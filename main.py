import os
import json
import re
import base64
import tempfile
import uuid
import asyncio
from collections import Counter
from datetime import datetime
from typing import Optional

import anthropic
import fitz  # PyMuPDF
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    raise RuntimeError("ANTHROPIC_API_KEY environment variable is required")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
sessions: dict[str, dict] = {}
jobs: dict[str, dict] = {}
_running_tasks: set = set()


class UpdateTransactionRequest(BaseModel):
    include: bool
    reason: Optional[str] = None


def extract_balanced(text: str, open_ch: str, close_ch: str, start: int) -> tuple[int, int]:
    depth = 0
    in_string = False
    escape_next = False
    for i, ch in enumerate(text[start:], start):
        if escape_next:
            escape_next = False
            continue
        if ch == "\\" and in_string:
            escape_next = True
            continue
        if ch == '"' and not escape_next:
            in_string = not in_string
            continue
        if in_string:
            continue
        if ch == open_ch:
            depth += 1
        elif ch == close_ch:
            depth -= 1
            if depth == 0:
                return start, i
    return start, -1


def extract_json_transactions(text: str) -> list:
    text = re.sub(r"^```[a-z]*\n?", "", text.strip(), flags=re.MULTILINE)
    text = re.sub(r"\n?```$", "", text.strip(), flags=re.MULTILINE)
    text = text.strip()
    try:
        result = json.loads(text)
        if isinstance(result, list):
            return result
        if isinstance(result, dict) and "transactions" in result:
            return result["transactions"]
    except json.JSONDecodeError:
        pass
    obj_start = text.find("{")
    if obj_start != -1:
        _, obj_end = extract_balanced(text, "{", "}", obj_start)
        if obj_end != -1:
            try:
                obj = json.loads(text[obj_start: obj_end + 1])
                if isinstance(obj, dict) and "transactions" in obj:
                    return obj["transactions"]
            except json.JSONDecodeError:
                pass
    arr_start = text.find("[")
    if arr_start == -1:
        raise ValueError("No JSON array or object found in response")
    _, arr_end = extract_balanced(text, "[", "]", arr_start)
    if arr_end != -1:
        try:
            return json.loads(text[arr_start: arr_end + 1])
        except json.JSONDecodeError:
            pass
    partial = text[arr_start:]
    last_complete = max(partial.rfind("},"), partial.rfind("}\n"))
    if last_complete != -1:
        partial = partial[: last_complete + 1] + "]"
        try:
            return json.loads(partial)
        except json.JSONDecodeError:
            pass
    raise ValueError(f"Could not parse JSON from response. Length={len(text)}")


def normalize_transaction(tx: dict) -> dict:
    if "suggested_include" in tx and "include" not in tx:
        tx["include"] = tx["suggested_include"].strip().upper() == "Y"
    try:
        tx["amount"] = float(tx.get("amount", 0))
    except (ValueError, TypeError):
        tx["amount"] = 0.0
    if tx.get("type") == "expense" and tx["amount"] > 0:
        tx["amount"] = -abs(tx["amount"])
    elif tx.get("type") == "income":
        tx["amount"] = abs(tx["amount"])
    tx.setdefault("date", "")
    tx.setdefault("description", "")
    tx.setdefault("category", "Other")
    tx.setdefault("reason", "")
    tx.setdefault("month", "")
    return tx


def deduplicate_categories(transactions: list) -> list:
    """Post-processing pass: normalize inconsistent category/include for same description+type."""
    key_cats: dict = {}
    key_incs: dict = {}
    for tx in transactions:
        key = (tx.get("description", "").strip().lower(), tx.get("type", ""))
        if key not in key_cats:
            key_cats[key] = Counter()
            key_incs[key] = Counter()
        key_cats[key][tx.get("category", "")] += 1
        key_incs[key][tx.get("include", True)] += 1
    canonical_cat = {k: v.most_common(1)[0][0] for k, v in key_cats.items()}
    canonical_inc = {k: v.most_common(1)[0][0] for k, v in key_incs.items()}
    for tx in transactions:
        key = (tx.get("description", "").strip().lower(), tx.get("type", ""))
        tx["category"] = canonical_cat[key]
        tx["include"] = canonical_inc[key]
    return transactions


PROMPT = """MORTGAGE UNDERWRITING — BANK STATEMENT INCOME & EXPENSE ANALYSIS
You are a mortgage underwriting analyst. Analyze the uploaded bank statement and return a structured JSON object with every transaction. Do not produce the Excel directly. Return clean structured data for human review first.
EXTRACTION: Extract every single transaction without exception — do not skip, group, summarize, or combine any lines. For each transaction return: date (YYYY-MM-DD), description (exact text from statement), amount (positive number), type (income or expense), month (e.g. Apr-25), category, suggested_include (Y or N), reason (plain English explanation). Do not include opening balance, closing balance, or running balance lines as transactions. Only extract actual debit and credit transactions.
LARGE STATEMENTS: If the statement has hundreds of transactions, you must still extract every single one. Do not stop early. Do not summarize groups of transactions. Every line item must appear as its own row.
CATEGORIES: Do not use a fixed list. Read the statement and assign the most specific label from the payee name or transfer type exactly as it appears. Income examples: E-Transfer In, ATM Deposit, Mobile Deposit, Cheque Deposit, Wire Transfer, Cash Deposit, Lendcare Loan, Internal Transfer In, NSF Reversal, Govt Rebate, or exact sender name. Expense examples: exact payee name such as Enbridge Gas or Cogeco, or type such as E-Transfer Out, Cash Withdrawal, Bank Charges, Cheque, CC Transfer, Transfer Out. Use Other Income or Other Expense only if truly unclear and explain in reason field. SPECIAL CATEGORY RULE: MSP fees, merchant services fees, point of sale fees, and terminal fees — always categorize these as Merchant Services Fee and always set suggested_include to Y. Never categorize these as Bank Charges. Any description starting with D/L or containing INT followed by a number string is a direct debit or interest payment — categorize as the exact payee name if identifiable or as Direct Debit Payment if unclear, and set suggested_include Y.
CONSISTENCY RULE: This rule is absolute and non-negotiable. Identical or near-identical transaction descriptions must always receive identical category and identical suggested_include values across the entire statement and across all pages. If the same payee or description pattern appears 50 times, all 50 must have identical category and suggested_include. Never make a different judgment on the same type of transaction on different dates. Before finalizing output, mentally scan all transactions and ensure no duplicates have conflicting categories or include values.
INCLUDE / EXCLUDE LOGIC: Every transaction must appear in the output regardless of Y or N. Y or N only controls whether it counts in summary totals. The reviewer can flip any Y or N in the review screen.
INCOME — Auto-set N, flag as excluded: Internal transfers from own accounts — any transfer that appears to come from the same account holder at the same or another bank, look for keywords like TFR-FR, transfer from, own account. NSF reversals and re-credits after returned payments. Government rebates such as HST rebate, GST credit, Ontario Trillium Benefit. Any wire transfer regardless of amount — set N and reason: Review: wire transfer — verify source and whether this is qualifying income. Any deposit that is a reversal or return of a prior outgoing transaction — including e-transfers that were sent out and came back, if an e-transfer was sent out and then deposited back it is NOT income, set N and reason: Returned outgoing e-transfer — not new income. Any deposit with no clear identifiable source — reason must say: Review: unusual deposit — source unclear. Lendcare loan deposits — always set N with reason: Lendcare loan — verify if this should be counted as qualifying income. Any single deposit that is significantly larger than the average deposit amount in the statement — set N and reason: Review: unusually large deposit — verify source. Any deposit that appears to be a loan, line of credit draw, financing, or borrowing from any lender or financial institution — set N with reason: Review: possible loan deposit — verify if this is qualifying income. Look for keywords like loan, LOC, credit, financing, advance, lendcare, or any known lender name in the description. Any transaction description that appears to be an outgoing or debit transaction must never be classified as income regardless of how it appears. If a transaction is clearly money leaving the account it is always an expense. Double check every income transaction — if there is any doubt whether it is truly incoming money, set it as expense.
INCOME — Auto-set Y: E-transfers from clearly external senders. ATM deposits, mobile deposits, cash deposits, cheque deposits from third parties. Regular recurring deposits that appear employment or business related. When uncertain default Y and note uncertainty in reason.
EXPENSE — Auto-set N, always excluded: ALL bank fees without exception — monthly plan fees, NSF fees, overdraft interest, e-transfer send fees, wire fees, service charges — always N, never override. Credit card payments and CC transfers. Transfers to own accounts at same or other institutions. Credit card bill payments of any kind — look for keywords like MC, VISA, AMEX, MASTERCARD, CAN TIRE MC, TD VISA, CIBC VISA, RBC VISA, SCOTIA VISA, BMO MC, or any description that combines a card issuer name with alphanumeric characters. Always N never override.
EXPENSE — Auto-set Y: Insurance payments. Utilities such as gas, hydro, internet, cable. Loan and mortgage payments. E-transfers out to clearly external payees. Cheques payable to third party individuals or businesses — always Y unless there is a clear reason not to include. Rent payments. Any regular identifiable recurring expense. When uncertain default Y and note uncertainty in reason.
FLAGGING UNUSUAL ITEMS: For any transaction that seems unusual, one-time, very large, or unclear, always set reason to start with Review: followed by your concern. This applies to both income and expenses.
OUTPUT FORMAT: Return only valid JSON with no prose and no markdown fences. Structure: statement_period, account_holder, bank, transactions array where each transaction has date, description, amount, type, month, category, suggested_include, reason."""


HTML = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Bank Statement Analyzer</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <style>
    @keyframes spin { to { transform: rotate(360deg); } }
    .spin { animation: spin 1s linear infinite; }
    .clamp2 { overflow: hidden; display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; }
    body { font-family: system-ui, -apple-system, sans-serif; }
  </style>
</head>
<body class="m-0 p-0 bg-gray-50">
<div id="app"></div>
<script>
(function() {

  var state = {
    page: 'upload',
    files: [],
    isDragging: false,
    isAnalyzing: false,
    analyzeError: null,
    jobId: null,
    progressInfo: {},
    sessionId: null,
    transactions: [],
    filter: 'all',
    search: '',
    updating: {},
    downloading: false,
    downloaded: false
  };

  var pollTimer = null;

  function h(str) {
    return String(str)
      .replace(/&/g,'&amp;').replace(/</g,'&lt;')
      .replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  }

  function fmt(n) {
    return new Intl.NumberFormat('en-US',{style:'currency',currency:'USD'}).format(n);
  }

  function getFiltered() {
    var r = state.transactions.slice();
    if (state.filter === 'income') r = r.filter(function(t){return t.type==='income';});
    else if (state.filter === 'expense') r = r.filter(function(t){return t.type==='expense';});
    else if (state.filter === 'included') r = r.filter(function(t){return t.include;});
    else if (state.filter === 'excluded') r = r.filter(function(t){return !t.include;});
    if (state.search.trim()) {
      var q = state.search.toLowerCase();
      r = r.filter(function(t){
        return t.description.toLowerCase().indexOf(q)>=0 ||
               t.category.toLowerCase().indexOf(q)>=0 ||
               t.date.indexOf(q)>=0 ||
               t.reason.toLowerCase().indexOf(q)>=0;
      });
    }
    return r;
  }

  function getTotals() {
    var inc = state.transactions.filter(function(t){return t.include;});
    var income = inc.filter(function(t){return t.type==='income';}).reduce(function(s,t){return s+t.amount;},0);
    var expenses = inc.filter(function(t){return t.type==='expense';}).reduce(function(s,t){return s+Math.abs(t.amount);},0);
    return {income:income, expenses:expenses, net:income-expenses,
            count:state.transactions.length, includedCount:inc.length};
  }

  function setState(updates) {
    Object.assign(state, updates);
    render();
  }

  function render() {
    var activeId = document.activeElement ? document.activeElement.id : null;
    var scrollY = window.scrollY;
    var app = document.getElementById('app');
    if (state.page === 'upload') app.innerHTML = renderUpload();
    else if (state.page === 'progress') app.innerHTML = renderProgress();
    else if (state.page === 'review') app.innerHTML = renderReview();
    else app.innerHTML = renderDownload();
    attachListeners();
    if (activeId) {
      var el = document.getElementById(activeId);
      if (el) {
        el.focus();
        if ((el.tagName==='INPUT'||el.tagName==='TEXTAREA') && el.type!=='file') {
          try { el.setSelectionRange(el.value.length, el.value.length); } catch(e){}
        }
      }
    }
    window.scrollTo(0, scrollY);
  }

  /* ============================================================
     UPLOAD PAGE
  ============================================================ */
  function renderUpload() {
    var fileRows = state.files.map(function(f,i){
      return '<div class="flex items-center justify-between bg-white border border-gray-200 rounded-xl px-4 py-3 shadow-sm">' +
        '<div class="flex items-center gap-3 min-w-0">' +
          '<div class="w-8 h-8 rounded-lg bg-red-50 flex items-center justify-center flex-shrink-0 text-red-500 text-xs font-bold">PDF</div>' +
          '<div class="min-w-0">' +
            '<p class="text-sm font-medium text-gray-800 truncate">' + h(f.name) + '</p>' +
            '<p class="text-xs text-gray-400">' + (f.size/1024).toFixed(1) + ' KB</p>' +
          '</div>' +
        '</div>' +
        '<button data-remove="' + i + '" class="ml-3 p-1.5 rounded-lg hover:bg-red-50 text-gray-300 hover:text-red-500 transition-colors text-lg leading-none">&times;</button>' +
      '</div>';
    }).join('');

    var dz = state.isDragging
      ? 'border-blue-500 bg-blue-50 scale-105'
      : 'border-gray-300 bg-white hover:border-blue-400';
    var btnDis = state.isAnalyzing || state.files.length === 0;
    var btnCls = btnDis
      ? 'bg-gray-200 text-gray-400 cursor-not-allowed'
      : 'bg-blue-600 text-white hover:bg-blue-700 shadow-md hover:shadow-lg active:scale-95';

    return '<div class="min-h-screen flex flex-col items-center justify-center p-6 bg-gray-50">' +
      '<div class="w-full max-w-2xl">' +
        '<div class="mb-8 text-center">' +
          '<div class="inline-flex items-center justify-center w-14 h-14 rounded-xl bg-blue-600 mb-4">' +
            '<svg class="w-7 h-7 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>' +
          '</div>' +
          '<h1 class="text-3xl font-bold text-gray-900 tracking-tight">Bank Statement Analyzer</h1>' +
          '<p class="mt-2 text-gray-500">Upload PDF bank statements for mortgage underwriting analysis</p>' +
        '</div>' +
        '<div id="drop-zone" class="border-2 border-dashed rounded-2xl p-10 text-center cursor-pointer transition-all duration-200 ' + dz + '">' +
          '<input id="file-input" type="file" accept=".pdf,application/pdf" multiple class="hidden">' +
          '<div class="flex flex-col items-center gap-3">' +
            '<div class="p-4 rounded-full ' + (state.isDragging?'bg-blue-100':'bg-gray-100') + '">' +
              '<svg class="w-8 h-8 ' + (state.isDragging?'text-blue-600':'text-gray-400') + '" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>' +
            '</div>' +
            '<div>' +
              '<p class="text-base font-semibold text-gray-700">' + (state.isDragging ? 'Drop your PDFs here' : 'Drag &amp; drop PDFs here') + '</p>' +
              '<p class="text-sm text-gray-400 mt-1">or <span class="text-blue-600 font-medium">browse files</span> &mdash; multiple files supported</p>' +
            '</div>' +
          '</div>' +
        '</div>' +
        (state.files.length > 0 ? '<div class="mt-4 space-y-2">' + fileRows + '</div>' : '') +
        (state.analyzeError ? '<div class="mt-4 flex items-center gap-3 bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-red-600"><span class="text-lg">&#9888;</span><p class="text-sm">' + h(state.analyzeError) + '</p></div>' : '') +
        '<button id="analyze-btn" ' + (btnDis?'disabled':'') + ' class="mt-6 w-full py-4 rounded-xl font-semibold text-base flex items-center justify-center gap-3 transition-all duration-200 ' + btnCls + '">' +
          (state.isAnalyzing
            ? '<svg class="spin w-5 h-5" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Uploading files…'
            : '<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>Analyze Statements') +
        '</button>' +
      '</div>' +
    '</div>';
  }

  /* ============================================================
     PROGRESS PAGE
  ============================================================ */
  function renderProgress() {
    var info = state.progressInfo || {};
    var isError = info.status === 'error';
    var pct = (info.total_pages > 0) ? Math.round(info.pages_done / info.total_pages * 100) : 0;

    return '<div class="min-h-screen flex flex-col items-center justify-center p-6 bg-gray-50">' +
      '<div class="w-full max-w-xl">' +
        '<div class="mb-8 text-center">' +
          '<div class="inline-flex items-center justify-center w-14 h-14 rounded-xl ' + (isError ? 'bg-red-500' : 'bg-blue-600') + ' mb-4">' +
            (isError
              ? '<svg class="w-7 h-7 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-3L13.732 4c-.77-1.333-2.694-1.333-3.464 0L3.34 16c-.77 1.333.192 3 1.732 3z"/></svg>'
              : '<svg class="spin w-7 h-7 text-white" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>') +
          '</div>' +
          '<h1 class="text-3xl font-bold text-gray-900 tracking-tight">' + (isError ? 'Analysis Failed' : 'Analyzing Statements') + '</h1>' +
          '<p class="mt-2 text-gray-500">' + (isError ? 'An error occurred during processing.' : 'Processing your bank statements with Claude AI…') + '</p>' +
        '</div>' +
        '<div class="bg-white rounded-2xl border border-gray-200 shadow-sm p-6 mb-6">' +
          '<p class="text-sm font-semibold text-gray-700 mb-4">' + h(info.current_file || 'Starting…') + '</p>' +
          (info.total_pages > 0
            ? '<div class="mb-1"><div class="flex justify-between text-xs text-gray-400 mb-2"><span>Pages processed</span><span>' + info.pages_done + ' / ' + info.total_pages + '</span></div>' +
              '<div class="w-full bg-gray-200 rounded-full h-2"><div class="bg-blue-600 h-2 rounded-full transition-all duration-500" style="width:' + pct + '%"></div></div></div>'
            : '<div class="flex items-center gap-2 text-gray-400 text-sm"><div class="w-2 h-2 rounded-full bg-blue-500 animate-pulse"></div><span>Preparing files…</span></div>') +
          (isError ? '<p class="mt-4 text-sm text-red-600 bg-red-50 rounded-lg p-3">' + h(info.error || 'Unknown error') + '</p>' : '') +
        '</div>' +
        (isError
          ? '<button id="retry-btn" class="w-full py-4 rounded-xl font-semibold text-base bg-blue-600 text-white hover:bg-blue-700 shadow-md hover:shadow-lg transition-all">← Back to Upload</button>'
          : '<p class="text-center text-sm text-gray-400">Checking progress every 5 seconds — do not close this tab.</p>') +
      '</div>' +
    '</div>';
  }

  /* ============================================================
     REVIEW PAGE
  ============================================================ */
  function renderReview() {
    var totals = getTotals();
    var filtered = getFiltered();

    var filterBtns = ['all','income','expense','included','excluded'].map(function(v){
      var labels = {all:'All',income:'Income',expense:'Expenses',included:'Y Only',excluded:'N Only'};
      var active = state.filter === v;
      return '<button data-filter="' + v + '" class="px-4 py-1.5 rounded-lg text-sm font-medium transition-all ' +
        (active ? 'bg-blue-600 text-white shadow-sm' : 'bg-gray-100 text-gray-600 hover:bg-gray-200') + '">' + labels[v] + '</button>';
    }).join('');

    var rows = filtered.length === 0
      ? '<tr><td colspan="7" class="px-4 py-12 text-center text-gray-400">No transactions match your filters.</td></tr>'
      : filtered.map(function(tx){
          var upd = state.updating[tx.id];
          var rBg = tx.include ? 'bg-green-50 hover:bg-green-100' : 'bg-orange-50 hover:bg-orange-100';
          var bCls = upd
            ? 'opacity-50 cursor-wait bg-gray-300 text-white'
            : (tx.include ? 'bg-green-500 text-white hover:bg-green-600' : 'bg-orange-400 text-white hover:bg-orange-500');
          var aCls = tx.amount >= 0 ? 'text-green-700' : 'text-red-600';
          var typeCls = tx.type === 'income' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-600';
          return '<tr class="' + rBg + ' transition-colors">' +
            '<td class="px-4 py-3 font-mono text-xs text-gray-500 whitespace-nowrap">' + h(tx.date) + '</td>' +
            '<td class="px-4 py-3 text-gray-800 max-w-xs"><span class="clamp2">' + h(tx.description) + '</span></td>' +
            '<td class="px-4 py-3 text-right font-semibold tabular-nums whitespace-nowrap ' + aCls + '">' + (tx.amount>=0?'+':'') + fmt(tx.amount) + '</td>' +
            '<td class="px-4 py-3 text-center"><span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ' + typeCls + '">' + h(tx.type) + '</span></td>' +
            '<td class="px-4 py-3 text-gray-500 text-xs">' + h(tx.category) + '</td>' +
            '<td class="px-4 py-3 text-center"><button data-toggle="' + h(tx.id) + '" ' + (upd?'disabled':'') + ' class="inline-flex items-center justify-center w-12 h-7 rounded-md text-xs font-bold transition-all ' + bCls + '">' +
              (upd ? '<svg class="spin w-3 h-3" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>' : (tx.include?'Y':'N')) +
            '</button></td>' +
            '<td class="px-4 py-3 text-gray-400 text-xs max-w-xs"><span class="clamp2">' + h(tx.reason) + '</span></td>' +
          '</tr>';
        }).join('');

    function sc(icon, label, value, sub, color) {
      var c = {blue:'text-blue-600 bg-blue-50',green:'text-green-600 bg-green-50',orange:'text-orange-500 bg-orange-50',red:'text-red-500 bg-red-50'};
      return '<div class="flex items-center gap-3">' +
        '<div class="w-9 h-9 rounded-lg flex items-center justify-center flex-shrink-0 font-bold ' + c[color] + '">' + icon + '</div>' +
        '<div><p class="text-xs text-gray-400 font-medium">' + label + '</p>' +
        '<p class="text-sm font-bold text-gray-800">' + value + '</p>' +
        '<p class="text-xs text-gray-400">' + sub + '</p></div></div>';
    }

    return '<div class="min-h-screen bg-gray-50 flex flex-col">' +
      '<div class="bg-gray-900 text-white px-6 py-4">' +
        '<div class="max-w-7xl mx-auto">' +
          '<h1 class="text-xl font-bold tracking-tight">Bank Statement Review</h1>' +
          '<p class="text-gray-400 text-sm mt-0.5">Review and toggle transactions for mortgage underwriting</p>' +
        '</div>' +
      '</div>' +
      '<div class="bg-white border-b border-gray-200 px-6 py-4">' +
        '<div class="max-w-7xl mx-auto grid grid-cols-2 sm:grid-cols-4 gap-4">' +
          sc('#','Transactions',totals.includedCount+' / '+totals.count,'included','blue') +
          sc('&uarr;','Total Annual Income',fmt(totals.income),'included only','green') +
          sc('&darr;','Total Annual Expenses',fmt(totals.expenses),'included only','orange') +
          sc('$','Net Annual Income',fmt(totals.net),'income minus expenses',totals.net>=0?'green':'red') +
        '</div>' +
      '</div>' +
      '<div class="bg-white border-b border-gray-200 px-6 py-3">' +
        '<div class="max-w-7xl mx-auto flex flex-col sm:flex-row gap-3 items-start sm:items-center justify-between">' +
          '<div class="flex items-center gap-2 flex-wrap">' + filterBtns + '</div>' +
          '<input id="search-input" type="search" placeholder="Search transactions..." value="' + h(state.search) + '" ' +
            'class="pl-4 pr-4 py-2 text-sm border border-gray-200 rounded-lg bg-gray-50 focus:outline-none focus:ring-2 focus:ring-blue-300 focus:border-blue-500 w-60">' +
        '</div>' +
      '</div>' +
      '<div class="flex-1 px-6 py-4 overflow-auto">' +
        '<div class="max-w-7xl mx-auto">' +
          '<p class="text-sm text-gray-400 mb-3">Showing ' + filtered.length + ' of ' + state.transactions.length + ' transactions</p>' +
          '<div class="bg-white rounded-2xl border border-gray-200 overflow-hidden shadow-sm">' +
            '<div class="overflow-x-auto">' +
              '<table class="w-full text-sm">' +
                '<thead><tr class="bg-gray-900 text-gray-300">' +
                  '<th class="px-4 py-3 text-left font-semibold text-xs uppercase tracking-wider w-28">Date</th>' +
                  '<th class="px-4 py-3 text-left font-semibold text-xs uppercase tracking-wider">Description</th>' +
                  '<th class="px-4 py-3 text-right font-semibold text-xs uppercase tracking-wider w-28">Amount</th>' +
                  '<th class="px-4 py-3 text-center font-semibold text-xs uppercase tracking-wider w-24">Type</th>' +
                  '<th class="px-4 py-3 text-left font-semibold text-xs uppercase tracking-wider w-40">Category</th>' +
                  '<th class="px-4 py-3 text-center font-semibold text-xs uppercase tracking-wider w-20">Include</th>' +
                  '<th class="px-4 py-3 text-left font-semibold text-xs uppercase tracking-wider">Reason</th>' +
                '</tr></thead>' +
                '<tbody class="divide-y divide-gray-100">' + rows + '</tbody>' +
              '</table>' +
            '</div>' +
          '</div>' +
        '</div>' +
      '</div>' +
      '<div class="bg-white border-t border-gray-200 px-6 py-4">' +
        '<div class="max-w-7xl mx-auto flex justify-end">' +
          '<button id="to-download-btn" class="flex items-center gap-2 px-6 py-3 bg-green-500 text-white rounded-xl font-semibold hover:bg-green-600 transition-all shadow-md hover:shadow-lg">' +
            '<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>' +
            'Generate Excel Report' +
          '</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  /* ============================================================
     DOWNLOAD PAGE
  ============================================================ */
  function renderDownload() {
    var inc = state.transactions.filter(function(t){return t.include;});
    var totalIncome = inc.filter(function(t){return t.type==='income';}).reduce(function(s,t){return s+t.amount;},0);
    var totalExpenses = inc.filter(function(t){return t.type==='expense';}).reduce(function(s,t){return s+Math.abs(t.amount);},0);
    var net = totalIncome - totalExpenses;
    var netColor = net >= 0 ? 'text-green-700' : 'text-red-600';
    var dlBtnCls = (state.downloaded ? 'bg-green-500' : 'bg-blue-600 hover:bg-blue-700 hover:shadow-lg') +
      (state.downloading ? ' opacity-70 cursor-wait' : '');

    return '<div class="min-h-screen flex flex-col items-center justify-center p-6 bg-gray-50">' +
      '<div class="w-full max-w-xl">' +
        '<div class="text-center mb-8">' +
          '<div class="inline-flex items-center justify-center w-16 h-16 rounded-2xl bg-green-500 mb-4 shadow-lg">' +
            '<svg class="w-8 h-8 text-white" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 17v-2m3 2v-4m3 4v-6m2 10H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>' +
          '</div>' +
          '<h1 class="text-3xl font-bold text-gray-900 tracking-tight">Analysis Complete</h1>' +
          '<p class="mt-2 text-gray-500">' + inc.length + ' of ' + state.transactions.length + ' transactions included in the report</p>' +
        '</div>' +
        '<div class="bg-white rounded-2xl border border-gray-200 shadow-sm overflow-hidden mb-6">' +
          '<div class="bg-gray-900 px-6 py-4"><h2 class="text-sm font-semibold text-gray-300 uppercase tracking-wider">Underwriting Summary</h2></div>' +
          '<div class="divide-y divide-gray-100">' +
            '<div class="flex items-center justify-between px-6 py-4">' +
              '<div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl bg-green-100 text-green-600 flex items-center justify-center font-bold">&uarr;</div><p class="text-sm font-medium text-gray-700">Total Annual Income</p></div>' +
              '<span class="text-xl font-bold text-green-600">' + fmt(totalIncome) + '</span>' +
            '</div>' +
            '<div class="flex items-center justify-between px-6 py-4">' +
              '<div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl bg-orange-100 text-orange-500 flex items-center justify-center font-bold">&darr;</div><p class="text-sm font-medium text-gray-700">Total Annual Expenses</p></div>' +
              '<span class="text-xl font-bold text-orange-500">' + fmt(totalExpenses) + '</span>' +
            '</div>' +
            '<div class="flex items-center justify-between px-6 py-4 bg-gray-50">' +
              '<div class="flex items-center gap-3">' +
                '<div class="w-10 h-10 rounded-xl flex items-center justify-center font-bold ' + (net>=0?'bg-green-100 text-green-700':'bg-red-100 text-red-600') + '">$</div>' +
                '<div><p class="text-xs text-gray-400 font-medium uppercase tracking-wide">Net Annual Income</p><p class="text-xs text-gray-400">Income minus Expenses</p></div>' +
              '</div>' +
              '<span class="text-2xl font-bold ' + netColor + '">' + fmt(net) + '</span>' +
            '</div>' +
          '</div>' +
        '</div>' +
        '<div class="bg-white rounded-2xl border border-gray-200 shadow-sm p-6 mb-6">' +
          '<h3 class="text-sm font-semibold text-gray-700 mb-3">What&#39;s in the Excel report</h3>' +
          '<ul class="space-y-2 text-sm text-gray-500">' +
            '<li>&#10003; Sheet 1: Summary — monthly income &amp; expense tables, annual totals, key metrics</li>' +
            '<li>&#10003; Sheet 2: Income Breakdown — every deposit with Y/N colour coding</li>' +
            '<li>&#10003; Sheet 3: Expense Breakdown — every withdrawal with Y/N colour coding</li>' +
            '<li>&#10003; AI-generated category and reason for every transaction</li>' +
          '</ul>' +
        '</div>' +
        '<button id="download-btn" ' + (state.downloading?'disabled':'') + ' class="w-full py-4 rounded-xl font-semibold text-base flex items-center justify-center gap-3 transition-all shadow-md mb-4 text-white ' + dlBtnCls + '">' +
          (state.downloading
            ? '<svg class="spin w-5 h-5" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Generating Excel file...'
            : state.downloaded
              ? '&#10003; Downloaded! Click to download again'
              : '<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"/></svg>Download Excel Report (.xlsx)') +
        '</button>' +
        '<div class="flex gap-3">' +
          '<button id="back-btn" class="flex-1 py-3 rounded-xl font-semibold text-sm flex items-center justify-center gap-2 border border-gray-200 bg-white text-gray-600 hover:bg-gray-50 transition-all">&larr; Back to Review</button>' +
          '<button id="reset-btn" class="flex-1 py-3 rounded-xl font-semibold text-sm flex items-center justify-center gap-2 bg-gray-100 text-gray-600 hover:bg-gray-200 transition-all">Analyze New Statements</button>' +
        '</div>' +
      '</div>' +
    '</div>';
  }

  /* ============================================================
     EVENT LISTENERS
  ============================================================ */
  function attachListeners() {
    if (state.page === 'upload') attachUpload();
    else if (state.page === 'progress') attachProgress();
    else if (state.page === 'review') attachReview();
    else attachDownload();
  }

  function attachUpload() {
    var dz = document.getElementById('drop-zone');
    var fi = document.getElementById('file-input');
    var ab = document.getElementById('analyze-btn');
    if (dz) {
      dz.addEventListener('click', function(){ if(!state.isAnalyzing) fi.click(); });
      dz.addEventListener('dragover', function(e){ e.preventDefault(); if(!state.isDragging) setState({isDragging:true}); });
      dz.addEventListener('dragleave', function(){ if(state.isDragging) setState({isDragging:false}); });
      dz.addEventListener('drop', function(e){ e.preventDefault(); setState({isDragging:false}); addFiles(e.dataTransfer.files); });
    }
    if (fi) fi.addEventListener('change', function(e){ addFiles(e.target.files); });
    document.querySelectorAll('[data-remove]').forEach(function(btn){
      btn.addEventListener('click', function(e){
        e.stopPropagation();
        var idx = parseInt(btn.getAttribute('data-remove'));
        setState({files: state.files.filter(function(_,i){return i!==idx;})});
      });
    });
    if (ab) ab.addEventListener('click', handleAnalyze);
  }

  function attachProgress() {
    var rb = document.getElementById('retry-btn');
    if (rb) rb.addEventListener('click', function(){
      stopPolling();
      setState({page:'upload', jobId:null, progressInfo:{}, analyzeError:null});
    });
  }

  function attachReview() {
    document.querySelectorAll('[data-filter]').forEach(function(btn){
      btn.addEventListener('click', function(){ setState({filter: btn.getAttribute('data-filter')}); });
    });
    var si = document.getElementById('search-input');
    if (si) si.addEventListener('input', function(e){ setState({search: e.target.value}); });
    document.querySelectorAll('[data-toggle]').forEach(function(btn){
      btn.addEventListener('click', function(){ handleToggle(btn.getAttribute('data-toggle')); });
    });
    var td = document.getElementById('to-download-btn');
    if (td) td.addEventListener('click', function(){ setState({page:'download', downloaded:false}); });
  }

  function attachDownload() {
    var db = document.getElementById('download-btn');
    if (db) db.addEventListener('click', handleDownload);
    var bb = document.getElementById('back-btn');
    if (bb) bb.addEventListener('click', function(){ setState({page:'review'}); });
    var rb = document.getElementById('reset-btn');
    if (rb) rb.addEventListener('click', function(){
      stopPolling();
      setState({page:'upload',files:[],sessionId:null,transactions:[],
                filter:'all',search:'',analyzeError:null,downloaded:false,
                jobId:null,progressInfo:{}});
    });
  }

  /* ============================================================
     ACTIONS
  ============================================================ */
  function addFiles(list) {
    if (!list) return;
    var pdfs = Array.from(list).filter(function(f){ return f.type==='application/pdf'||f.name.toLowerCase().endsWith('.pdf'); });
    var existing = {};
    state.files.forEach(function(f){ existing[f.name+f.size]=1; });
    var fresh = pdfs.filter(function(f){ return !existing[f.name+f.size]; });
    setState({files: state.files.concat(fresh), analyzeError:null});
  }

  async function handleAnalyze() {
    if (state.files.length===0) { setState({analyzeError:'Please add at least one PDF file.'}); return; }
    setState({isAnalyzing:true, analyzeError:null});
    try {
      var fd = new FormData();
      state.files.forEach(function(f){ fd.append('files',f); });
      var res = await fetch('/analyze',{method:'POST',body:fd});
      if (!res.ok) {
        var err = await res.json().catch(function(){return {detail:'Unknown error'};});
        throw new Error(err.detail||'Error '+res.status);
      }
      var data = await res.json();
      setState({
        isAnalyzing: false,
        jobId: data.job_id,
        page: 'progress',
        progressInfo: {status:'processing', current_file:'Starting…', pages_done:0, total_pages:0}
      });
      startPolling(data.job_id);
    } catch(e) {
      setState({isAnalyzing:false, analyzeError:e.message||'Failed to analyze. Please try again.'});
    }
  }

  function startPolling(jobId) {
    stopPolling();
    setTimeout(function(){ pollJob(jobId); }, 2000);
    pollTimer = setInterval(function(){ pollJob(jobId); }, 5000);
  }

  function stopPolling() {
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
  }

  async function pollJob(jobId) {
    if (state.page !== 'progress') { stopPolling(); return; }
    try {
      var res = await fetch('/jobs/'+jobId+'/status');
      if (!res.ok) return;
      var data = await res.json();
      if (data.status === 'complete') {
        stopPolling();
        var txRes = await fetch('/sessions/'+data.session_id+'/transactions');
        if (!txRes.ok) {
          setState({progressInfo:{status:'error', error:'Failed to load transactions. Please try again.', current_file:'', pages_done:0, total_pages:0}});
          return;
        }
        var transactions = await txRes.json();
        setState({
          page: 'review',
          sessionId: data.session_id,
          transactions: transactions,
          progressInfo: data,
          filter: 'all',
          search: ''
        });
      } else if (data.status === 'error') {
        stopPolling();
        setState({progressInfo: data});
      } else {
        setState({progressInfo: data});
      }
    } catch(e) {
      // silent — will retry on next interval
    }
  }

  async function handleToggle(txId) {
    var tx = state.transactions.find(function(t){return t.id===txId;});
    if (!tx || state.updating[txId]) return;
    var newInclude = !tx.include;

    // Optimistic update: flip include immediately so totals update in real time
    var optimisticTxs = state.transactions.map(function(t){
      return t.id===txId ? Object.assign({},t,{include:newInclude}) : t;
    });
    var upd = Object.assign({},state.updating); upd[txId]=true;
    setState({transactions:optimisticTxs, updating:upd});

    try {
      var res = await fetch('/sessions/'+state.sessionId+'/transactions/'+txId,{
        method:'PUT', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({include:newInclude})
      });
      if (!res.ok) throw new Error('Failed');
      var updated = await res.json();
      var txs = state.transactions.map(function(t){return t.id===updated.id?updated:t;});
      var done = Object.assign({},state.updating); delete done[txId];
      setState({transactions:txs, updating:done});
    } catch(e) {
      // Revert optimistic update on failure
      var revertTxs = state.transactions.map(function(t){
        return t.id===txId ? Object.assign({},t,{include:!newInclude}) : t;
      });
      var done2 = Object.assign({},state.updating); delete done2[txId];
      setState({transactions:revertTxs, updating:done2});
    }
  }

  async function handleDownload() {
    setState({downloading:true});
    try {
      var res = await fetch('/sessions/'+state.sessionId+'/export');
      if (!res.ok) throw new Error('Failed to generate export');
      var blob = await res.blob();
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href=url; a.download='bank_statement_analysis_'+new Date().toISOString().slice(0,10)+'.xlsx';
      document.body.appendChild(a); a.click(); document.body.removeChild(a);
      URL.revokeObjectURL(url);
      setState({downloading:false, downloaded:true});
    } catch(e) {
      setState({downloading:false});
      alert('Failed to download. Please try again.');
    }
  }

  render();
})();
</script>
</body>
</html>"""


app = FastAPI(title="Mortgage Bank Statement Analyzer")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])


@app.get("/")
def index():
    return HTMLResponse(HTML)


@app.get("/health")
def health():
    return {"status": "ok"}


def _extract_pdf_text(pdf_bytes: bytes, filename: str) -> str:
    """Stage 1: Extract all text from every page using PyMuPDF, concatenated into one string."""
    pages_text: list[str] = []
    try:
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        for page_num, page in enumerate(doc, 1):
            text = page.get_text()
            if text.strip():
                pages_text.append(f"--- Page {page_num} ---\n{text}")
        doc.close()
    except Exception as e:
        print(f"[WARN] fitz extraction error for {filename}: {e}")
    full_text = "\n\n".join(pages_text)
    print(f"[INFO] {filename}: extracted {len(full_text)} chars from {len(pages_text)} pages")
    return full_text


def _identify_transaction_lines(full_text: str) -> list[str]:
    lines = full_text.split('\n')
    kept = []
    skip_keywords = ['opening balance', 'closing balance', 'balance forward',
                     'statement period', 'account number', 'page ', 'total ']
    for line in lines:
        line = line.strip()
        if not line:
            continue
        if len(line) < 3:
            continue
        if line.replace('-', '').replace('=', '').replace('*', '').strip() == '':
            continue
        if any(kw in line.lower() for kw in skip_keywords):
            continue
        kept.append(line)
    return kept


async def _process_job(job_id: str, file_data: list[tuple[str, bytes]]):
    """Background task: fitz extraction → transaction line identification → batched Claude categorization."""
    BATCH_SIZE = 50
    all_transactions: list = []
    total_files = len(file_data)

    try:
        for file_idx, (filename, pdf_bytes) in enumerate(file_data):
            jobs[job_id].update({
                "current_file": f"Processing file {file_idx + 1} of {total_files}: {filename} (extracting text…)",
                "pages_done": 0,
                "total_pages": 0,
            })

            # Stage 1 — fitz full-text extraction
            full_text = await asyncio.to_thread(_extract_pdf_text, pdf_bytes, filename)

            if not full_text.strip():
                print(f"[WARN] {filename}: no text extracted, skipping")
                continue

            # Identify transaction lines via date + amount regex
            tx_lines = _identify_transaction_lines(full_text)
            print(f"[INFO] {filename}: {len(tx_lines)} transaction lines identified")

            if not tx_lines:
                print(f"[WARN] {filename}: no transaction lines found, skipping")
                continue

            total_batches = (len(tx_lines) + BATCH_SIZE - 1) // BATCH_SIZE
            jobs[job_id]["total_pages"] = total_batches

            file_transactions: list = []
            category_summary: dict[str, str] = {}  # {category: "Y"/"N"} for cross-batch consistency

            # Stage 2 — sequential non-overlapping batches of 50 lines each
            for batch_num, i in enumerate(range(0, len(tx_lines), BATCH_SIZE), 1):
                batch = tx_lines[i:i + BATCH_SIZE]  # lines i..i+49, never overlapping
                jobs[job_id]["current_file"] = (
                    f"Processing batch {batch_num} of {total_batches} for {filename}"
                )

                system_prompt = PROMPT
                if batch_num > 1 and category_summary:
                    system_prompt += (
                        f"\n\nPreviously seen categories and their include decisions: {category_summary}"
                    )

                prompt_text = system_prompt + "\n\n" + "\n".join(batch)

                def _sync(pt=prompt_text) -> str:
                    with client.messages.stream(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=32000,
                        messages=[{"role": "user", "content": pt}],
                    ) as stream:
                        return stream.get_final_text().strip()

                try:
                    response_text = await asyncio.to_thread(_sync)
                    raw = extract_json_transactions(response_text)
                    for tx in raw:
                        tx = normalize_transaction(tx)
                        tx["id"] = str(uuid.uuid4())
                        file_transactions.append(tx)
                        cat = tx.get("category", "")
                        if cat and cat not in category_summary:
                            category_summary[cat] = "Y" if tx.get("include", True) else "N"
                    print(f"[INFO] {filename} batch {batch_num}/{total_batches}: {len(raw)} transactions")
                except Exception as e:
                    print(f"[WARN] Claude error for {filename} batch {batch_num}: {e}")

                jobs[job_id]["pages_done"] = batch_num

            all_transactions.extend(file_transactions)
            print(f"[INFO] {filename}: complete — {len(file_transactions)} transactions")

        if all_transactions:
            all_transactions = deduplicate_categories(all_transactions)

        session_id = str(uuid.uuid4())
        sessions[session_id] = {"transactions": all_transactions}
        jobs[job_id].update({
            "status": "complete",
            "session_id": session_id,
            "transaction_count": len(all_transactions),
            "current_file": f"Complete — {len(all_transactions)} transactions extracted from {total_files} file(s)",
        })
        print(f"[INFO] Job {job_id} complete: {len(all_transactions)} transactions, session {session_id}")

    except Exception as e:
        jobs[job_id] = {
            "status": "error",
            "error": str(e),
            "current_file": "",
            "pages_done": 0,
            "total_pages": 0,
        }
        print(f"[ERROR] Job {job_id} failed: {e}")


@app.post("/analyze")
async def analyze_statements(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    for file in files:
        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"{file.filename} is not a PDF")

    file_data = [(file.filename, await file.read()) for file in files]
    job_id = str(uuid.uuid4())
    jobs[job_id] = {
        "status": "processing",
        "current_file": f"Starting — 0 of {len(file_data)} files processed",
        "pages_done": 0,
        "total_pages": 0,
    }

    task = asyncio.create_task(_process_job(job_id, file_data))
    _running_tasks.add(task)
    task.add_done_callback(_running_tasks.discard)

    return {"job_id": job_id}


@app.get("/jobs/{job_id}/status")
def get_job_status(job_id: str):
    if job_id not in jobs:
        raise HTTPException(status_code=404, detail="Job not found")
    job = jobs[job_id]
    response = {
        "status": job["status"],
        "current_file": job.get("current_file", ""),
        "pages_done": job.get("pages_done", 0),
        "total_pages": job.get("total_pages", 0),
    }
    if job["status"] == "complete":
        response["session_id"] = job["session_id"]
        response["transaction_count"] = job["transaction_count"]
    elif job["status"] == "error":
        response["error"] = job.get("error", "Unknown error")
    return response


@app.get("/sessions/{session_id}/transactions")
def get_session_transactions(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    return sessions[session_id]["transactions"]


@app.put("/sessions/{session_id}/transactions/{transaction_id}")
def update_transaction(session_id: str, transaction_id: str, body: UpdateTransactionRequest):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")
    for tx in sessions[session_id]["transactions"]:
        if tx["id"] == transaction_id:
            tx["include"] = body.include
            if body.reason is not None:
                tx["reason"] = body.reason
            return tx
    raise HTTPException(status_code=404, detail="Transaction not found")


@app.get("/sessions/{session_id}/export")
def export_excel(session_id: str):
    if session_id not in sessions:
        raise HTTPException(status_code=404, detail="Session not found")

    transactions = sessions[session_id]["transactions"]
    income_txs = [tx for tx in transactions if tx.get("type") == "income"]
    expense_txs = [tx for tx in transactions if tx.get("type") == "expense"]

    def parse_month(m):
        try:
            return datetime.strptime(m, "%b-%y")
        except Exception:
            return datetime.min

    all_months = sorted(set(tx.get("month", "") for tx in transactions if tx.get("month")), key=parse_month)
    income_months = sorted(set(tx.get("month", "") for tx in income_txs if tx.get("month")), key=parse_month)
    expense_months = sorted(set(tx.get("month", "") for tx in expense_txs if tx.get("month")), key=parse_month)
    income_cats = sorted(set(tx.get("category", "") for tx in income_txs if tx.get("category")))
    expense_cats = sorted(set(tx.get("category", "") for tx in expense_txs if tx.get("category")))

    n_inc = len(income_months) or 1
    n_exp = len(expense_months) or 1
    n_all = len(all_months) or 1

    # ── Styles ────────────────────────────────────────────────────────────────
    green_row  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    orange_row = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
    inc_y_fill = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    inc_n_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    hdr_fill   = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    sec_fill   = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    sub_fill   = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    tot_fill   = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    met_fill   = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")
    net_pos    = PatternFill(start_color="E8F8E8", end_color="E8F8E8", fill_type="solid")

    hdr_font  = Font(name="Arial", color="FFFFFF", bold=True, size=10)
    sec_font  = Font(name="Arial", color="FFFFFF", bold=True, size=11)
    bold_font = Font(name="Arial", bold=True, size=10)
    norm_font = Font(name="Arial", size=10)
    wb_font   = Font(name="Arial", color="FFFFFF", bold=True, size=10)
    net_p_fnt = Font(name="Arial", bold=True, size=10, color="27AE60")
    net_n_fnt = Font(name="Arial", bold=True, size=10, color="E74C3C")

    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    mfmt = "#,##0.00"
    pfmt = "0.0%"
    MIN_H = 20

    # ── Helpers ───────────────────────────────────────────────────────────────
    def autofit_columns(ws, min_width=8, max_width=60):
        """Size columns from literal cell content; enable wrap text; enforce min row height."""
        col_widths: dict[int, int] = {}
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                try:
                    curr = cell.alignment if cell.alignment else Alignment()
                    cell.alignment = Alignment(
                        horizontal=curr.horizontal,
                        vertical=curr.vertical or "center",
                        wrap_text=True,
                        indent=curr.indent,
                    )
                    if cell.value is not None and not str(cell.value).startswith("="):
                        col = cell.column
                        length = max(len(ln) for ln in str(cell.value).split("\n")) + 2
                        col_widths[col] = max(col_widths.get(col, min_width), length)
                except Exception:
                    pass
        for col, width in col_widths.items():
            ws.column_dimensions[get_column_letter(col)].width = min(max(width, min_width), max_width)
        for ri in ws.row_dimensions:
            if (ws.row_dimensions[ri].height or 0) < MIN_H:
                ws.row_dimensions[ri].height = MIN_H

    def sec_hdr(ws, r, text, ncols):
        c = ws.cell(row=r, column=1, value=text)
        c.font = sec_font; c.fill = sec_fill
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = 22
        if ncols > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)

    def col_hdrs(ws, r, hdrs, start_col=1):
        for ci, h in enumerate(hdrs, start_col):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = hdr_font; c.fill = sub_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin
        ws.row_dimensions[r].height = MIN_H

    def num_cell(ws, r, ci, formula, font=None, fill=None):
        c = ws.cell(row=r, column=ci, value=formula)
        c.font = font or norm_font; c.border = thin; c.number_format = mfmt
        c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        if fill:
            c.fill = fill

    def lbl_cell(ws, r, ci, text, font=None, fill=None):
        c = ws.cell(row=r, column=ci, value=text)
        c.font = font or norm_font; c.border = thin
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if fill:
            c.fill = fill

    # ── Breakdown sheet column layout (A–G) ──────────────────────────────────
    # A=Date  B=Description  C=Amount  D=Month  E=Category  F=Include  G=Reason
    # Data starts at row 3 (row 1 = banner, row 2 = headers)
    BD_DATA_START = 3
    BD_MAX_ROW = 2000   # fixed range ceiling for SUMPRODUCT formulas

    INC_SHEET = "Income Breakdown"
    EXP_SHEET = "Expense Breakdown"

    def sp(sheet, col, r1=BD_DATA_START, r2=BD_MAX_ROW):
        """Absolute fixed-range reference for SUMPRODUCT: 'Sheet'!$X$r1:$X$r2"""
        return f"'{sheet}'!${col}${r1}:${col}${r2}"

    def sumproduct_cat_month(sheet, cat, month):
        return (
            f"=SUMPRODUCT("
            f"({sp(sheet,'E')}=\"{cat}\")*"
            f"({sp(sheet,'D')}=\"{month}\")*"
            f"({sp(sheet,'F')}=\"Y\")*"
            f"{sp(sheet,'C')})"
        )

    def sumproduct_cat_annual(sheet, cat):
        return (
            f"=SUMPRODUCT("
            f"({sp(sheet,'E')}=\"{cat}\")*"
            f"({sp(sheet,'F')}=\"Y\")*"
            f"{sp(sheet,'C')})"
        )

    def sumproduct_cat_has_y(sheet, cat):
        return (
            f"=IF(SUMPRODUCT("
            f"({sp(sheet,'E')}=\"{cat}\")*"
            f"({sp(sheet,'F')}=\"Y\"))>0,\"Y\",\"N\")"
        )

    # ── Summary sheet row layout ──────────────────────────────────────────────
    # Section 1 (Monthly Income)
    S1_HDR        = 1
    S1_COLHDR     = 2
    S1_DATA_START = 3
    S1_DATA_END   = S1_DATA_START + max(len(income_months), 1) - 1
    S1_GRAND      = S1_DATA_END + 1
    S1_AVG        = S1_GRAND + 1

    # Section 2 (Monthly Expenses)
    S2_HDR        = S1_AVG + 2
    S2_COLHDR     = S2_HDR + 1
    S2_DATA_START = S2_COLHDR + 1
    S2_DATA_END   = S2_DATA_START + max(len(expense_months), 1) - 1
    S2_GRAND      = S2_DATA_END + 1
    S2_AVG        = S2_GRAND + 1

    # Section 3 (Annual Category Totals)
    S3_HDR        = S2_AVG + 2
    S3_COLHDR     = S3_HDR + 1
    S3_DATA_START = S3_COLHDR + 1
    n_cat_rows    = max(len(income_cats), len(expense_cats), 1)
    S3_DATA_END   = S3_DATA_START + n_cat_rows - 1

    # Section 4 (Key Metrics)
    S4_HDR        = S3_DATA_END + 2
    S4_DATA_START = S4_HDR + 1

    # Summary column counts
    # Section 1/2: col 1=Month, cols 2..1+n_cats=categories, col 2+n_cats=Total
    inc_total_col = len(income_cats) + 2   # "Total Income" column index
    exp_total_col = len(expense_cats) + 2  # "Total Expenses" column index
    max_data_cols = max(inc_total_col, exp_total_col, 9)

    # Key cell references used across sections
    inc_grand_ref = f"{get_column_letter(inc_total_col)}{S1_GRAND}"
    exp_grand_ref = f"{get_column_letter(exp_total_col)}{S2_GRAND}"

    # ── Build Workbook ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 1 — Monthly Income Breakdown
    # ════════════════════════════════════════════════════════════════════════
    sec_hdr(ws, S1_HDR, "SECTION 1 — MONTHLY INCOME BREAKDOWN  (Y transactions only)", inc_total_col)
    col_hdrs(ws, S1_COLHDR, ["Month"] + income_cats + ["Total Income"])

    for mi, mo in enumerate(income_months):
        r = S1_DATA_START + mi
        lbl_cell(ws, r, 1, mo)
        for ci, cat in enumerate(income_cats, 2):
            num_cell(ws, r, ci, sumproduct_cat_month(INC_SHEET, cat, mo))
        # Total Income = SUM of category cells in this row
        if income_cats:
            tot_f = f"=SUM({get_column_letter(2)}{r}:{get_column_letter(1+len(income_cats))}{r})"
        else:
            tot_f = "=0"
        num_cell(ws, r, inc_total_col, tot_f)
        ws.row_dimensions[r].height = MIN_H

    if not income_months:
        ws.cell(row=S1_DATA_START, column=1, value="(no income data)").font = norm_font

    # Grand Total row — SUM each column over all data rows
    lbl_cell(ws, S1_GRAND, 1, "Grand Total", bold_font, tot_fill)
    for ci in range(2, inc_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S1_GRAND, ci, f"=SUM({cl}{S1_DATA_START}:{cl}{S1_DATA_END})", bold_font, tot_fill)
    ws.row_dimensions[S1_GRAND].height = MIN_H

    # Monthly Average row — Grand Total / n_inc
    lbl_cell(ws, S1_AVG, 1, "Monthly Average", bold_font, tot_fill)
    for ci in range(2, inc_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S1_AVG, ci, f"={cl}{S1_GRAND}/{n_inc}", bold_font, tot_fill)
    ws.row_dimensions[S1_AVG].height = MIN_H

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 2 — Monthly Expense Breakdown
    # ════════════════════════════════════════════════════════════════════════
    sec_hdr(ws, S2_HDR, "SECTION 2 — MONTHLY EXPENSE BREAKDOWN  (Y transactions only)", exp_total_col)
    col_hdrs(ws, S2_COLHDR, ["Month"] + expense_cats + ["Total Expenses"])

    for mi, mo in enumerate(expense_months):
        r = S2_DATA_START + mi
        lbl_cell(ws, r, 1, mo)
        for ci, cat in enumerate(expense_cats, 2):
            num_cell(ws, r, ci, sumproduct_cat_month(EXP_SHEET, cat, mo))
        if expense_cats:
            tot_f = f"=SUM({get_column_letter(2)}{r}:{get_column_letter(1+len(expense_cats))}{r})"
        else:
            tot_f = "=0"
        num_cell(ws, r, exp_total_col, tot_f)
        ws.row_dimensions[r].height = MIN_H

    if not expense_months:
        ws.cell(row=S2_DATA_START, column=1, value="(no expense data)").font = norm_font

    lbl_cell(ws, S2_GRAND, 1, "Grand Total", bold_font, tot_fill)
    for ci in range(2, exp_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S2_GRAND, ci, f"=SUM({cl}{S2_DATA_START}:{cl}{S2_DATA_END})", bold_font, tot_fill)
    ws.row_dimensions[S2_GRAND].height = MIN_H

    lbl_cell(ws, S2_AVG, 1, "Monthly Average", bold_font, tot_fill)
    for ci in range(2, exp_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S2_AVG, ci, f"={cl}{S2_GRAND}/{n_exp}", bold_font, tot_fill)
    ws.row_dimensions[S2_AVG].height = MIN_H

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 3 — Annual Category Totals
    # ════════════════════════════════════════════════════════════════════════
    # Income cols 1-4, spacer col 5, Expense cols 6-9
    sec_hdr(ws, S3_HDR, "SECTION 3 — ANNUAL CATEGORY TOTALS", max(max_data_cols, 9))
    col_hdrs(ws, S3_COLHDR, ["Category", "Annual Total", "% of Total Income", "Include Status"], 1)
    col_hdrs(ws, S3_COLHDR, ["Category", "Annual Total", "% of Total Expenses", "Include Status"], 6)

    for idx in range(n_cat_rows):
        r = S3_DATA_START + idx
        ws.row_dimensions[r].height = MIN_H

        # ── Income side (cols 1-4) ──
        if idx < len(income_cats):
            cat = income_cats[idx]
            has_y = any(t.get("include", True) for t in income_txs if t.get("category") == cat)

            c1 = ws.cell(row=r, column=1, value=cat)
            c1.font = norm_font; c1.border = thin
            c1.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            c2 = ws.cell(row=r, column=2, value=sumproduct_cat_annual(INC_SHEET, cat))
            c2.font = norm_font; c2.border = thin; c2.number_format = mfmt
            c2.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

            c3 = ws.cell(row=r, column=3,
                         value=f"=IF({inc_grand_ref}<>0,B{r}/{inc_grand_ref},0)")
            c3.font = norm_font; c3.border = thin; c3.number_format = pfmt
            c3.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

            c4 = ws.cell(row=r, column=4, value=sumproduct_cat_has_y(INC_SHEET, cat))
            c4.font = wb_font; c4.fill = inc_y_fill if has_y else inc_n_fill
            c4.border = thin
            c4.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # ── Expense side (cols 6-9) ──
        if idx < len(expense_cats):
            cat = expense_cats[idx]
            has_y = any(t.get("include", True) for t in expense_txs if t.get("category") == cat)

            c6 = ws.cell(row=r, column=6, value=cat)
            c6.font = norm_font; c6.border = thin
            c6.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)

            c7 = ws.cell(row=r, column=7, value=sumproduct_cat_annual(EXP_SHEET, cat))
            c7.font = norm_font; c7.border = thin; c7.number_format = mfmt
            c7.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

            c8 = ws.cell(row=r, column=8,
                         value=f"=IF({exp_grand_ref}<>0,G{r}/{exp_grand_ref},0)")
            c8.font = norm_font; c8.border = thin; c8.number_format = pfmt
            c8.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

            c9 = ws.cell(row=r, column=9, value=sumproduct_cat_has_y(EXP_SHEET, cat))
            c9.font = wb_font; c9.fill = inc_y_fill if has_y else inc_n_fill
            c9.border = thin
            c9.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Conditional formatting: Include Status cells update color when formula flips Y↔N
    s3_inc_rng = f"D{S3_DATA_START}:D{S3_DATA_END}"
    s3_exp_rng = f"I{S3_DATA_START}:I{S3_DATA_END}"
    for rng in (s3_inc_rng, s3_exp_rng):
        ws.conditional_formatting.add(rng,
            CellIsRule(operator="equal", formula=['"Y"'], fill=inc_y_fill, font=wb_font))
        ws.conditional_formatting.add(rng,
            CellIsRule(operator="equal", formula=['"N"'], fill=inc_n_fill, font=wb_font))

    # ════════════════════════════════════════════════════════════════════════
    # SECTION 4 — Key Metrics (all cell-reference formulas)
    # ════════════════════════════════════════════════════════════════════════
    sec_hdr(ws, S4_HDR, "SECTION 4 — KEY METRICS", 4)

    metrics = [
        ("Total Annual Income",       f"={inc_grand_ref}",                           False),
        ("Total Annual Expenses",     f"={exp_grand_ref}",                           False),
        ("Net Annual Income",         f"={inc_grand_ref}-{exp_grand_ref}",           True),
        ("Monthly Average Income",    f"={inc_grand_ref}/{n_inc}",                   False),
        ("Monthly Average Expenses",  f"={exp_grand_ref}/{n_exp}",                   False),
        ("Monthly Average Net Income",f"=({inc_grand_ref}-{exp_grand_ref})/{n_all}", True),
    ]
    for i, (lbl, formula, is_net) in enumerate(metrics):
        r = S4_DATA_START + i
        lbl_cell(ws, r, 1, lbl, bold_font, met_fill)
        vc = ws.cell(row=r, column=2, value=formula)
        vc.fill = met_fill; vc.border = thin; vc.number_format = mfmt; vc.font = bold_font
        vc.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = MIN_H

    # Conditional formatting for Net rows: green ≥ 0, red < 0
    for i in (2, 5):   # Net Annual Income and Monthly Average Net Income
        nc = f"B{S4_DATA_START + i}"
        ws.conditional_formatting.add(nc, CellIsRule(
            operator="greaterThanOrEqual", formula=["0"], fill=net_pos, font=net_p_fnt))
        ws.conditional_formatting.add(nc, CellIsRule(
            operator="lessThan", formula=["0"], fill=met_fill, font=net_n_fnt))

    autofit_columns(ws)

    # ════════════════════════════════════════════════════════════════════════
    # BREAKDOWN SHEETS (Income Breakdown + Expense Breakdown)
    # Columns: A=Date  B=Description  C=Amount  D=Month  E=Category
    #          F=Include (Y/N dropdown)  G=Notes/Reason
    # ════════════════════════════════════════════════════════════════════════
    def build_breakdown_sheet(ws_b, txs):
        banner = (
            "GREEN rows (Y) are counted in Summary totals.  "
            "ORANGE rows (N) are excluded but visible for reference.  "
            "Change the Include dropdown (F column) to Y or N — "
            "the Summary sheet will recalculate automatically."
        )
        bc = ws_b.cell(row=1, column=1, value=banner)
        bc.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        bc.fill = hdr_fill
        bc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws_b.merge_cells("A1:G1")
        ws_b.row_dimensions[1].height = 40

        for ci, h in enumerate(
            ["Date", "Description", "Amount", "Month", "Category", "Include", "Notes / Reason"], 1
        ):
            c = ws_b.cell(row=2, column=ci, value=h)
            c.font = hdr_font; c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin
        ws_b.row_dimensions[2].height = MIN_H

        sorted_txs = sorted(txs, key=lambda t: t.get("date", ""))

        for ri, tx in enumerate(sorted_txs, BD_DATA_START):
            inc = tx.get("include", True)
            rf = green_row if inc else orange_row
            vals = [
                tx.get("date", ""),
                tx.get("description", ""),
                abs(tx.get("amount", 0)),
                tx.get("month", ""),
                tx.get("category", ""),
                "Y" if inc else "N",
                tx.get("reason", ""),
            ]
            for ci, val in enumerate(vals, 1):
                c = ws_b.cell(row=ri, column=ci, value=val)
                c.fill = rf; c.border = thin
                if ci == 3:      # Amount
                    c.font = norm_font; c.number_format = mfmt
                    c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
                elif ci == 6:    # Include — green/red cell
                    c.font = wb_font
                    c.fill = inc_y_fill if inc else inc_n_fill
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif ci == 7:    # Reason
                    c.font = norm_font
                    c.alignment = Alignment(vertical="center", wrap_text=True)
                else:
                    c.font = norm_font
                    c.alignment = Alignment(vertical="center", wrap_text=True)
            ws_b.row_dimensions[ri].height = MIN_H

        # Y/N dropdown on Include column (F), covering all data rows plus spare rows
        dv_last = max(len(sorted_txs) + BD_DATA_START + 50, BD_MAX_ROW)
        dv = DataValidation(
            type="list",
            formula1='"Y,N"',
            allow_blank=False,
            showDropDown=False,      # False = show the dropdown arrow in Excel
            showErrorMessage=True,
            error="Enter Y or N",
            errorTitle="Invalid value",
        )
        dv.sqref = f"F{BD_DATA_START}:F{dv_last}"
        ws_b.add_data_validation(dv)

        autofit_columns(ws_b)

    ws_inc = wb.create_sheet(INC_SHEET)
    build_breakdown_sheet(ws_inc, income_txs)
    ws_exp = wb.create_sheet(EXP_SHEET)
    build_breakdown_sheet(ws_exp, expense_txs)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    wb.save(tmp_path)

    return FileResponse(
        tmp_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"bank_statement_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )



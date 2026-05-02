import os
import io
import json
import re
import base64
import tempfile
import uuid
from datetime import datetime
from typing import Optional

import anthropic
from pypdf import PdfReader, PdfWriter
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse, HTMLResponse
from pydantic import BaseModel
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY")
if not ANTHROPIC_API_KEY:
    raise RuntimeError("ANTHROPIC_API_KEY environment variable is required")

client = anthropic.Anthropic(api_key=ANTHROPIC_API_KEY)
sessions: dict[str, dict] = {}


class UpdateTransactionRequest(BaseModel):
    include: bool
    reason: Optional[str] = None


def split_pdf_into_chunks(pdf_bytes: bytes, pages_per_chunk: int = 8) -> list[bytes]:
    reader = PdfReader(io.BytesIO(pdf_bytes))
    total = len(reader.pages)
    chunks = []
    for start in range(0, total, pages_per_chunk):
        writer = PdfWriter()
        for page_idx in range(start, min(start + pages_per_chunk, total)):
            writer.add_page(reader.pages[page_idx])
        buf = io.BytesIO()
        writer.write(buf)
        chunks.append(buf.getvalue())
    return chunks


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


PROMPT = """MORTGAGE UNDERWRITING — BANK STATEMENT INCOME & EXPENSE ANALYSIS
You are a mortgage underwriting analyst. Analyze the uploaded bank statement and return a structured JSON object with every transaction. Do not produce the Excel directly. Return clean structured data for human review first.
EXTRACTION: Extract every single transaction without exception — do not skip, group, summarize, or combine any lines. For each transaction return: date (YYYY-MM-DD), description (exact text from statement), amount (positive number), type (income or expense), month (e.g. Apr-25), category, suggested_include (Y or N), reason (plain English explanation). Do not include opening balance, closing balance, or running balance lines as transactions. Only extract actual debit and credit transactions.
LARGE STATEMENTS: If the statement has hundreds of transactions, you must still extract every single one. Do not stop early. Do not summarize groups of transactions. Every line item must appear as its own row.
CATEGORIES: Do not use a fixed list. Read the statement and assign the most specific label from the payee name or transfer type exactly as it appears. Income examples: E-Transfer In, ATM Deposit, Mobile Deposit, Cheque Deposit, Wire Transfer, Cash Deposit, Lendcare Loan, Internal Transfer In, NSF Reversal, Govt Rebate, or exact sender name. Expense examples: exact payee name such as Enbridge Gas or Cogeco, or type such as E-Transfer Out, Cash Withdrawal, Bank Charges, Cheque, CC Transfer, Transfer Out. Use Other Income or Other Expense only if truly unclear and explain in reason field. SPECIAL CATEGORY RULE: MSP fees, merchant services fees, point of sale fees, and terminal fees — always categorize these as Merchant Services Fee and always set suggested_include to Y. Never categorize these as Bank Charges.
CONSISTENCY RULE: If the same payee or description appears multiple times, it must always get the same category and the same suggested_include value every single time. Never categorize the same type of transaction differently on different dates.
INCLUDE / EXCLUDE LOGIC: Every transaction must appear in the output regardless of Y or N. Y or N only controls whether it counts in summary totals. The reviewer can flip any Y or N in the review screen.
INCOME — Auto-set N, flag as excluded: Internal transfers from own accounts — any transfer that appears to come from the same account holder at the same or another bank, look for keywords like TFR-FR, transfer from, own account. NSF reversals and re-credits after returned payments. Government rebates such as HST rebate, GST credit, Ontario Trillium Benefit. Any wire transfer regardless of amount — set N and reason: Review: wire transfer — verify source and whether this is qualifying income. Any deposit that is a reversal or return of a prior outgoing transaction — including e-transfers that were sent out and came back, if an e-transfer was sent out and then deposited back it is NOT income, set N and reason: Returned outgoing e-transfer — not new income. Any deposit with no clear identifiable source — reason must say: Review: unusual deposit — source unclear. Lendcare loan deposits — always set N with reason: Lendcare loan — verify if this should be counted as qualifying income. Any single deposit that is significantly larger than the average deposit amount in the statement — set N and reason: Review: unusually large deposit — verify source. Any deposit that appears to be a loan, line of credit draw, financing, or borrowing from any lender or financial institution — set N with reason: Review: possible loan deposit — verify if this is qualifying income. Look for keywords like loan, LOC, credit, financing, advance, lendcare, or any known lender name in the description.
INCOME — Auto-set Y: E-transfers from clearly external senders. ATM deposits, mobile deposits, cash deposits, cheque deposits from third parties. Regular recurring deposits that appear employment or business related. When uncertain default Y and note uncertainty in reason.
EXPENSE — Auto-set N, always excluded: ALL bank fees without exception — monthly plan fees, NSF fees, overdraft interest, e-transfer send fees, wire fees, service charges — always N, never override. Credit card payments and CC transfers. Transfers to own accounts at same or other institutions.
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
    sessionId: null,
    transactions: [],
    filter: 'all',
    search: '',
    updating: {},
    downloading: false,
    downloaded: false
  };

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
            ? '<svg class="spin w-5 h-5" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Analyzing with Claude AI &mdash; processing all pages...'
            : '<svg class="w-5 h-5" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>Analyze Statements') +
        '</button>' +
        (state.isAnalyzing ? '<p class="mt-3 text-center text-sm text-gray-400">Claude is reading every page of your statement(s) — larger files take several minutes. Please keep this tab open.</p>' : '') +
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
          sc('&uarr;','Total Income',fmt(totals.income),'included only','green') +
          sc('&darr;','Total Expenses',fmt(totals.expenses),'included only','orange') +
          sc('$','Net Income',fmt(totals.net),'income minus expenses',totals.net>=0?'green':'red') +
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
              '<div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl bg-green-100 text-green-600 flex items-center justify-center font-bold">&uarr;</div><p class="text-sm font-medium text-gray-700">Total Income</p></div>' +
              '<span class="text-xl font-bold text-green-600">' + fmt(totalIncome) + '</span>' +
            '</div>' +
            '<div class="flex items-center justify-between px-6 py-4">' +
              '<div class="flex items-center gap-3"><div class="w-10 h-10 rounded-xl bg-orange-100 text-orange-500 flex items-center justify-center font-bold">&darr;</div><p class="text-sm font-medium text-gray-700">Total Expenses</p></div>' +
              '<span class="text-xl font-bold text-orange-500">' + fmt(totalExpenses) + '</span>' +
            '</div>' +
            '<div class="flex items-center justify-between px-6 py-4 bg-gray-50">' +
              '<div class="flex items-center gap-3">' +
                '<div class="w-10 h-10 rounded-xl flex items-center justify-center font-bold ' + (net>=0?'bg-green-100 text-green-700':'bg-red-100 text-red-600') + '">$</div>' +
                '<div><p class="text-xs text-gray-400 font-medium uppercase tracking-wide">Net Income</p><p class="text-xs text-gray-400">Income minus Expenses</p></div>' +
              '</div>' +
              '<span class="text-2xl font-bold ' + netColor + '">' + fmt(net) + '</span>' +
            '</div>' +
          '</div>' +
        '</div>' +
        '<div class="bg-white rounded-2xl border border-gray-200 shadow-sm p-6 mb-6">' +
          '<h3 class="text-sm font-semibold text-gray-700 mb-3">What\'s in the Excel report</h3>' +
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
      setState({page:'upload',files:[],sessionId:null,transactions:[],
                filter:'all',search:'',analyzeError:null,downloaded:false});
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
      setState({isAnalyzing:false, sessionId:data.session_id, transactions:data.transactions, page:'review', filter:'all', search:''});
    } catch(e) {
      setState({isAnalyzing:false, analyzeError:e.message||'Failed to analyze. Please try again.'});
    }
  }

  async function handleToggle(txId) {
    var tx = state.transactions.find(function(t){return t.id===txId;});
    if (!tx || state.updating[txId]) return;
    var upd = Object.assign({},state.updating); upd[txId]=true;
    setState({updating:upd});
    try {
      var res = await fetch('/sessions/'+state.sessionId+'/transactions/'+txId,{
        method:'PUT', headers:{'Content-Type':'application/json'},
        body:JSON.stringify({include:!tx.include})
      });
      if (!res.ok) throw new Error('Failed');
      var updated = await res.json();
      var txs = state.transactions.map(function(t){return t.id===updated.id?updated:t;});
      var done = Object.assign({},state.updating); delete done[txId];
      setState({transactions:txs, updating:done});
    } catch(e) {
      var done2 = Object.assign({},state.updating); delete done2[txId];
      setState({updating:done2});
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


@app.post("/analyze")
async def analyze_statements(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")

    all_transactions = []

    for file in files:
        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"{file.filename} is not a PDF")

        pdf_bytes = await file.read()

        try:
            page_chunks = split_pdf_into_chunks(pdf_bytes, pages_per_chunk=8)
        except Exception as e:
            print(f"[WARN] Could not split {file.filename}: {e}")
            page_chunks = [pdf_bytes]

        total_chunks = len(page_chunks)
        print(f"[INFO] {file.filename}: {total_chunks} chunk(s)")

        for idx, chunk_bytes in enumerate(page_chunks):
            print(f"[INFO] Chunk {idx+1}/{total_chunks} — {file.filename}")
            pdf_b64 = base64.standard_b64encode(chunk_bytes).decode("utf-8")

            with client.messages.stream(
                model="claude-haiku-4-5-20251001",
                max_tokens=32000,
                messages=[{
                    "role": "user",
                    "content": [
                        {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": pdf_b64}},
                        {"type": "text", "text": PROMPT + f"\n\n(Page chunk {idx+1} of {total_chunks})"},
                    ],
                }],
            ) as stream:
                response_text = stream.get_final_text().strip()

            try:
                raw = extract_json_transactions(response_text)
            except Exception as e:
                print(f"[WARN] Parse error chunk {idx+1}: {e}")
                continue

            added = 0
            for tx in raw:
                tx = normalize_transaction(tx)
                tx["id"] = str(uuid.uuid4())
                all_transactions.append(tx)
                added += 1
            print(f"[INFO] Chunk {idx+1}/{total_chunks}: {added} transactions extracted")

    session_id = str(uuid.uuid4())
    sessions[session_id] = {"transactions": all_transactions}
    return {"session_id": session_id, "transactions": all_transactions, "count": len(all_transactions)}


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

    all_months = sorted(set(tx.get("month","") for tx in transactions if tx.get("month")), key=parse_month)
    income_months = sorted(set(tx.get("month","") for tx in income_txs if tx.get("month")), key=parse_month)
    expense_months = sorted(set(tx.get("month","") for tx in expense_txs if tx.get("month")), key=parse_month)
    income_cats = sorted(set(tx.get("category","") for tx in income_txs if tx.get("category")))
    expense_cats = sorted(set(tx.get("category","") for tx in expense_txs if tx.get("category")))

    green_row   = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    orange_row  = PatternFill(start_color="FFE0B2", end_color="FFE0B2", fill_type="solid")
    inc_y_fill  = PatternFill(start_color="00B050", end_color="00B050", fill_type="solid")
    inc_n_fill  = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    hdr_fill    = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    sec_fill    = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    sub_fill    = PatternFill(start_color="BDD7EE", end_color="BDD7EE", fill_type="solid")
    tot_fill    = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    met_fill    = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")

    hdr_font  = Font(name="Arial", color="FFFFFF", bold=True, size=10)
    sec_font  = Font(name="Arial", color="FFFFFF", bold=True, size=11)
    bold_font = Font(name="Arial", bold=True, size=10)
    norm_font = Font(name="Arial", size=10)
    wb_font   = Font(name="Arial", color="FFFFFF", bold=True, size=10)

    thin = Border(left=Side(style="thin"), right=Side(style="thin"),
                  top=Side(style="thin"), bottom=Side(style="thin"))
    mfmt = "#,##0.00"
    pfmt = "0.0%"

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"
    row = 1

    def sec_hdr(ws, r, text, ncols):
        c = ws.cell(row=r, column=1, value=text)
        c.font = sec_font; c.fill = sec_fill
        c.alignment = Alignment(horizontal="left", vertical="center")
        ws.row_dimensions[r].height = 22
        if ncols > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncols)
        return r + 1

    def col_hdrs(ws, r, hdrs):
        for ci, h in enumerate(hdrs, 1):
            c = ws.cell(row=r, column=ci, value=h)
            c.font = hdr_font; c.fill = sub_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = thin
        ws.row_dimensions[r].height = 18
        return r + 1

    # Section 1 — Monthly Income
    inc_cols = 1 + len(income_cats) + 1
    row = sec_hdr(ws, row, "SECTION 1 — MONTHLY INCOME BREAKDOWN  (Y transactions only)", inc_cols)
    row = col_hdrs(ws, row, ["Month"] + income_cats + ["Total Income"])
    cat_inc_tot = {c: 0.0 for c in income_cats}
    mon_inc_tot = {}
    for mo in income_months:
        mtxs = [t for t in income_txs if t.get("month")==mo and t.get("include",True)]
        rd = [mo]; rt = 0.0
        for cat in income_cats:
            v = sum(abs(t.get("amount",0)) for t in mtxs if t.get("category")==cat)
            cat_inc_tot[cat] += v; rd.append(v or None); rt += v
        rd.append(rt); mon_inc_tot[mo] = rt
        for ci, val in enumerate(rd, 1):
            c = ws.cell(row=row, column=ci, value=val); c.font = norm_font; c.border = thin
            if ci == 1: c.alignment = Alignment(horizontal="left", vertical="center")
            else: c.number_format = mfmt; c.alignment = Alignment(horizontal="right", vertical="center")
        row += 1
    grand_inc = sum(mon_inc_tot.values()); n_inc = len(income_months) or 1
    for lbl, vals in [("Grand Total",[cat_inc_tot[c] for c in income_cats]+[grand_inc]),
                      ("Monthly Average",[cat_inc_tot[c]/n_inc for c in income_cats]+[grand_inc/n_inc])]:
        for ci, val in enumerate([lbl]+vals, 1):
            c = ws.cell(row=row, column=ci, value=val); c.font = bold_font; c.fill = tot_fill; c.border = thin
            if ci == 1: c.alignment = Alignment(horizontal="left", vertical="center")
            else: c.number_format = mfmt; c.alignment = Alignment(horizontal="right", vertical="center")
        row += 1
    row += 1

    # Section 2 — Monthly Expenses
    exp_cols = 1 + len(expense_cats) + 1
    mx = max(inc_cols, exp_cols)
    row = sec_hdr(ws, row, "SECTION 2 — MONTHLY EXPENSE BREAKDOWN  (Y transactions only)", mx)
    row = col_hdrs(ws, row, ["Month"] + expense_cats + ["Total Expenses"])
    cat_exp_tot = {c: 0.0 for c in expense_cats}
    mon_exp_tot = {}
    for mo in expense_months:
        mtxs = [t for t in expense_txs if t.get("month")==mo and t.get("include",True)]
        rd = [mo]; rt = 0.0
        for cat in expense_cats:
            v = sum(abs(t.get("amount",0)) for t in mtxs if t.get("category")==cat)
            cat_exp_tot[cat] += v; rd.append(v or None); rt += v
        rd.append(rt); mon_exp_tot[mo] = rt
        for ci, val in enumerate(rd, 1):
            c = ws.cell(row=row, column=ci, value=val); c.font = norm_font; c.border = thin
            if ci == 1: c.alignment = Alignment(horizontal="left", vertical="center")
            else: c.number_format = mfmt; c.alignment = Alignment(horizontal="right", vertical="center")
        row += 1
    grand_exp = sum(mon_exp_tot.values()); n_exp = len(expense_months) or 1
    for lbl, vals in [("Grand Total",[cat_exp_tot[c] for c in expense_cats]+[grand_exp]),
                      ("Monthly Average",[cat_exp_tot[c]/n_exp for c in expense_cats]+[grand_exp/n_exp])]:
        for ci, val in enumerate([lbl]+vals, 1):
            c = ws.cell(row=row, column=ci, value=val); c.font = bold_font; c.fill = tot_fill; c.border = thin
            if ci == 1: c.alignment = Alignment(horizontal="left", vertical="center")
            else: c.number_format = mfmt; c.alignment = Alignment(horizontal="right", vertical="center")
        row += 1
    row += 1

    # Section 3 — Annual Category Totals
    row = sec_hdr(ws, row, "SECTION 3 — ANNUAL CATEGORY TOTALS", max(mx, 9))
    for ci, h in enumerate(["Category","Annual Total","% of Total Income","Include Status"], 1):
        c = ws.cell(row=row, column=ci, value=h); c.font = hdr_font; c.fill = sub_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = thin
    for ci, h in enumerate(["Category","Annual Total","% of Total Expenses","Include Status"], 6):
        c = ws.cell(row=row, column=ci, value=h); c.font = hdr_font; c.fill = sub_fill
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True); c.border = thin
    row += 1

    def has_included(txs, cat):
        return any(t.get("include", True) for t in txs if t.get("category") == cat)

    for r in range(max(len(income_cats), len(expense_cats), 1)):
        if r < len(income_cats):
            cat = income_cats[r]; tot = cat_inc_tot[cat]
            pct = tot/grand_inc if grand_inc else 0.0
            st = "Y" if has_included(income_txs, cat) else "N"
            for ci, val in enumerate([cat, tot, pct, st], 1):
                c = ws.cell(row=row+r, column=ci, value=val); c.border = thin
                if ci==1: c.font=norm_font
                elif ci==2: c.font=norm_font; c.number_format=mfmt; c.alignment=Alignment(horizontal="right")
                elif ci==3: c.font=norm_font; c.number_format=pfmt; c.alignment=Alignment(horizontal="right")
                else: c.font=wb_font; c.fill=inc_y_fill if st=="Y" else inc_n_fill; c.alignment=Alignment(horizontal="center")
        if r < len(expense_cats):
            cat = expense_cats[r]; tot = cat_exp_tot[cat]
            pct = tot/grand_exp if grand_exp else 0.0
            st = "Y" if has_included(expense_txs, cat) else "N"
            for ci, val in enumerate([cat, tot, pct, st], 6):
                c = ws.cell(row=row+r, column=ci, value=val); c.border = thin
                if ci==6: c.font=norm_font
                elif ci==7: c.font=norm_font; c.number_format=mfmt; c.alignment=Alignment(horizontal="right")
                elif ci==8: c.font=norm_font; c.number_format=pfmt; c.alignment=Alignment(horizontal="right")
                else: c.font=wb_font; c.fill=inc_y_fill if st=="Y" else inc_n_fill; c.alignment=Alignment(horizontal="center")
    row += max(len(income_cats), len(expense_cats), 1) + 1

    # Section 4 — Key Metrics
    row = sec_hdr(ws, row, "SECTION 4 — KEY METRICS", 4)
    n_all = len(all_months) or 1
    net = grand_inc - grand_exp
    for lbl, val in [("Total Annual Income",grand_inc),("Total Annual Expenses",grand_exp),("Net Annual Income",net),
                     ("Monthly Average Income",grand_inc/n_all),("Monthly Average Expenses",grand_exp/n_all),
                     ("Monthly Average Net Income",net/n_all)]:
        lc = ws.cell(row=row, column=1, value=lbl); vc = ws.cell(row=row, column=2, value=val)
        lc.font=bold_font; lc.fill=met_fill; lc.border=thin; lc.alignment=Alignment(horizontal="left",vertical="center")
        vc.fill=met_fill; vc.border=thin; vc.number_format=mfmt; vc.alignment=Alignment(horizontal="right",vertical="center")
        vc.font = Font(name="Arial",bold=True,size=10,color=("27AE60" if val>=0 else "E74C3C")) if "Net" in lbl else bold_font
        row += 1

    ws.column_dimensions["A"].width = 24
    for i in range(2, max(mx, 10)+1):
        ws.column_dimensions[get_column_letter(i)].width = 16

    # Breakdown sheet helper
    def build_sheet(ws_b, txs):
        banner = ("GREEN rows (Y) are counted in summary totals.  "
                  "ORANGE rows (N) are excluded but visible for reference.  "
                  "Change Include in the Review screen before re-exporting.")
        bc = ws_b.cell(row=1, column=1, value=banner)
        bc.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        bc.fill = hdr_fill; bc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws_b.merge_cells("A1:G1"); ws_b.row_dimensions[1].height = 36

        hdrs = ["Date","Description","Amount","Month","Category","Include","Notes / Reason"]
        widths = [14, 45, 14, 10, 26, 10, 50]
        for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
            c = ws_b.cell(row=2, column=ci, value=h)
            c.font = hdr_font; c.fill = hdr_fill
            c.alignment = Alignment(horizontal="center", vertical="center"); c.border = thin
            ws_b.column_dimensions[get_column_letter(ci)].width = w
        ws_b.row_dimensions[2].height = 18

        for ri, tx in enumerate(sorted(txs, key=lambda t: t.get("date","")), 3):
            inc = tx.get("include", True)
            rf = green_row if inc else orange_row
            vals = [tx.get("date",""), tx.get("description",""), abs(tx.get("amount",0)),
                    tx.get("month",""), tx.get("category",""), "Y" if inc else "N", tx.get("reason","")]
            for ci, val in enumerate(vals, 1):
                c = ws_b.cell(row=ri, column=ci, value=val); c.fill = rf; c.border = thin
                if ci == 3:
                    c.font = norm_font; c.number_format = mfmt; c.alignment = Alignment(horizontal="right",vertical="center")
                elif ci == 6:
                    c.font = wb_font; c.fill = inc_y_fill if inc else inc_n_fill
                    c.alignment = Alignment(horizontal="center",vertical="center")
                elif ci == 7:
                    c.font = norm_font; c.alignment = Alignment(vertical="center",wrap_text=True)
                else:
                    c.font = norm_font; c.alignment = Alignment(vertical="center")
            ws_b.row_dimensions[ri].height = 15

    ws_inc = wb.create_sheet("Income Breakdown")
    build_sheet(ws_inc, income_txs)
    ws_exp = wb.create_sheet("Expense Breakdown")
    build_sheet(ws_exp, expense_txs)

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp_path = tmp.name
    wb.save(tmp_path)

    return FileResponse(
        tmp_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=f"bank_statement_analysis_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
    )

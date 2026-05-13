import os
import json
import re
import uuid
import asyncio
import tempfile
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


# ─────────────────────────────────────────────────────────────────────────────
# JSON extraction utilities (robust parsing of Claude output)
# ─────────────────────────────────────────────────────────────────────────────

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


def deduplicate_categories(transactions: list) -> list:
    """Post-processing: normalize inconsistent category/include for same description+type."""
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


# ─────────────────────────────────────────────────────────────────────────────
# Stage 1 — Python PDF parser (no AI, deterministic)
# Works for BMO and most Canadian bank statements with month-day date format.
# ─────────────────────────────────────────────────────────────────────────────

# A transaction line in the PyMuPDF output always starts with a 3-letter month + day
DATE_RE   = re.compile(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}$', re.I)
AMOUNT_RE = re.compile(r'^-?[\d,]+\.\d{2}$')
PERIOD_RE = re.compile(r'for the period ending\s+\S+\s+\d{1,2},?\s+(\d{4})', re.I)
YEAR_RE   = re.compile(r'\b(20\d{2})\b')

MONTH_TO_NUM = {
    'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
    'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
    'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12',
}
NUM_TO_MONTH = {v: k.capitalize() for k, v in MONTH_TO_NUM.items()}

# Lines to always discard (matched against lowercased content)
DISCARD_EXACT = {
    'transaction details', 'transaction details (continued)',
    'amounts debited', 'amounts credited',
    'date', 'description',
    'from your account ($)', 'to your account ($)', 'balance ($)',
    '(continued)', 'continued',
    'business banking statement', 'business banking',
    'summary of account', 'number of items processed',
    'your branch address:', 'your branch', 'your plan',
    'direct banking', 'www.bmo.com',
    'essential plan $0 monthly fee',
    'for questions about your', 'statement call',
    'debited ($)', 'credited ($)', 'balance ($) on',
    'opening', 'total', 'closing', 'account', '-', '+', '=',
}

DISCARD_PATTERNS = [
    re.compile(r'^page \d+ of \d+$', re.I),           # "Page 1 of 17"
    re.compile(r'^\.*$'),                               # blank or all-dots
    re.compile(r'^\.{2,}'),                             # lines starting with dots
    re.compile(r'^-{3,}$'),                             # separator lines
    re.compile(r'^\d+$'),                               # bare integers (e.g. item counts 76, 516)
    re.compile(r'^business account\b', re.I),           # "Business Account # ..."
    re.compile(r'^business name:', re.I),
    re.compile(r'^#\s*\d'),                             # "# 3050 1981-377"
    re.compile(r'^transit number:', re.I),
    re.compile(r'^\(\d{3}\)\s*\d{3}'),                 # phone numbers
    re.compile(r'^1-\d{3}-\d{3}-\d{4}$'),              # 1-800 numbers
    re.compile(r'^www\.', re.I),                        # URLs
    re.compile(r'^[a-z]\d[a-z]\s*\d[a-z]\d$', re.I),  # Canadian postal codes like L6R3P4
    re.compile(r'^\d{5}\s+[A-Z]{2}'),                  # US-style zip + state
    re.compile(r'^for the period ending', re.I),        # period header line
]

# Date-prefixed lines that are NOT real transactions
NON_TX_DESCRIPTIONS = {
    'opening balance', 'closing totals', 'closing balance', 'balance forward',
}


def _should_discard(line: str) -> bool:
    low = line.lower()
    if low in DISCARD_EXACT:
        return True
    return any(p.match(low) for p in DISCARD_PATTERNS)


def _parse_amount(s: str) -> float:
    return float(s.replace(',', ''))


def extract_transactions_from_pdf(pdf_bytes: bytes, filename: str) -> list[dict]:
    """
    Pure Python extraction from a bank statement PDF.

    PyMuPDF linearises the table layout so each field (date, description,
    continuation lines, amount, balance) appears on its own text line.

    Algorithm:
      1. Collect all non-garbage lines from every page.
      2. Group into blocks, each starting with a date line (Mon DD).
      3. Within each block, the last two numbers are amount + balance.
         Everything in between is the transaction description.
      4. Determine credit/debit from balance delta — no guessing from text.
    """
    all_lines = []
    raw_pages = []

    try:
        doc = fitz.open(stream=pdf_bytes, filetype='pdf')
        for page in doc:
            page_text = page.get_text()
            raw_pages.append(page_text)
            for line in page_text.split('\n'):
                line = line.strip()
                if line and not _should_discard(line):
                    all_lines.append(line)
        doc.close()
    except Exception as e:
        print(f"[WARN] fitz error for {filename}: {e}")
        return []

    full_text = '\n'.join(raw_pages)

    # Detect statement year from "For the period ending … YYYY"
    year = datetime.now().year
    m = PERIOD_RE.search(full_text)
    if m:
        year = int(m.group(1))
    else:
        found = YEAR_RE.findall(full_text[:800])
        if found:
            year = int(found[0])

    print(f"[INFO] {filename}: year={year}, {len(all_lines)} lines after filtering")

    # Group lines into blocks — each block starts with a date line
    blocks = []
    current: list[str] = []
    for line in all_lines:
        if DATE_RE.match(line):
            if current:
                blocks.append(current)
            current = [line]
        elif current:
            current.append(line)
    if current:
        blocks.append(current)

    print(f"[INFO] {filename}: {len(blocks)} date-prefixed blocks found")

    transactions = []
    prev_balance: float | None = None

    for block in blocks:
        date_line = block[0]
        rest      = block[1:]

        # Pull trailing numbers off the end: second-to-last = amount, last = balance
        trailing: list[str] = []
        desc_parts: list[str] = []
        for item in reversed(rest):
            if AMOUNT_RE.match(item) and len(trailing) < 2:
                trailing.insert(0, item)
            else:
                desc_parts.insert(0, item)

        description = ' '.join(desc_parts).strip()

        # Only 1 trailing number → opening balance row (no transaction amount)
        if len(trailing) < 2:
            if trailing:
                try:
                    prev_balance = _parse_amount(trailing[0])
                except ValueError:
                    pass
            continue

        # Skip known non-transaction date rows
        if description.lower() in NON_TX_DESCRIPTIONS:
            continue

        amount_str, balance_str = trailing[0], trailing[1]
        try:
            amount  = _parse_amount(amount_str)
            balance = _parse_amount(balance_str)
        except ValueError:
            continue

        # Determine direction from balance delta — 100% accurate, no description guessing
        direction = 'debit'  # safe default
        if prev_balance is not None:
            delta = round(balance - prev_balance, 2)
            if abs(delta - amount) < 0.02:      # balance went up → credit
                direction = 'credit'
            elif abs(delta + amount) < 0.02:    # balance went down → debit
                direction = 'debit'
            else:
                # Delta mismatch — fall back to description keywords
                dl = description.lower()
                if any(k in dl for k in ['received', 'direct deposit', 'deposit', 'refund']):
                    direction = 'credit'
                elif any(k in dl for k in ['sent', 'purchase', 'payment', 'withdrawal',
                                            'fee', 'charge', 'draft', 'transfer out']):
                    direction = 'debit'
                else:
                    direction = 'debit'

        # Build date and month strings
        parts   = date_line.split()
        m_abbr  = parts[0].capitalize()
        day     = parts[1].zfill(2)
        mon_num = MONTH_TO_NUM.get(m_abbr.lower(), '01')
        full_date   = f"{year}-{mon_num}-{day}"
        month_field = f"{m_abbr}-{str(year)[2:]}"   # "Apr-26"

        transactions.append({
            'id':          str(uuid.uuid4()),
            'date':        full_date,
            'description': description,
            'amount':      amount,
            'direction':   direction,
            'month':       month_field,
        })
        prev_balance = balance

    credits = sum(1 for t in transactions if t['direction'] == 'credit')
    debits  = sum(1 for t in transactions if t['direction'] == 'debit')
    print(f"[INFO] {filename}: extracted {len(transactions)} transactions "
          f"({credits} credits / {debits} debits)")
    return transactions


# ─────────────────────────────────────────────────────────────────────────────
# Stage 2 — Claude categorisation prompt (categorisation only, no extraction)
# ─────────────────────────────────────────────────────────────────────────────

CAT_PROMPT = """MORTGAGE UNDERWRITING — TRANSACTION CATEGORISATION

You are a mortgage underwriting analyst. Bank transactions have been pre-extracted by a parser. Your only job is to assign category, suggested_include (Y or N), and reason to each transaction.

CRITICAL: The "direction" field is pre-determined from the bank statement and is always correct.
  direction "credit" = money coming IN  (income transaction)
  direction "debit"  = money going OUT  (expense transaction)
Never override the direction. Never reclassify a debit as income or a credit as expense.

Return ONLY a valid JSON array — no prose, no markdown fences.
Each element: {"id": "...", "category": "...", "suggested_include": "Y", "reason": "..."}
Every input transaction must appear in the output with its exact id.

── CATEGORIES ──────────────────────────────────────────────────────────────
Use the most specific label from the payee name or transfer type.
Income: INTERAC e-Transfer In, ATM Deposit, Mobile Deposit, Cheque Deposit, Direct Deposit, Wire Transfer, Cash Deposit, Internal Transfer In, NSF Reversal, Government Rebate, or exact sender name.
Expense: exact payee name (e.g. "Enbridge Gas", "CHIT CHATS BC") or transfer type such as INTERAC e-Transfer Out, Cash Withdrawal, Bank Charges, Cheque, CC Transfer, Transfer Out, Canadian Draft, Merchant Services Fee.
Special: MSP fees / merchant services fees / POS fees / terminal fees → always "Merchant Services Fee", always Y.
Use "Other Income" or "Other Expense" only if truly unclear; explain in reason.

── CONSISTENCY RULE (absolute) ─────────────────────────────────────────────
Identical or near-identical descriptions must always get identical category and identical suggested_include.
Before finalising, scan all transactions and fix any conflicts.

── INCOME RULES (direction = credit) ───────────────────────────────────────

Auto-set N:
- Internal own-account transfers: keywords TFR-FR, transfer from own, same account holder → reason: Internal transfer — not external income.
- NSF reversals and re-credits → reason: NSF reversal — not new income.
- Government rebates: HST rebate, GST credit, Ontario Trillium Benefit → reason: Government rebate — excluded.
- Wire transfers (any amount) → reason: Review: wire transfer — verify source and whether qualifying income.
- Loan or LOC deposits: keywords loan, LOC, credit line, financing, advance, lendcare, or any lender name → reason: Review: possible loan deposit — verify if qualifying income.
- Single deposit significantly larger than average deposits in this batch → reason: Review: unusually large deposit — verify source.
- Returned outgoing e-transfers (sent out then returned) → reason: Returned outgoing e-transfer — not new income.

Auto-set Y:
- INTERAC e-transfers from clearly external senders.
- ATM / mobile / cash / cheque deposits from third parties.
- Regular recurring deposits that appear employment or business-related.
- Direct deposits from known employers or processors (Square, Stripe, PayPal, etc.).
- When uncertain, default Y and note in reason.

── EXPENSE RULES (direction = debit) ───────────────────────────────────────

Auto-set N — never override these rules:
- ALL bank fees without exception: monthly plan fees, NSF fees, overdraft charges, overdraft per item charges, e-transfer fees, wire fees, service charges, transaction fees, draft fees, excess item fees → category: Bank Charges, reason: Bank fee — excluded.
- Credit card payments: keywords MC, VISA, AMEX, MASTERCARD, CAN TIRE MC, TD VISA, CIBC VISA, RBC VISA, SCOTIA VISA, BMO MC, or any description combining a card brand with alphanumeric characters → reason: Credit card payment — excluded.
- Transfers to own accounts at same or other institutions → reason: Internal transfer — excluded.
- CRA, Receiver General, Canada Revenue Agency, or any government tax remittance: keywords CRA, CANADA REVENUE, RECEIVER GENERAL, HST REMIT, GST REMIT, TAX REMIT → category: CRA / Tax Remittance, reason: CRA / government tax remittance — excluded from qualifying expenses. This rule is absolute — never set Y for CRA or Receiver General payments.

Auto-set Y:
- Insurance payments.
- Utilities: gas, hydro, internet, cable, phone, electricity.
- Loan and mortgage payments to external lenders.
- INTERAC e-transfers out to clearly external payees.
- Cheques to third-party individuals or businesses.
- Rent payments.
- Debit card purchases at identifiable merchants.
- Any regular, identifiable recurring expense.
- When uncertain, default Y.

── FLAGGING ────────────────────────────────────────────────────────────────
Unusual, one-time, very large, or unclear transactions → start reason with "Review:" followed by concern.
"""


# ─────────────────────────────────────────────────────────────────────────────
# Frontend HTML + JavaScript
# ─────────────────────────────────────────────────────────────────────────────

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
  var dragCounter = 0;   // fix: counter prevents jitter from child-element drag events

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
    if (state.filter === 'income')    r = r.filter(function(t){return t.type==='income';});
    else if (state.filter === 'expense')  r = r.filter(function(t){return t.type==='expense';});
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
    var income   = inc.filter(function(t){return t.type==='income';}).reduce(function(s,t){return s+t.amount;},0);
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
    var scrollY  = window.scrollY;
    var app = document.getElementById('app');
    if      (state.page === 'upload')   app.innerHTML = renderUpload();
    else if (state.page === 'progress') app.innerHTML = renderProgress();
    else if (state.page === 'review')   app.innerHTML = renderReview();
    else                                app.innerHTML = renderDownload();
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
      ? 'border-blue-500 bg-blue-50'
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

        '<!-- Hidden file input — triggered by label below or drag-drop -->' +
        '<input id="file-input" type="file" accept=".pdf,application/pdf" multiple class="hidden">' +

        '<!-- Drop zone -->' +
        '<div id="drop-zone" class="border-2 border-dashed rounded-2xl p-10 text-center transition-all duration-150 ' + dz + '">' +
          '<div class="flex flex-col items-center gap-4">' +
            '<div class="p-4 rounded-full ' + (state.isDragging ? 'bg-blue-100' : 'bg-gray-100') + '">' +
              '<svg class="w-8 h-8 ' + (state.isDragging ? 'text-blue-600' : 'text-gray-400') + '" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/></svg>' +
            '</div>' +
            '<div>' +
              '<p class="text-base font-semibold text-gray-700">' + (state.isDragging ? 'Drop your PDFs here' : 'Drag &amp; drop PDFs here') + '</p>' +
              '<p class="text-sm text-gray-400 mt-1">or use the button below to browse</p>' +
            '</div>' +
            '<!-- Label-based button: works natively on desktop and mobile without JS tricks -->' +
            '<label for="file-input" class="inline-flex items-center gap-2 px-5 py-2.5 bg-blue-600 text-white text-sm font-semibold rounded-xl cursor-pointer hover:bg-blue-700 active:scale-95 transition-all shadow-sm select-none">' +
              '<svg class="w-4 h-4" fill="none" viewBox="0 0 24 24" stroke="currentColor"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M15.172 7l-6.586 6.586a2 2 0 102.828 2.828l6.414-6.586a4 4 0 00-5.656-5.656l-6.415 6.585a6 6 0 108.486 8.486L20.5 13"/></svg>' +
              'Choose PDF Files' +
            '</label>' +
            '<p class="text-xs text-gray-400">Multiple files supported &mdash; one per month</p>' +
          '</div>' +
        '</div>' +

        (state.files.length > 0 ? '<div class="mt-4 space-y-2">' + fileRows + '</div>' : '') +
        (state.analyzeError ? '<div class="mt-4 flex items-center gap-3 bg-red-50 border border-red-200 rounded-xl px-4 py-3 text-red-600"><span class="text-lg">&#9888;</span><p class="text-sm">' + h(state.analyzeError) + '</p></div>' : '') +
        '<button id="analyze-btn" ' + (btnDis ? 'disabled' : '') + ' class="mt-6 w-full py-4 rounded-xl font-semibold text-base flex items-center justify-center gap-3 transition-all duration-200 ' + btnCls + '">' +
          (state.isAnalyzing
            ? '<svg class="spin w-5 h-5" viewBox="0 0 24 24" fill="none"><circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"/><path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4z"/></svg>Uploading files\u2026'
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
          '<p class="mt-2 text-gray-500">' + (isError ? 'An error occurred during processing.' : 'Processing your bank statements\u2026') + '</p>' +
        '</div>' +
        '<div class="bg-white rounded-2xl border border-gray-200 shadow-sm p-6 mb-6">' +
          '<p class="text-sm font-semibold text-gray-700 mb-4">' + h(info.current_file || 'Starting\u2026') + '</p>' +
          (info.total_pages > 0
            ? '<div class="mb-1"><div class="flex justify-between text-xs text-gray-400 mb-2"><span>Batches categorised</span><span>' + info.pages_done + ' / ' + info.total_pages + '</span></div>' +
              '<div class="w-full bg-gray-200 rounded-full h-2"><div class="bg-blue-600 h-2 rounded-full transition-all duration-500" style="width:' + pct + '%"></div></div></div>'
            : '<div class="flex items-center gap-2 text-gray-400 text-sm"><div class="w-2 h-2 rounded-full bg-blue-500 animate-pulse"></div><span>Extracting transactions\u2026</span></div>') +
          (isError ? '<p class="mt-4 text-sm text-red-600 bg-red-50 rounded-lg p-3">' + h(info.error || 'Unknown error') + '</p>' : '') +
        '</div>' +
        (isError
          ? '<button id="retry-btn" class="w-full py-4 rounded-xl font-semibold text-base bg-blue-600 text-white hover:bg-blue-700 shadow-md hover:shadow-lg transition-all">\u2190 Back to Upload</button>'
          : '<p class="text-center text-sm text-gray-400">Checking progress every 5 seconds \u2014 do not close this tab.</p>') +
      '</div>' +
    '</div>';
  }

  /* ============================================================
     REVIEW PAGE
  ============================================================ */
  function renderReview() {
    var totals   = getTotals();
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
          var upd  = state.updating[tx.id];
          var rBg  = tx.include ? 'bg-green-50 hover:bg-green-100' : 'bg-orange-50 hover:bg-orange-100';
          var bCls = upd
            ? 'opacity-50 cursor-wait bg-gray-300 text-white'
            : (tx.include ? 'bg-green-500 text-white hover:bg-green-600' : 'bg-orange-400 text-white hover:bg-orange-500');
          var aCls   = tx.amount >= 0 ? 'text-green-700' : 'text-red-600';
          var typCls = tx.type === 'income' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-600';
          return '<tr class="' + rBg + ' transition-colors">' +
            '<td class="px-4 py-3 font-mono text-xs text-gray-500 whitespace-nowrap">' + h(tx.date) + '</td>' +
            '<td class="px-4 py-3 text-gray-800 max-w-xs"><span class="clamp2">' + h(tx.description) + '</span></td>' +
            '<td class="px-4 py-3 text-right font-semibold tabular-nums whitespace-nowrap ' + aCls + '">' + (tx.amount>=0?'+':'') + fmt(tx.amount) + '</td>' +
            '<td class="px-4 py-3 text-center"><span class="inline-flex items-center px-2.5 py-0.5 rounded-full text-xs font-medium ' + typCls + '">' + h(tx.type) + '</span></td>' +
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
    var inc          = state.transactions.filter(function(t){return t.include;});
    var totalIncome  = inc.filter(function(t){return t.type==='income';}).reduce(function(s,t){return s+t.amount;},0);
    var totalExpenses= inc.filter(function(t){return t.type==='expense';}).reduce(function(s,t){return s+Math.abs(t.amount);},0);
    var net          = totalIncome - totalExpenses;
    var netColor     = net >= 0 ? 'text-green-700' : 'text-red-600';
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
            '<li>&#10003; Sheet 1: Summary \u2014 monthly income &amp; expense tables, annual totals, key metrics</li>' +
            '<li>&#10003; Sheet 2: Income Breakdown \u2014 every deposit with Y/N colour coding</li>' +
            '<li>&#10003; Sheet 3: Expense Breakdown \u2014 every withdrawal with Y/N colour coding</li>' +
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
    if      (state.page === 'upload')   attachUpload();
    else if (state.page === 'progress') attachProgress();
    else if (state.page === 'review')   attachReview();
    else                                attachDownload();
  }

  function attachUpload() {
    dragCounter = 0;   // reset counter whenever upload page re-renders
    var dz = document.getElementById('drop-zone');
    var fi = document.getElementById('file-input');
    var ab = document.getElementById('analyze-btn');

    if (dz) {
      // dragenter/dragleave counter prevents jitter when mouse passes over child elements
      dz.addEventListener('dragenter', function(e){
        e.preventDefault();
        dragCounter++;
        if (!state.isDragging) setState({isDragging: true});
      });
      dz.addEventListener('dragleave', function(e){
        dragCounter--;
        if (dragCounter <= 0) {
          dragCounter = 0;
          if (state.isDragging) setState({isDragging: false});
        }
      });
      dz.addEventListener('dragover', function(e){ e.preventDefault(); });
      dz.addEventListener('drop', function(e){
        e.preventDefault();
        dragCounter = 0;
        setState({isDragging: false});
        addFiles(e.dataTransfer.files);
      });
    }

    // File input change — fires after user picks files via the label button or drop
    if (fi) fi.addEventListener('change', function(e){
      addFiles(e.target.files);
      // Reset so the same file can be re-selected if removed then re-added
      fi.value = '';
    });

    // Remove buttons
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
    var pdfs = Array.from(list).filter(function(f){
      return f.type==='application/pdf' || f.name.toLowerCase().endsWith('.pdf');
    });
    var existing = {};
    state.files.forEach(function(f){ existing[f.name+f.size]=1; });
    var fresh = pdfs.filter(function(f){ return !existing[f.name+f.size]; });
    if (fresh.length) setState({files: state.files.concat(fresh), analyzeError:null});
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
        progressInfo: {status:'processing', current_file:'Starting\u2026', pages_done:0, total_pages:0}
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
          setState({progressInfo:{status:'error', error:'Failed to load transactions.', current_file:'', pages_done:0, total_pages:0}});
          return;
        }
        var transactions = await txRes.json();
        setState({page:'review', sessionId:data.session_id, transactions:transactions,
                  progressInfo:data, filter:'all', search:''});
      } else if (data.status === 'error') {
        stopPolling();
        setState({progressInfo: data});
      } else {
        setState({progressInfo: data});
      }
    } catch(e) { /* silent — retries on next interval */ }
  }

  async function handleToggle(txId) {
    var tx = state.transactions.find(function(t){return t.id===txId;});
    if (!tx || state.updating[txId]) return;
    var newInclude = !tx.include;

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
      var url  = URL.createObjectURL(blob);
      var a    = document.createElement('a');
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


# ─────────────────────────────────────────────────────────────────────────────
# FastAPI app
# ─────────────────────────────────────────────────────────────────────────────

app = FastAPI(title="Mortgage Bank Statement Analyzer")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=True,
                   allow_methods=["*"], allow_headers=["*"])


@app.get("/")
def index():
    return HTMLResponse(HTML)


@app.get("/health")
def health():
    return {"status": "ok"}


# ─────────────────────────────────────────────────────────────────────────────
# Background job: Stage 1 (Python extraction) → Stage 2 (Claude categorisation)
# ─────────────────────────────────────────────────────────────────────────────

async def _process_job(job_id: str, file_data: list[tuple[str, bytes]]):
    """
    For each uploaded PDF:
      Stage 1 — Python parser extracts exact structured transactions (no AI).
      Stage 2 — Claude categorises in batches of 150; only assigns category,
                 suggested_include, and reason — direction/amount already known.
      Stage 3 — Python dedup pass normalises any cross-batch inconsistencies.
    """
    BATCH_SIZE   = 150
    all_transactions: list[dict] = []
    total_files = len(file_data)

    try:
        for file_idx, (filename, pdf_bytes) in enumerate(file_data):
            jobs[job_id].update({
                "current_file": f"File {file_idx+1}/{total_files}: {filename} — extracting transactions…",
                "pages_done":   0,
                "total_pages":  0,
            })

            # ── Stage 1: Python structured extraction ─────────────────────────
            raw_txs = await asyncio.to_thread(
                extract_transactions_from_pdf, pdf_bytes, filename
            )

            if not raw_txs:
                print(f"[WARN] {filename}: 0 transactions extracted — "
                      f"check that this is a supported bank statement format")
                continue

            print(f"[INFO] {filename}: {len(raw_txs)} transactions ready for categorisation")

            # ── Stage 2: Claude categorisation ────────────────────────────────
            total_batches = (len(raw_txs) + BATCH_SIZE - 1) // BATCH_SIZE
            jobs[job_id]["total_pages"] = total_batches

            categorized: dict[str, dict] = {}   # id → {category, suggested_include, reason}

            for batch_num, start in enumerate(range(0, len(raw_txs), BATCH_SIZE), 1):
                batch = raw_txs[start: start + BATCH_SIZE]
                jobs[job_id]["current_file"] = (
                    f"File {file_idx+1}/{total_files}: {filename} — "
                    f"categorising batch {batch_num}/{total_batches}…"
                )

                # Send only what Claude needs — not re-sending what Python already knows
                batch_input = [
                    {
                        "id":          tx["id"],
                        "description": tx["description"],
                        "amount":      tx["amount"],
                        "direction":   tx["direction"],
                        "date":        tx["date"],
                        "month":       tx["month"],
                    }
                    for tx in batch
                ]

                prompt = CAT_PROMPT + "\n\nTransactions to categorise:\n" + json.dumps(batch_input)

                def _call(p=prompt) -> str:
                    with client.messages.stream(
                        model="claude-haiku-4-5-20251001",
                        max_tokens=16000,
                        messages=[{"role": "user", "content": p}],
                    ) as stream:
                        return stream.get_final_text().strip()

                try:
                    response_text = await asyncio.to_thread(_call)
                    results       = extract_json_transactions(response_text)
                    for item in results:
                        if "id" in item:
                            categorized[item["id"]] = item
                    print(f"[INFO] {filename} batch {batch_num}/{total_batches}: "
                          f"{len(results)} categorised")
                except Exception as e:
                    print(f"[WARN] Claude error — {filename} batch {batch_num}: {e}")

                jobs[job_id]["pages_done"] = batch_num

            # ── Stage 3: Merge ────────────────────────────────────────────────
            file_transactions: list[dict] = []
            for tx in raw_txs:
                cat      = categorized.get(tx["id"], {})
                tx_type  = "income" if tx["direction"] == "credit" else "expense"
                inc_flag = str(cat.get("suggested_include", "Y")).strip().upper() == "Y"
                # Expenses stored as negative, income as positive
                signed   = abs(tx["amount"]) if tx_type == "income" else -abs(tx["amount"])

                file_transactions.append({
                    "id":               tx["id"],
                    "date":             tx["date"],
                    "description":      tx["description"],
                    "amount":           signed,
                    "type":             tx_type,
                    "month":            tx["month"],
                    "category":         cat.get("category", "Other"),
                    "suggested_include": cat.get("suggested_include", "Y"),
                    "reason":           cat.get("reason", ""),
                    "include":          inc_flag,
                })

            all_transactions.extend(file_transactions)
            print(f"[INFO] {filename}: complete — {len(file_transactions)} transactions merged")

        # Dedup categories across all files for cross-batch consistency
        if all_transactions:
            all_transactions = deduplicate_categories(all_transactions)

        session_id = str(uuid.uuid4())
        sessions[session_id] = {"transactions": all_transactions}
        jobs[job_id].update({
            "status":            "complete",
            "session_id":        session_id,
            "transaction_count": len(all_transactions),
            "current_file":      f"Complete — {len(all_transactions)} transactions from {total_files} file(s)",
        })
        print(f"[INFO] Job {job_id} complete: {len(all_transactions)} transactions, "
              f"session {session_id}")

    except Exception as e:
        jobs[job_id] = {
            "status":      "error",
            "error":       str(e),
            "current_file": "",
            "pages_done":  0,
            "total_pages": 0,
        }
        print(f"[ERROR] Job {job_id} failed: {e}")


# ─────────────────────────────────────────────────────────────────────────────
# API routes
# ─────────────────────────────────────────────────────────────────────────────

@app.post("/analyze")
async def analyze_statements(files: list[UploadFile] = File(...)):
    if not files:
        raise HTTPException(status_code=400, detail="No files uploaded")
    for file in files:
        if not file.filename.lower().endswith(".pdf"):
            raise HTTPException(status_code=400, detail=f"{file.filename} is not a PDF")

    file_data = [(f.filename, await f.read()) for f in files]
    job_id    = str(uuid.uuid4())
    jobs[job_id] = {
        "status":      "processing",
        "current_file": f"Starting — 0 of {len(file_data)} files processed",
        "pages_done":  0,
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
    resp = {
        "status":       job["status"],
        "current_file": job.get("current_file", ""),
        "pages_done":   job.get("pages_done", 0),
        "total_pages":  job.get("total_pages", 0),
    }
    if job["status"] == "complete":
        resp["session_id"]        = job["session_id"]
        resp["transaction_count"] = job["transaction_count"]
    elif job["status"] == "error":
        resp["error"] = job.get("error", "Unknown error")
    return resp


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
    income_txs   = [tx for tx in transactions if tx.get("type") == "income"]
    expense_txs  = [tx for tx in transactions if tx.get("type") == "expense"]

    def parse_month(m):
        try:    return datetime.strptime(m, "%b-%y")
        except: return datetime.min

    all_months     = sorted(set(tx.get("month","") for tx in transactions if tx.get("month")), key=parse_month)
    income_months  = sorted(set(tx.get("month","") for tx in income_txs  if tx.get("month")), key=parse_month)
    expense_months = sorted(set(tx.get("month","") for tx in expense_txs if tx.get("month")), key=parse_month)
    income_cats    = sorted(set(tx.get("category","") for tx in income_txs  if tx.get("category")))
    expense_cats   = sorted(set(tx.get("category","") for tx in expense_txs if tx.get("category")))

    n_inc = len(income_months)  or 1
    n_exp = len(expense_months) or 1
    n_all = len(all_months)     or 1

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
                  top=Side(style="thin"),  bottom=Side(style="thin"))
    mfmt = "#,##0.00"
    pfmt = "0.0%"
    MIN_H = 20

    # ── Helpers ───────────────────────────────────────────────────────────────
    def autofit_columns(ws, min_width=8, max_width=60):
        col_widths: dict[int, int] = {}
        for row_cells in ws.iter_rows():
            for cell in row_cells:
                try:
                    curr = cell.alignment if cell.alignment else Alignment()
                    cell.alignment = Alignment(
                        horizontal=curr.horizontal, vertical=curr.vertical or "center",
                        wrap_text=True, indent=curr.indent,
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
        if fill: c.fill = fill

    def lbl_cell(ws, r, ci, text, font=None, fill=None):
        c = ws.cell(row=r, column=ci, value=text)
        c.font = font or norm_font; c.border = thin
        c.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        if fill: c.fill = fill

    # ── Breakdown sheet constants ─────────────────────────────────────────────
    BD_DATA_START = 3
    BD_MAX_ROW    = 2000
    INC_SHEET     = "Income Breakdown"
    EXP_SHEET     = "Expense Breakdown"

    def sp(sheet, col, r1=BD_DATA_START, r2=BD_MAX_ROW):
        return f"'{sheet}'!${col}${r1}:${col}${r2}"

    def sumproduct_cat_month(sheet, cat, month):
        return (f"=SUMPRODUCT("
                f"({sp(sheet,'E')}=\"{cat}\")*"
                f"({sp(sheet,'D')}=\"{month}\")*"
                f"({sp(sheet,'F')}=\"Y\")*"
                f"{sp(sheet,'C')})")

    def sumproduct_cat_annual(sheet, cat):
        return (f"=SUMPRODUCT("
                f"({sp(sheet,'E')}=\"{cat}\")*"
                f"({sp(sheet,'F')}=\"Y\")*"
                f"{sp(sheet,'C')})")

    def sumproduct_cat_has_y(sheet, cat):
        return (f"=IF(SUMPRODUCT("
                f"({sp(sheet,'E')}=\"{cat}\")*"
                f"({sp(sheet,'F')}=\"Y\"))>0,\"Y\",\"N\")")

    # ── Summary sheet row layout ──────────────────────────────────────────────
    S1_HDR        = 1
    S1_COLHDR     = 2
    S1_DATA_START = 3
    S1_DATA_END   = S1_DATA_START + max(len(income_months),  1) - 1
    S1_GRAND      = S1_DATA_END + 1
    S1_AVG        = S1_GRAND + 1

    S2_HDR        = S1_AVG + 2
    S2_COLHDR     = S2_HDR + 1
    S2_DATA_START = S2_COLHDR + 1
    S2_DATA_END   = S2_DATA_START + max(len(expense_months), 1) - 1
    S2_GRAND      = S2_DATA_END + 1
    S2_AVG        = S2_GRAND + 1

    S3_HDR        = S2_AVG + 2
    S3_COLHDR     = S3_HDR + 1
    S3_DATA_START = S3_COLHDR + 1
    n_cat_rows    = max(len(income_cats), len(expense_cats), 1)
    S3_DATA_END   = S3_DATA_START + n_cat_rows - 1

    S4_HDR        = S3_DATA_END + 2
    S4_DATA_START = S4_HDR + 1

    inc_total_col = len(income_cats)  + 2
    exp_total_col = len(expense_cats) + 2
    max_data_cols = max(inc_total_col, exp_total_col, 9)

    inc_grand_ref = f"{get_column_letter(inc_total_col)}{S1_GRAND}"
    exp_grand_ref = f"{get_column_letter(exp_total_col)}{S2_GRAND}"

    # ── Build Workbook ────────────────────────────────────────────────────────
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Summary"

    # Section 1 — Monthly Income
    sec_hdr(ws, S1_HDR, "SECTION 1 — MONTHLY INCOME BREAKDOWN  (Y transactions only)", inc_total_col)
    col_hdrs(ws, S1_COLHDR, ["Month"] + income_cats + ["Total Income"])
    for mi, mo in enumerate(income_months):
        r = S1_DATA_START + mi
        lbl_cell(ws, r, 1, mo)
        for ci, cat in enumerate(income_cats, 2):
            num_cell(ws, r, ci, sumproduct_cat_month(INC_SHEET, cat, mo))
        tot_f = f"=SUM({get_column_letter(2)}{r}:{get_column_letter(1+len(income_cats))}{r})" if income_cats else "=0"
        num_cell(ws, r, inc_total_col, tot_f)
        ws.row_dimensions[r].height = MIN_H
    if not income_months:
        ws.cell(row=S1_DATA_START, column=1, value="(no income data)").font = norm_font
    lbl_cell(ws, S1_GRAND, 1, "Grand Total", bold_font, tot_fill)
    for ci in range(2, inc_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S1_GRAND, ci, f"=SUM({cl}{S1_DATA_START}:{cl}{S1_DATA_END})", bold_font, tot_fill)
    ws.row_dimensions[S1_GRAND].height = MIN_H
    lbl_cell(ws, S1_AVG, 1, "Monthly Average", bold_font, tot_fill)
    for ci in range(2, inc_total_col + 1):
        cl = get_column_letter(ci)
        num_cell(ws, S1_AVG, ci, f"={cl}{S1_GRAND}/{n_inc}", bold_font, tot_fill)
    ws.row_dimensions[S1_AVG].height = MIN_H

    # Section 2 — Monthly Expenses
    sec_hdr(ws, S2_HDR, "SECTION 2 — MONTHLY EXPENSE BREAKDOWN  (Y transactions only)", exp_total_col)
    col_hdrs(ws, S2_COLHDR, ["Month"] + expense_cats + ["Total Expenses"])
    for mi, mo in enumerate(expense_months):
        r = S2_DATA_START + mi
        lbl_cell(ws, r, 1, mo)
        for ci, cat in enumerate(expense_cats, 2):
            num_cell(ws, r, ci, sumproduct_cat_month(EXP_SHEET, cat, mo))
        tot_f = f"=SUM({get_column_letter(2)}{r}:{get_column_letter(1+len(expense_cats))}{r})" if expense_cats else "=0"
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

    # Section 3 — Annual Category Totals
    sec_hdr(ws, S3_HDR, "SECTION 3 — ANNUAL CATEGORY TOTALS", max(max_data_cols, 9))
    col_hdrs(ws, S3_COLHDR, ["Category", "Annual Total", "% of Total Income",    "Include Status"], 1)
    col_hdrs(ws, S3_COLHDR, ["Category", "Annual Total", "% of Total Expenses",  "Include Status"], 6)
    for idx in range(n_cat_rows):
        r = S3_DATA_START + idx
        ws.row_dimensions[r].height = MIN_H
        if idx < len(income_cats):
            cat   = income_cats[idx]
            has_y = any(t.get("include", True) for t in income_txs if t.get("category") == cat)
            c1 = ws.cell(row=r, column=1, value=cat);  c1.font=norm_font; c1.border=thin; c1.alignment=Alignment(horizontal="left",  vertical="center", wrap_text=True)
            c2 = ws.cell(row=r, column=2, value=sumproduct_cat_annual(INC_SHEET, cat)); c2.font=norm_font; c2.border=thin; c2.number_format=mfmt; c2.alignment=Alignment(horizontal="right", vertical="center", wrap_text=True)
            c3 = ws.cell(row=r, column=3, value=f"=IF({inc_grand_ref}<>0,B{r}/{inc_grand_ref},0)"); c3.font=norm_font; c3.border=thin; c3.number_format=pfmt; c3.alignment=Alignment(horizontal="right", vertical="center", wrap_text=True)
            c4 = ws.cell(row=r, column=4, value=sumproduct_cat_has_y(INC_SHEET, cat)); c4.font=wb_font; c4.fill=inc_y_fill if has_y else inc_n_fill; c4.border=thin; c4.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
        if idx < len(expense_cats):
            cat   = expense_cats[idx]
            has_y = any(t.get("include", True) for t in expense_txs if t.get("category") == cat)
            c6 = ws.cell(row=r, column=6, value=cat);  c6.font=norm_font; c6.border=thin; c6.alignment=Alignment(horizontal="left",  vertical="center", wrap_text=True)
            c7 = ws.cell(row=r, column=7, value=sumproduct_cat_annual(EXP_SHEET, cat)); c7.font=norm_font; c7.border=thin; c7.number_format=mfmt; c7.alignment=Alignment(horizontal="right", vertical="center", wrap_text=True)
            c8 = ws.cell(row=r, column=8, value=f"=IF({exp_grand_ref}<>0,G{r}/{exp_grand_ref},0)"); c8.font=norm_font; c8.border=thin; c8.number_format=pfmt; c8.alignment=Alignment(horizontal="right", vertical="center", wrap_text=True)
            c9 = ws.cell(row=r, column=9, value=sumproduct_cat_has_y(EXP_SHEET, cat)); c9.font=wb_font; c9.fill=inc_y_fill if has_y else inc_n_fill; c9.border=thin; c9.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)

    for rng in (f"D{S3_DATA_START}:D{S3_DATA_END}", f"I{S3_DATA_START}:I{S3_DATA_END}"):
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Y"'], fill=inc_y_fill, font=wb_font))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"N"'], fill=inc_n_fill, font=wb_font))

    # Section 4 — Key Metrics
    sec_hdr(ws, S4_HDR, "SECTION 4 — KEY METRICS", 4)
    metrics = [
        ("Total Annual Income",        f"={inc_grand_ref}",                          False),
        ("Total Annual Expenses",      f"={exp_grand_ref}",                          False),
        ("Net Annual Income",          f"={inc_grand_ref}-{exp_grand_ref}",          True),
        ("Monthly Average Income",     f"={inc_grand_ref}/{n_inc}",                  False),
        ("Monthly Average Expenses",   f"={exp_grand_ref}/{n_exp}",                  False),
        ("Monthly Average Net Income", f"=({inc_grand_ref}-{exp_grand_ref})/{n_all}",True),
    ]
    for i, (lbl, formula, _) in enumerate(metrics):
        r  = S4_DATA_START + i
        lbl_cell(ws, r, 1, lbl, bold_font, met_fill)
        vc = ws.cell(row=r, column=2, value=formula)
        vc.fill=met_fill; vc.border=thin; vc.number_format=mfmt; vc.font=bold_font
        vc.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)
        ws.row_dimensions[r].height = MIN_H
    for i in (2, 5):
        nc = f"B{S4_DATA_START + i}"
        ws.conditional_formatting.add(nc, CellIsRule(operator="greaterThanOrEqual", formula=["0"],  fill=net_pos,  font=net_p_fnt))
        ws.conditional_formatting.add(nc, CellIsRule(operator="lessThan",           formula=["0"],  fill=met_fill, font=net_n_fnt))

    autofit_columns(ws)

    # ── Breakdown sheets ──────────────────────────────────────────────────────
    def build_breakdown_sheet(ws_b, txs):
        banner = ("GREEN rows (Y) are counted in Summary totals.  "
                  "ORANGE rows (N) are excluded but visible for reference.  "
                  "Change the Include dropdown (F column) to Y or N — "
                  "the Summary sheet will recalculate automatically.")
        bc = ws_b.cell(row=1, column=1, value=banner)
        bc.font = Font(name="Arial", bold=True, size=10, color="FFFFFF")
        bc.fill = hdr_fill
        bc.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws_b.merge_cells("A1:G1")
        ws_b.row_dimensions[1].height = 40

        for ci, h in enumerate(["Date","Description","Amount","Month","Category","Include","Notes / Reason"], 1):
            c = ws_b.cell(row=2, column=ci, value=h)
            c.font=hdr_font; c.fill=hdr_fill
            c.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border=thin
        ws_b.row_dimensions[2].height = MIN_H

        sorted_txs = sorted(txs, key=lambda t: t.get("date",""))
        for ri, tx in enumerate(sorted_txs, BD_DATA_START):
            inc  = tx.get("include", True)
            rf   = green_row if inc else orange_row
            vals = [
                tx.get("date",""),
                tx.get("description",""),
                abs(tx.get("amount", 0)),
                tx.get("month",""),
                tx.get("category",""),
                "Y" if inc else "N",
                tx.get("reason",""),
            ]
            for ci, val in enumerate(vals, 1):
                c = ws_b.cell(row=ri, column=ci, value=val)
                c.fill=rf; c.border=thin
                if ci == 3:
                    c.font=norm_font; c.number_format=mfmt
                    c.alignment=Alignment(horizontal="right", vertical="center", wrap_text=True)
                elif ci == 6:
                    c.font=wb_font; c.fill=inc_y_fill if inc else inc_n_fill
                    c.alignment=Alignment(horizontal="center", vertical="center", wrap_text=True)
                elif ci == 7:
                    c.font=norm_font; c.alignment=Alignment(vertical="center", wrap_text=True)
                else:
                    c.font=norm_font; c.alignment=Alignment(vertical="center", wrap_text=True)
            ws_b.row_dimensions[ri].height = MIN_H

        dv_last = max(len(sorted_txs) + BD_DATA_START + 50, BD_MAX_ROW)
        dv = DataValidation(
            type="list", formula1='"Y,N"', allow_blank=False,
            showDropDown=False, showErrorMessage=True,
            error="Enter Y or N", errorTitle="Invalid value",
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

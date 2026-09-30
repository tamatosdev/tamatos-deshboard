// ============================================
// TAMATOS FINANCE DASHBOARD - DATA FETCHER
// Source: Tamatos_Daily_Finance_Upload_Template_1
// Tabs: Invoices, Expenses, Bank_Transactions
// ============================================

const SHEET_ID = "15X-awF8nl5mtW7wz0lQA5YYPGEdpIxoTC5LcM18Knwk";

function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  const action = params.action ? params.action.toLowerCase() : "overdue";

  // Optional global date-range filter. Absent params => null => no filtering
  // (fully backward compatible with the original no-param calls).
  const fromDate = toRangeStart_(params.from);
  const toDate = toRangeEnd_(params.to);

  try {
    if (action === "requestotp") return requestOTP(params.email);
    if (action === "verifyotp") return verifyOTP(params.email, params.code);
    if (action === "revenue") return getRevenueSummary();
    if (action === "breakdown") return getBreakdown(fromDate, toDate);
    if (action === "expenses") return getUpcomingPayments();
    if (action === "expensesummary") return getExpenseSummary();
    if (action === "accounts") return getAccounts();
    if (action === "debug") return getDebugInfo();
    if (action === "flagged") return getFlaggedSpend();
    if (action === "all") return getCombined(fromDate, toDate);
    return getOverdueInvoices(fromDate, toDate);
  } catch (error) {
    return errorResponse(error.toString());
  }
}

// ============================================
// EMAIL OTP LOGIN
// NOTE: this gates the dashboard UI only. The data actions below are not
// token-checked, so anyone holding the SCRIPT_URL can still call ?action=all
// directly. Accepted limitation for now.
// ============================================
const ALLOWED_EMAILS = [
  "ebad.khan@tamatos.com"
  // Add up to 5 more once provided, e.g.:
  // , "someone@tamatos.com"
];
const OTP_EXPIRY_MINUTES = 10;

function requestOTP(email) {
  const normalized = String(email || "").trim().toLowerCase();
  const allowed = ALLOWED_EMAILS.map(e => e.toLowerCase());
  if (!normalized || allowed.indexOf(normalized) === -1) {
    return jsonResponse({ success: false, error: "This email is not authorized." });
  }

  const code = String(Math.floor(100000 + Math.random() * 900000));
  const expiresAt = Date.now() + OTP_EXPIRY_MINUTES * 60 * 1000;
  PropertiesService.getScriptProperties().setProperty("otp_" + normalized, code + "|" + expiresAt);
  // A fresh code clears any previous brute-force lockout
  PropertiesService.getScriptProperties().deleteProperty("otp_attempts_" + normalized);

  try {
    MailApp.sendEmail({
      to: normalized,
      subject: "Your Tamatos Dashboard login code",
      body: "Your login code is: " + code + "\n\nThis code expires in " + OTP_EXPIRY_MINUTES + " minutes.\n\nIf you didn't request this, ignore this email."
    });
  } catch (err) {
    return jsonResponse({ success: false, error: "Could not send email: " + err.toString() });
  }

  return jsonResponse({ success: true, message: "Code sent" });
}

function verifyOTP(email, code) {
  const normalized = String(email || "").trim().toLowerCase();
  const submittedCode = String(code || "").trim();
  const props = PropertiesService.getScriptProperties();
  const key = "otp_" + normalized;
  const attemptsKey = "otp_attempts_" + normalized;
  const stored = props.getProperty(key);

  if (!stored) {
    return jsonResponse({ success: false, error: "No code was requested for this email, or it already expired." });
  }

  const parts = stored.split("|");
  if (Date.now() > Number(parts[1])) {
    props.deleteProperty(key);
    return jsonResponse({ success: false, error: "Code expired. Please request a new code." });
  }

  // Brute-force lockout: 5 incorrect attempts within one code's lifetime.
  // Checked before the code comparison, so the correct code is refused too.
  const attempts = Number(props.getProperty(attemptsKey) || "0");
  if (attempts >= 5) {
    return jsonResponse({ success: false, error: "Too many incorrect attempts. Please request a new code." });
  }

  if (submittedCode !== parts[0]) {
    props.setProperty(attemptsKey, String(attempts + 1));
    return jsonResponse({ success: false, error: "Incorrect code." });
  }

  props.deleteProperty(key); // one-time use
  props.deleteProperty(attemptsKey);
  return jsonResponse({ success: true, token: Utilities.getUuid(), email: normalized });
}

// ============================================
// Shared helpers
// ============================================
function getTab_(tabName) {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  const sheet = spreadsheet.getSheetByName(tabName);
  if (!sheet) throw new Error('Tab "' + tabName + '" not found');
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues() : [];
  return { headers, rows };
}

function findCol(headers, searchTerms) {
  const terms = Array.isArray(searchTerms) ? searchTerms : [searchTerms];
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).toLowerCase().trim();
    if (terms.some(term => header.includes(term.toLowerCase()))) {
      return i + 1;
    }
  }
  return null;
}

function formatDate(date) {
  if (!date || !(date instanceof Date)) return "N/A";
  return date.toLocaleDateString('en-US', { year: 'numeric', month: 'short', day: 'numeric' });
}

// Handles real Date objects, Apps Script's date-like objects (which sometimes fail
// strict instanceof Date checks), and dates typed/pasted as plain text.
function toDate_(value) {
  if (value === null || value === undefined || value === "") return null;

  if (value instanceof Date && !isNaN(value.getTime())) return value;

  // Apps Script occasionally returns a date-like object that isn't strictly instanceof Date
  if (typeof value === "object" && typeof value.getTime === "function") {
    const t = value.getTime();
    if (!isNaN(t)) return new Date(t);
  }

  if (typeof value === "string") {
    const parsed = new Date(value.trim());
    if (!isNaN(parsed.getTime())) return parsed;
  }

  if (typeof value === "number") {
    const parsed = new Date(value);
    if (!isNaN(parsed.getTime())) return parsed;
  }

  return null;
}

// ---- Global date-range filter helpers ----
// Range bounds arrive as ISO date strings ("2026-06-01"). new Date("2026-06-01")
// is parsed as UTC midnight, which can land on the previous local day depending
// on the script timezone — so the components are read off explicitly and the
// boundaries are widened to whole local days to keep the range inclusive.
function toRangeStart_(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const iso = String(value).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 0, 0, 0, 0);
  const d = toDate_(value);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 0, 0, 0, 0);
}

function toRangeEnd_(value) {
  if (value === null || value === undefined || String(value).trim() === "") return null;
  const iso = String(value).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (iso) return new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]), 23, 59, 59, 999);
  const d = toDate_(value);
  if (!d) return null;
  return new Date(d.getFullYear(), d.getMonth(), d.getDate(), 23, 59, 59, 999);
}

// True when no range is set, otherwise an inclusive [from, to] comparison.
function inRange_(date, fromDate, toDate) {
  if (!fromDate && !toDate) return true;
  if (!date) return false;
  if (fromDate && date < fromDate) return false;
  if (toDate && date > toDate) return false;
  return true;
}

function rangeMeta_(fromDate, toDate) {
  return {
    rangeApplied: !!(fromDate || toDate),
    from: fromDate ? formatDate(fromDate) : null,
    to: toDate ? formatDate(toDate) : null
  };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function errorResponse(error) {
  return jsonResponse({ success: false, error: error, lastFetched: new Date().toISOString() });
}

// ============================================
// INVOICES TAB
// ============================================

function invoiceCols_(headers) {
  return {
    invNumber: findCol(headers, ["Invoice #", "Invoice Number"]),
    company: findCol(headers, ["Company"]),
    client: findCol(headers, ["Client Name", "Client"]),
    type: findCol(headers, ["Type"]),
    currency: findCol(headers, ["Currency"]),
    amount: findCol(headers, ["Amount"]),
    issueDate: findCol(headers, ["Issue Date"]),
    dueDate: findCol(headers, ["Due Date"]),
    status: findCol(headers, ["Status"]),
    amountPaid: findCol(headers, ["Amount Paid"])
  };
}

// ---- Overdue invoices (60+ days, unpaid) ----
// fromDate/toDate are optional: pass them to restrict results to invoices whose
// Due Date falls inside [from, to] (inclusive). Omit them for the unfiltered list.
function getOverdueInvoices(fromDate, toDate) {
  const { headers, rows } = getTab_("Invoices");
  const c = invoiceCols_(headers);

  if (!c.invNumber || !c.amount || !c.dueDate) {
    return errorResponse("Invoices tab: required columns nahi mile");
  }

  const today = new Date();
  const invoices = [];

  rows.forEach((row) => {
    const invNumber = row[c.invNumber - 1];
    if (!invNumber || String(invNumber).trim() === "") return;

    const amount = Number(row[c.amount - 1]) || 0;
    const paid = c.amountPaid ? (Number(row[c.amountPaid - 1]) || 0) : 0;
    const outstanding = amount - paid;
    const status = c.status ? String(row[c.status - 1] || "").trim().toLowerCase() : "";
    const dueDate = toDate_(row[c.dueDate - 1]);

    if (outstanding <= 0 || status === "paid") return;
    if (!dueDate) return;

    // Global date-range filter — same guard style as the overdueDays filter above
    if (!inRange_(dueDate, fromDate, toDate)) return;

    const overdueDays = Math.floor((today - dueDate) / 86400000);
    if (overdueDays >= 60) {
      invoices.push({
        invNumber: String(invNumber).trim(),
        client: c.client ? String(row[c.client - 1] || "Unknown").trim() : "Unknown",
        company: c.company ? String(row[c.company - 1] || "Tamatos").trim() : "Tamatos",
        overdueDays: overdueDays,
        amount: outstanding,
        currency: c.currency ? String(row[c.currency - 1] || "PKR").trim() : "PKR",
        dueDate: formatDate(dueDate),
        dueDateRaw: dueDate.toISOString()
      });
    }
  });

  invoices.sort((a, b) => b.overdueDays - a.overdueDays);
  const meta = rangeMeta_(fromDate, toDate);
  return jsonResponse({
    success: true, count: invoices.length, invoices: invoices,
    rangeApplied: meta.rangeApplied, from: meta.from, to: meta.to,
    lastFetched: new Date().toISOString()
  });
}

// ---- Revenue totals (metrics + charts) ----
function getRevenueSummary() {
  const { headers, rows } = getTab_("Invoices");
  const c = invoiceCols_(headers);

  if (!c.amount) return errorResponse("Invoices tab: Amount column nahi mila");

  const monthly = {};
  const byCompany = {};
  const today = new Date();
  const d30 = new Date(today.getTime() - 30 * 86400000);
  const d60 = new Date(today.getTime() - 60 * 86400000);
  let last30Days = 0, prior30Days = 0;

  rows.forEach((row) => {
    const amount = Number(row[c.amount - 1]) || 0;
    if (!amount) return;

    const issueDate = c.issueDate ? toDate_(row[c.issueDate - 1]) : null;
    const company = c.company ? String(row[c.company - 1] || "Tamatos").trim() : "Tamatos";

    if (issueDate) {
      const key = issueDate.getFullYear() + "-" + String(issueDate.getMonth() + 1).padStart(2, "0");
      monthly[key] = (monthly[key] || 0) + amount;

      if (issueDate >= d30 && issueDate <= today) last30Days += amount;
      else if (issueDate >= d60 && issueDate < d30) prior30Days += amount;
    }

    byCompany[company] = (byCompany[company] || 0) + amount;
  });

  const pctChange = prior30Days > 0 ? Math.round(((last30Days - prior30Days) / prior30Days) * 100) : null;

  return jsonResponse({
    success: true, monthly: monthly, byCompany: byCompany,
    last30Days: last30Days, prior30Days: prior30Days, pctChange: pctChange,
    lastFetched: new Date().toISOString()
  });
}

// ---- AR aging buckets + Retainer/Project breakdown ----
// fromDate/toDate are optional: pass them to include only invoices whose Due Date
// falls inside [from, to] (inclusive). Omit them for the unfiltered calculation.
function getBreakdown(fromDate, toDate) {
  const { headers, rows } = getTab_("Invoices");
  const c = invoiceCols_(headers);

  if (!c.amount || !c.dueDate) return errorResponse("Invoices tab: required columns nahi mile");

  const today = new Date();
  const aging = { "0-30": 0, "31-60": 0, "61-90": 0, "90+": 0 };
  const typeBreakdown = {};
  const rangeActive = !!(fromDate || toDate);

  rows.forEach((row) => {
    const amount = Number(row[c.amount - 1]) || 0;
    const paid = c.amountPaid ? (Number(row[c.amountPaid - 1]) || 0) : 0;
    const outstanding = amount - paid;
    const status = c.status ? String(row[c.status - 1] || "").trim().toLowerCase() : "";
    const dueDate = toDate_(row[c.dueDate - 1]);

    // With no range active every row contributes, exactly as before.
    if (rangeActive && (!dueDate || !inRange_(dueDate, fromDate, toDate))) return;

    if (outstanding > 0 && status !== "paid" && dueDate) {
      const overdueDays = Math.max(0, Math.floor((today - dueDate) / 86400000));
      if (overdueDays <= 30) aging["0-30"] += outstanding;
      else if (overdueDays <= 60) aging["31-60"] += outstanding;
      else if (overdueDays <= 90) aging["61-90"] += outstanding;
      else aging["90+"] += outstanding;
    }

    if (c.type && amount > 0) {
      const type = String(row[c.type - 1] || "Other").trim() || "Other";
      typeBreakdown[type] = (typeBreakdown[type] || 0) + amount;
    }
  });

  const meta = rangeMeta_(fromDate, toDate);
  return jsonResponse({
    success: true, aging: aging, typeBreakdown: typeBreakdown,
    rangeApplied: meta.rangeApplied, from: meta.from, to: meta.to,
    lastFetched: new Date().toISOString()
  });
}

// ============================================
// EXPENSES TAB — upcoming payments
// ============================================

function getUpcomingPayments() {
  const { headers, rows } = getTab_("Expenses");

  const c = {
    vendor: findCol(headers, ["Vendor"]),
    company: findCol(headers, ["Company"]),
    category: findCol(headers, ["Category"]),
    currency: findCol(headers, ["Currency"]),
    amount: findCol(headers, ["Amount"]),
    dueDate: findCol(headers, ["Due Date"]),
    status: findCol(headers, ["Status"]),
    account: findCol(headers, ["Account"])
  };

  if (!c.amount || !c.dueDate) return errorResponse("Expenses tab: required columns nahi mile");

  const payments = [];

  rows.forEach((row) => {
    const status = c.status ? String(row[c.status - 1] || "").trim().toLowerCase() : "";
    if (status === "paid") return;

    const amount = Number(row[c.amount - 1]) || 0;
    if (!amount) return;

    const dueDate = toDate_(row[c.dueDate - 1]);

    payments.push({
      vendor: c.vendor ? String(row[c.vendor - 1] || "Unknown").trim() : "Unknown",
      company: c.company ? String(row[c.company - 1] || "Tamatos").trim() : "Tamatos",
      category: c.category ? String(row[c.category - 1] || "").trim() : "",
      account: c.account ? String(row[c.account - 1] || "").trim() : "",
      amount: amount,
      currency: c.currency ? String(row[c.currency - 1] || "PKR").trim() : "PKR",
      dueDate: dueDate ? formatDate(dueDate) : "N/A",
      dueDateRaw: dueDate ? dueDate.toISOString() : null
    });
  });

  payments.sort((a, b) => {
    if (!a.dueDateRaw) return 1;
    if (!b.dueDateRaw) return -1;
    return new Date(a.dueDateRaw) - new Date(b.dueDateRaw);
  });

  return jsonResponse({ success: true, count: payments.length, payments: payments, lastFetched: new Date().toISOString() });
}

// ---- Expense totals (paid + due, monthly + last 30 days) — for Net profit, P&L, margin ----
function getExpenseSummary() {
  const { headers, rows } = getTab_("Expenses");

  const c = {
    amount: findCol(headers, ["Amount"]),
    dueDate: findCol(headers, ["Due Date"]),
    paymentDate: findCol(headers, ["Payment Date"]),
    status: findCol(headers, ["Status"]),
    company: findCol(headers, ["Company"])
  };

  if (!c.amount) return errorResponse("Expenses tab: Amount column nahi mila");

  const monthly = {};   // "YYYY-MM" -> total expense (cash-basis: Payment Date if paid, else Due Date)
  const byCompany = {};
  const today = new Date();
  const d30 = new Date(today.getTime() - 30 * 86400000);
  const d60 = new Date(today.getTime() - 60 * 86400000);
  let last30Days = 0, prior30Days = 0;

  rows.forEach((row) => {
    const amount = Number(row[c.amount - 1]) || 0;
    if (!amount) return;

    const status = c.status ? String(row[c.status - 1] || "").trim().toLowerCase() : "";
    const paymentDate = c.paymentDate ? toDate_(row[c.paymentDate - 1]) : null;

    // Cash-basis: use actual Payment Date if this expense has been paid; otherwise skip
    // (an unpaid bill isn't a real cash outflow yet — it belongs in "upcoming payments", not spend-to-date).
    const effectiveDate = (status === "paid" && paymentDate) ? paymentDate : null;
    if (!effectiveDate) return;

    const company = c.company ? String(row[c.company - 1] || "Tamatos").trim() : "Tamatos";
    const key = effectiveDate.getFullYear() + "-" + String(effectiveDate.getMonth() + 1).padStart(2, "0");
    monthly[key] = (monthly[key] || 0) + amount;
    byCompany[company] = (byCompany[company] || 0) + amount;

    if (effectiveDate >= d30 && effectiveDate <= today) last30Days += amount;
    else if (effectiveDate >= d60 && effectiveDate < d30) prior30Days += amount;
  });

  return jsonResponse({
    success: true, monthly: monthly, byCompany: byCompany,
    last30Days: last30Days, prior30Days: prior30Days,
    lastFetched: new Date().toISOString()
  });
}

// ============================================
// BANK_TRANSACTIONS TAB — cash position + cash flow
// ============================================

function getAccounts() {
  const { headers, rows } = getTab_("Bank_Transactions");

  const c = {
    date: findCol(headers, ["Date"]),
    company: findCol(headers, ["Company"]),
    account: findCol(headers, ["Account"]),
    currency: findCol(headers, ["Currency"]),
    type: findCol(headers, ["Type"]),
    amount: findCol(headers, ["Amount"])
  };

  if (!c.account || !c.amount || !c.type) return errorResponse("Bank_Transactions tab: required columns nahi mile");

  const balances = {}; // "Account|Currency" -> balance
  const monthly = {};  // "YYYY-MM" -> { in, out }
  let totalCashPKR = 0;

  rows.forEach((row) => {
    const account = String(row[c.account - 1] || "Unknown").trim();
    const currency = c.currency ? String(row[c.currency - 1] || "PKR").trim() : "PKR";
    const type = String(row[c.type - 1] || "").trim().toLowerCase();
    const amount = Number(row[c.amount - 1]) || 0;
    if (!amount) return;

    const signedAmount = type === "in" ? amount : -amount;
    const key = account + "|" + currency;
    balances[key] = (balances[key] || 0) + signedAmount;

    if (currency === "PKR") totalCashPKR += signedAmount;

    const date = c.date ? toDate_(row[c.date - 1]) : null;
    if (date) {
      const mKey = date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
      if (!monthly[mKey]) monthly[mKey] = { in: 0, out: 0 };
      if (type === "in") monthly[mKey].in += amount;
      else monthly[mKey].out += amount;
    }
  });

  const accounts = Object.keys(balances).map((key) => {
    const [account, currency] = key.split("|");
    return { account: account, currency: currency, balance: balances[key] };
  });

  return jsonResponse({
    success: true, accounts: accounts, totalCashPKR: totalCashPKR, monthly: monthly,
    lastFetched: new Date().toISOString()
  });
}

// ---- New spend flagged: Pending expenses over a fixed threshold ----
const FLAGGED_SPEND_THRESHOLD_PKR = 100000;

function getFlaggedSpend() {
  const { headers, rows } = getTab_("Expenses");

  const c = {
    expNumber: findCol(headers, ["Bill / Expense #", "Bill/Expense #", "Expense #"]),
    vendor: findCol(headers, ["Vendor"]),
    company: findCol(headers, ["Company"]),
    category: findCol(headers, ["Category"]),
    currency: findCol(headers, ["Currency"]),
    amount: findCol(headers, ["Amount"]),
    dueDate: findCol(headers, ["Due Date"]),
    status: findCol(headers, ["Status"])
  };

  if (!c.amount || !c.status) return errorResponse("Expenses tab: required columns nahi mile");

  const flagged = [];

  rows.forEach((row) => {
    const status = String(row[c.status - 1] || "").trim().toLowerCase();
    if (status !== "pending") return;

    const amount = Number(row[c.amount - 1]) || 0;
    if (amount < FLAGGED_SPEND_THRESHOLD_PKR) return;

    const dueDate = c.dueDate ? toDate_(row[c.dueDate - 1]) : null;

    flagged.push({
      expNumber: c.expNumber ? String(row[c.expNumber - 1] || "").trim() : "",
      vendor: c.vendor ? String(row[c.vendor - 1] || "Unknown").trim() : "Unknown",
      company: c.company ? String(row[c.company - 1] || "Tamatos").trim() : "Tamatos",
      category: c.category ? String(row[c.category - 1] || "").trim() : "",
      amount: amount,
      currency: c.currency ? String(row[c.currency - 1] || "PKR").trim() : "PKR",
      dueDate: dueDate ? formatDate(dueDate) : "N/A"
    });
  });

  flagged.sort((a, b) => b.amount - a.amount);

  return jsonResponse({
    success: true, threshold: FLAGGED_SPEND_THRESHOLD_PKR,
    count: flagged.length, flagged: flagged,
    lastFetched: new Date().toISOString()
  });
}

// ============================================
// DEBUG — inspect raw headers + first row's value types
// ============================================

function getDebugInfo() {
  const tabs = ["Invoices", "Expenses", "Bank_Transactions"];
  const result = {};

  tabs.forEach((tabName) => {
    try {
      const { headers, rows } = getTab_(tabName);
      const firstRow = rows.length > 0 ? rows[0] : [];
      result[tabName] = {
        headers: headers,
        firstRowValues: firstRow.map((val, i) => ({
          header: headers[i],
          value: val,
          jsType: typeof val,
          isDateObject: val instanceof Date,
          stringified: String(val)
        }))
      };
    } catch (err) {
      result[tabName] = { error: err.toString() };
    }
  });

  return jsonResponse({ success: true, tabs: result, lastFetched: new Date().toISOString() });
}

// ============================================
// COMBINED
// ============================================

function getCombined(fromDate, toDate) {
  const overdue = JSON.parse(getOverdueInvoices(fromDate, toDate).getContent());
  const revenue = JSON.parse(getRevenueSummary().getContent());
  const breakdown = JSON.parse(getBreakdown(fromDate, toDate).getContent());
  const expenses = JSON.parse(getUpcomingPayments().getContent());
  const expenseSummary = JSON.parse(getExpenseSummary().getContent());
  const accounts = JSON.parse(getAccounts().getContent());
  const flagged = JSON.parse(getFlaggedSpend().getContent());

  return jsonResponse({
    success: true,
    overdue: overdue, revenue: revenue, breakdown: breakdown,
    expenses: expenses, expenseSummary: expenseSummary, accounts: accounts,
    flagged: flagged,
    lastFetched: new Date().toISOString()
  });
}
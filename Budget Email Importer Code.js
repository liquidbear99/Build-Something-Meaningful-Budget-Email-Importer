/*******************************************************
 * Budget Email Transaction Importer + OpenWebUI AI
 * Simplified vNext Lite + American Express Support
 *
 * Core sheets only:
 * - Expenses
 * - Pending Transactions
 * - Merchant Rules
 * - Ignored Merchants
 * - Review Settings
 * - Import Log
 *
 * Decision tree:
 * 1. Parse transaction
 * 2. Ignore merchant -> stop
 * 3. Duplicate -> stop
 * 4. Merchant rule confident + auto approve -> push to Expenses
 * 5. Otherwise use AI when enabled
 * 6. AI confident and under uncertainty threshold -> push to Expenses
 * 7. Otherwise send to Pending Review
 *******************************************************/

const CONFIG = {
  PENDING_SHEET_NAME: "Pending Transactions",
  EXPENSE_SHEET_NAME: "Expenses",
  MERCHANT_RULES_SHEET_NAME: "Merchant Rules",
  IGNORED_MERCHANTS_SHEET_NAME: "Ignored Merchants",
  REVIEW_SETTINGS_SHEET_NAME: "Review Settings",
  IMPORT_LOG_SHEET_NAME: "Import Log",

  PROCESSED_LABEL: "Budget/Processed",
  NEEDS_REVIEW_LABEL: "Budget/Needs Review",

  SEARCH_DAYS_BACK: 7,
  MAX_THREADS_PER_RUN: 50,

  // Example OpenWebUI URL: https://your-domain.example.com/api/chat/completions
  OPENWEBUI_API_URL: "PASTE_YOUR_OPENWEBUI_CHAT_COMPLETIONS_URL_HERE",
  OPENWEBUI_API_KEY: "PASTE_YOUR_OPENWEBUI_API_KEY_HERE",
  // Keep AUTO unless you want to force an exact ID returned by /api/models.
  OPENWEBUI_MODEL: "AUTO",

  DEFAULT_RULE_CONFIDENCE: 0.95,
  DEFAULT_AUTO_APPROVE_UNCERTAINTY: 0.10,

  ALLOWED_CATEGORIES: [
    "Groceries",
    "Restaurants",
    "Misc",
    "My Fun",
    "Spouse Fun",
    "Pet",
    "Hair/Cosmetics",
    "Home/Car Maintenance",
    "Entertainment/Clothing",
    "Medical",
    "Fuel",
    "Utilities",
    "Spouse Benefits",
    "My Benefits",
    "Vacation"
  ],
  // Paste the Google Form edit URL here only if PUSH_TO_GOOGLE_FORM is true.
  GOOGLE_FORM_URL: "PASTE_YOUR_GOOGLE_FORM_EDIT_URL_HERE",

PUSH_TO_GOOGLE_FORM: false,

SEND_REVIEW_EMAILS: false,
NOTIFICATION_EMAIL: "",
WEB_APP_URL: "",

FORM_CATEGORY_MAP: {
  "Groceries": "Groceries",
  "Restaurants": "Restaurants",
  "Misc": "Misc",
  "My Fun": "My Fun",
  "Spouse Fun": "Spouse Fun",
  "Pet": "Pet",
  "Hair/Cosmetics": "Hair/Cosmetics",
  "Home/Car Maintenance": "Home/Car Maintenance",
  "Entertainment/Clothing": "Entertainment/Clothing",
  "Medical": "Medical",
  "Fuel": "Fuel",
  "Utilities": "Utilities",
  "Spouse Benefits": "Spouse Benefits",
  "My Benefits": "My Benefits",
  "Vacation": "Vacation"
},
};

const PENDING_HEADERS = [
  "Date",
  "Merchant",
  "Amount",
  "Selected Category",
  "Rule Category",
  "Rule Confidence",
  "AI Category",
  "AI Confidence",
  "AI Reason",
  "Review Reason",
  "Status",
  "Duplicate Key",
  "Imported At",
  "Account",
  "Card Last 4",
  "Source",
  "Gmail Message ID"
];

const EXPENSE_HEADERS = [
  "Date",
  "Merchant",
  "Amount",
  "Category",
  "Account",
  "Card Last 4",
  "Source",
  "Duplicate Key",
  "Imported At",
  "Approval Source"
];

const RULE_HEADERS = [
  "Pattern",
  "Category",
  "Confidence",
  "Auto Approve",
  "Learning Allowed",
  "Notes",
  "Enabled"
];

const IGNORE_HEADERS = ["Pattern", "Notes", "Enabled"];
const SETTINGS_HEADERS = ["Setting", "Value", "Notes"];
const IMPORT_LOG_HEADERS = ["Timestamp", "Action", "Merchant", "Amount", "Category", "Duplicate Key", "Notes"];

/********************************************************
 * MENU + SETUP
 ********************************************************/

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Budget Importer")
      .addItem("1. Run Setup / Repair Sheets", "setupBudgetAutomation")
      .addSeparator()
      .addItem("2. Import Gmail Transactions", "importGmailTransactions")
      .addItem("3. Remove Ignored Pending Transactions", "removeIgnoredPendingTransactions")
      .addSeparator()
      .addItem("Test Chase Parser", "testParseExample")
      .addItem("Test Amex Parser", "testParseAmexExample")
      .addItem("Test Amex Statement Exclusion", "testParseAmexStatementExclusion")
      .addItem("Test OpenWebUI Categorizer", "testOpenWebUiCategorizerFromAppsScript")
      .addItem("Add Fake Pending Review Transaction", "addFakePendingReviewTransaction")
      .addSeparator()
      .addItem("Create Hourly Import Trigger", "createHourlyImportTrigger")
      .addItem("Delete Budget Triggers", "deleteBudgetTriggers")
      .addToUi();
  } catch (err) {
    console.log("Menu creation skipped: " + err.message);
  }
}

function setupBudgetAutomation() {
  const ss = getActiveSpreadsheetOrThrow_();

  setupPendingSheet_(getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME));
  setupExpenseSheet_(getOrCreateSheet_(ss, CONFIG.EXPENSE_SHEET_NAME));
  setupMerchantRulesSheet_(getOrCreateSheet_(ss, CONFIG.MERCHANT_RULES_SHEET_NAME));
  setupIgnoredMerchantsSheet_(getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME));
  setupReviewSettingsSheet_(getOrCreateSheet_(ss, CONFIG.REVIEW_SETTINGS_SHEET_NAME));
  setupImportLogSheet_(getOrCreateSheet_(ss, CONFIG.IMPORT_LOG_SHEET_NAME));

  getOrCreateGmailLabel_(CONFIG.PROCESSED_LABEL);
  getOrCreateGmailLabel_(CONFIG.NEEDS_REVIEW_LABEL);

  showMessage_(
    "Setup complete.\n\n" +
    "This starter version uses these support sheets:\n" +
    "Pending Transactions, Expenses, Merchant Rules, Ignored Merchants, Review Settings, and Import Log.\n\n" +
    "Next steps:\n" +
    "1. Paste your OpenWebUI URL/API key in CONFIG if using AI.\n" +
    "2. Paste your Google Form URL and set PUSH_TO_GOOGLE_FORM to true if using a form.\n" +
    "3. Deploy the web app, paste its URL into CONFIG.WEB_APP_URL, and optionally enable review emails."
  );
}

/********************************************************
 * WEB APP
 ********************************************************/

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("ReviewApp")
    .setTitle("Budget Review")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPendingReviewItems() {
  const ss = getActiveSpreadsheetOrThrow_();
  const sheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  setupPendingSheet_(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return { ok: true, items: [], categories: CONFIG.ALLOWED_CATEGORIES };
  }

  const values = sheet.getRange(2, 1, lastRow - 1, PENDING_HEADERS.length).getValues();
  const items = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const status = String(row[10] || "").trim();

    if (status !== "Pending Review") continue;
    if (!row[1] || row[2] === "") continue;

    const selectedCategory = sanitizeCategory_(row[3]);
    const aiCategory = sanitizeCategory_(row[6]);

    items.push({
      rowNumber: i + 2,
      transactionDate: formatDateForWeb_(row[0]),
      merchant: String(row[1] || ""),
      amount: Number(row[2] || 0),

      // New/legacy aliases so either ReviewApp version can read this safely.
      category: selectedCategory,
      selectedCategory: selectedCategory,
      ruleCategory: sanitizeCategory_(row[4]),
      ruleConfidence: clampNumber_(Number(row[5] || 0), 0, 1),
      aiSuggestion: aiCategory,
      aiCategory: aiCategory,
      aiConfidence: clampNumber_(Number(row[7] || 0), 0, 1),
      aiReason: String(row[8] || ""),
      reviewReason: String(row[9] || "Needs review."),
      needsReview: true,
      duplicateKey: String(row[11] || ""),
      aiAvailable: Boolean(row[6] && aiCategory !== "Misc")
    });
  }

  return { ok: true, items, categories: CONFIG.ALLOWED_CATEGORIES };
}

/********************************************************
 * REVIEW ACTIONS
 ********************************************************/

function approveSelectedByKey(duplicateKey, category) {
  return finalizePendingTransactionByKey_(duplicateKey, category, {
    approvalSource: "Manual",
    remember: false
  });
}

function approveSelectedRememberByKey(duplicateKey, category) {
  return finalizePendingTransactionByKey_(duplicateKey, category, {
    approvalSource: "Manual + Learned",
    remember: true
  });
}

function approveAiByKey(duplicateKey, aiCategoryFromClient) {
  const cleanAiCategory = validateCategory_(aiCategoryFromClient, "AI category");
  if (cleanAiCategory === "Misc") {
    throw new Error("AI category is Misc. Choose a category manually and use Approve Selected.");
  }

  return finalizePendingTransactionByKey_(duplicateKey, cleanAiCategory, {
    approvalSource: "AI Manual Approval",
    remember: false
  });
}

function approveAiRememberByKey(duplicateKey, aiCategoryFromClient) {
  const cleanAiCategory = validateCategory_(aiCategoryFromClient, "AI category");
  if (cleanAiCategory === "Misc") {
    throw new Error("AI category is Misc. Choose a category manually and use Approve Selected & Remember.");
  }

  return finalizePendingTransactionByKey_(duplicateKey, cleanAiCategory, {
    approvalSource: "AI + Learned",
    remember: true
  });
}

function skipPendingItemByKey(duplicateKey) {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  setupPendingSheet_(pendingSheet);

  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction.");

  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const merchant = String(row[1] || "");
  const amount = Number(row[2] || 0);

  // Skip means this transaction only. Delete it from the pending queue so it
  // does not keep cluttering the review sheet.
  pendingSheet.deleteRow(rowNumber);
  logImport_("Skipped", merchant, amount, "", duplicateKey, "Deleted from pending review app.");

  return { ok: true, message: merchant ? "Skipped: " + merchant + "." : "Skipped." };
}


/**
 * Row-number compatibility wrappers for ReviewApp.html.
 * The snappy card UI uses processReviewActions(), but these keep older
 * deployed HTML versions working too.
 */
function approveCurrentSelectionAndPush(rowNumber, category) {
  const key = getDuplicateKeyForPendingRow_(rowNumber);
  return approveSelectedByKey(key, category);
}

function approveCurrentSelectionLearnAndPush(rowNumber, category) {
  const key = getDuplicateKeyForPendingRow_(rowNumber);
  return approveSelectedRememberByKey(key, category);
}

function approveAiSuggestedPendingItem(rowNumber) {
  const tx = getPendingRowSummary_(rowNumber);
  return approveAiByKey(tx.duplicateKey, tx.aiCategory);
}

function approveAiSuggestedLearnAndPush(rowNumber) {
  const tx = getPendingRowSummary_(rowNumber);
  return approveAiRememberByKey(tx.duplicateKey, tx.aiCategory);
}

function skipPendingItem(rowNumber) {
  const key = getDuplicateKeyForPendingRow_(rowNumber);
  return skipPendingItemByKey(key);
}

function ignoreMerchantPendingItem(rowNumber) {
  const key = getDuplicateKeyForPendingRow_(rowNumber);
  return ignoreMerchantByKey(key);
}

function getDuplicateKeyForPendingRow_(rowNumber) {
  return getPendingRowSummary_(rowNumber).duplicateKey;
}

function getPendingRowSummary_(rowNumber) {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  setupPendingSheet_(pendingSheet);

  const rowNum = Number(rowNumber);
  if (!rowNum || rowNum < 2 || rowNum > pendingSheet.getLastRow()) {
    throw new Error("Invalid pending row number: " + rowNumber);
  }

  const row = pendingSheet.getRange(rowNum, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const duplicateKey = String(row[11] || "").trim();
  if (!duplicateKey) {
    throw new Error("Could not process row " + rowNumber + ": duplicate key is missing.");
  }

  return {
    duplicateKey: duplicateKey,
    aiCategory: sanitizeCategory_(row[6]),
    selectedCategory: sanitizeCategory_(row[3])
  };
}

function ignoreMerchantByKey(duplicateKey) {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  const ignoredSheet = getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME);
  setupPendingSheet_(pendingSheet);
  setupIgnoredMerchantsSheet_(ignoredSheet);

  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction.");

  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const merchant = String(row[1] || "").trim();
  const amount = Number(row[2] || 0);

  if (!merchant) throw new Error("Merchant is blank.");

  addIgnoredMerchantRule_(merchant, "Added from review app.");
  pendingSheet.deleteRow(rowNumber);
  logImport_("Ignored Merchant", merchant, amount, "", duplicateKey, "Added ignored merchant rule and deleted pending row.");

  return { ok: true, message: "Ignored merchant going forward: " + merchant + "." };
}

function finalizePendingTransactionByKey_(duplicateKey, category, options) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Another review action is already running.");
  }

  try {
    return finalizePendingTransactionByKeyNoLock_(duplicateKey, category, options || {});
  } finally {
    lock.releaseLock();
  }
}

function finalizePendingTransactionByKeyNoLock_(duplicateKey, category, options) {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  const expenseSheet = getOrCreateSheet_(ss, CONFIG.EXPENSE_SHEET_NAME);

  setupPendingSheet_(pendingSheet);
  setupExpenseSheet_(expenseSheet);

  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction.");

  const cleanCategory = validateCategory_(category, "Category");
  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];

  const tx = {
    transaction_date: row[0],
    merchant: String(row[1] || ""),
    amount: Number(row[2] || 0),
    category: cleanCategory,
    account: String(row[13] || ""),
    card_last4: String(row[14] || ""),
    source: String(row[15] || ""),
    duplicate_key: String(row[11] || duplicateKey || ""),
    gmail_message_id: String(row[16] || "")
  };

  if (!tx.transaction_date || !tx.merchant || tx.amount === "" || !tx.duplicate_key) {
    throw new Error("Pending transaction is missing required fields.");
  }

  if (expenseDuplicateExists_(expenseSheet, tx.duplicate_key)) {
    pendingSheet.getRange(rowNumber, 11).setValue("Duplicate");
    logImport_("Duplicate At Approval", tx.merchant, tx.amount, cleanCategory, tx.duplicate_key, "Already exists in Expenses.");
    return { ok: true, message: "Already existed in Expenses. Marked duplicate." };
  }

  appendExpenseRow_(expenseSheet, tx, options.approvalSource || "Manual");
  pendingSheet.getRange(rowNumber, 4).setValue(cleanCategory);
  pendingSheet.getRange(rowNumber, 11).setValue("Moved");

  if (options.remember) {
    addMerchantRule_(tx.merchant, cleanCategory, {
      confidence: CONFIG.DEFAULT_RULE_CONFIDENCE,
      autoApprove: true,
      learningAllowed: true,
      notes: "Learned from review app: " + (options.approvalSource || "Manual")
    });
  }

  logImport_(options.approvalSource || "Manual", tx.merchant, tx.amount, cleanCategory, tx.duplicate_key, "Approved from review app.");

  return {
    ok: true,
    message: "Approved as " + cleanCategory + "."
  };
}



/**
 * Batch review actions for optimistic UI.
 * Browser removes cards immediately, queues decisions locally, and sends them here in batches.
 */
function processReviewActions(actions) {
  if (!Array.isArray(actions) || actions.length === 0) {
    return { ok: true, results: [] };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Another review batch is already running.");
  }

  const results = [];

  try {
    const ss = getActiveSpreadsheetOrThrow_();
    const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
    const expenseSheet = getOrCreateSheet_(ss, CONFIG.EXPENSE_SHEET_NAME);
    const ignoredSheet = getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME);

    setupPendingSheet_(pendingSheet);
    setupExpenseSheet_(expenseSheet);
    setupIgnoredMerchantsSheet_(ignoredSheet);

    for (const rawAction of actions) {
      const action = rawAction || {};
      const key = String(action.duplicateKey || "").trim();
      const type = String(action.action || "").trim();

      try {
        if (!key) throw new Error("Missing duplicate key.");

        let result;
        if (type === "approve-selected") {
          result = finalizePendingTransactionFast_(pendingSheet, expenseSheet, key, action.category, { approvalSource: "Manual", remember: false });
        } else if (type === "approve-selected-remember") {
          result = finalizePendingTransactionFast_(pendingSheet, expenseSheet, key, action.category, { approvalSource: "Manual + Learned", remember: true });
        } else if (type === "approve-ai") {
          const aiCategory = validateCategory_(action.aiCategory, "AI category");
          if (aiCategory === "Misc") throw new Error("AI category is Misc. Choose a category manually and use Approve Selected.");
          result = finalizePendingTransactionFast_(pendingSheet, expenseSheet, key, aiCategory, { approvalSource: "AI Manual Approval", remember: false });
        } else if (type === "approve-ai-remember") {
          const aiCategory = validateCategory_(action.aiCategory, "AI category");
          if (aiCategory === "Misc") throw new Error("AI category is Misc. Choose a category manually and use Approve Selected & Remember.");
          result = finalizePendingTransactionFast_(pendingSheet, expenseSheet, key, aiCategory, { approvalSource: "AI + Learned", remember: true });
        } else if (type === "skip") {
          result = skipPendingItemFast_(pendingSheet, key);
        } else if (type === "ignore") {
          result = ignoreMerchantFast_(pendingSheet, ignoredSheet, key);
        } else {
          throw new Error("Unknown review action: " + type);
        }

        results.push({ ok: true, duplicateKey: key, action: type, message: result.message || "Done." });
      } catch (err) {
        results.push({ ok: false, duplicateKey: key, action: type, message: err.message });
      }
    }

    SpreadsheetApp.flush();
    return { ok: true, results };
  } finally {
    lock.releaseLock();
  }
}

function finalizePendingTransactionFast_(pendingSheet, expenseSheet, duplicateKey, category, options) {
  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction. It may already have been processed.");

  const cleanCategory = validateCategory_(category, "Category");
  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];

  const tx = {
    transaction_date: row[0],
    merchant: String(row[1] || ""),
    amount: Number(row[2] || 0),
    category: cleanCategory,
    account: String(row[13] || ""),
    card_last4: String(row[14] || ""),
    source: String(row[15] || ""),
    duplicate_key: String(row[11] || duplicateKey || ""),
    gmail_message_id: String(row[16] || "")
  };

  if (!tx.transaction_date || !tx.merchant || tx.amount === "" || !tx.duplicate_key) {
    throw new Error("Pending transaction is missing required fields.");
  }

  if (expenseDuplicateExists_(expenseSheet, tx.duplicate_key)) {
    pendingSheet.getRange(rowNumber, 11).setValue("Duplicate");
    logImport_("Duplicate At Approval", tx.merchant, tx.amount, cleanCategory, tx.duplicate_key, "Already exists in Expenses.");
    return { ok: true, message: "Already existed in Expenses. Marked duplicate." };
  }

  appendExpenseRow_(expenseSheet, tx, options.approvalSource || "Manual");
  pendingSheet.getRange(rowNumber, 4).setValue(cleanCategory);
  pendingSheet.getRange(rowNumber, 11).setValue("Moved");

  if (options.remember) {
    addMerchantRule_(tx.merchant, cleanCategory, {
      confidence: CONFIG.DEFAULT_RULE_CONFIDENCE,
      autoApprove: true,
      learningAllowed: true,
      notes: "Learned from review app: " + (options.approvalSource || "Manual")
    });
  }

  logImport_(options.approvalSource || "Manual", tx.merchant, tx.amount, cleanCategory, tx.duplicate_key, "Approved from review app batch.");
  return { ok: true, message: "Approved as " + cleanCategory + "." };
}

function skipPendingItemFast_(pendingSheet, duplicateKey) {
  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction. It may already have been processed.");

  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const merchant = String(row[1] || "");
  const amount = Number(row[2] || 0);

  // Skip means this transaction only. Delete it from pending review.
  pendingSheet.deleteRow(rowNumber);
  logImport_("Skipped", merchant, amount, "", duplicateKey, "Deleted from optimistic review batch.");

  return { ok: true, message: merchant ? "Skipped: " + merchant + "." : "Skipped." };
}

function ignoreMerchantFast_(pendingSheet, ignoredSheet, duplicateKey) {
  const rowNumber = findPendingRowByDuplicateKey_(pendingSheet, duplicateKey);
  if (!rowNumber) throw new Error("Could not find pending transaction. It may already have been processed.");

  const row = pendingSheet.getRange(rowNumber, 1, 1, PENDING_HEADERS.length).getValues()[0];
  const merchant = String(row[1] || "").trim();
  const amount = Number(row[2] || 0);

  if (!merchant) throw new Error("Merchant is blank.");

  addIgnoredMerchantRule_(merchant, "Added from review app.");

  // Ignore Merchant means future matching transactions should be excluded.
  // Delete the current pending row after creating the ignore rule.
  pendingSheet.deleteRow(rowNumber);
  logImport_("Ignored Merchant", merchant, amount, "", duplicateKey, "Added ignored merchant rule and deleted pending row from optimistic review batch.");

  return { ok: true, message: "Ignored merchant going forward: " + merchant + "." };
}

/********************************************************
 * IMPORT DECISION ENGINE
 ********************************************************/

function importGmailTransactions() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    console.log("Another import run is already in progress. Exiting.");
    return;
  }

  try {
    const ss = getActiveSpreadsheetOrThrow_();

    const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
    const expenseSheet = getOrCreateSheet_(ss, CONFIG.EXPENSE_SHEET_NAME);
    const rulesSheet = getOrCreateSheet_(ss, CONFIG.MERCHANT_RULES_SHEET_NAME);
    const ignoredSheet = getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME);
    const settingsSheet = getOrCreateSheet_(ss, CONFIG.REVIEW_SETTINGS_SHEET_NAME);

    setupPendingSheet_(pendingSheet);
    setupExpenseSheet_(expenseSheet);
    setupMerchantRulesSheet_(rulesSheet);
    setupIgnoredMerchantsSheet_(ignoredSheet);
    setupReviewSettingsSheet_(settingsSheet);
    setupImportLogSheet_(getOrCreateSheet_(ss, CONFIG.IMPORT_LOG_SHEET_NAME));

    const settings = loadReviewSettings_(settingsSheet);
    const rules = loadMerchantRules_(rulesSheet);
    const ignoredRules = loadIgnoredRules_(ignoredSheet);
    const processedLabel = getOrCreateGmailLabel_(CONFIG.PROCESSED_LABEL);
    const reviewLabel = getOrCreateGmailLabel_(CONFIG.NEEDS_REVIEW_LABEL);

    const existingKeys = new Set([
      ...getExistingPendingDuplicateKeys_(pendingSheet),
      ...getExistingExpenseDuplicateKeys_(expenseSheet)
    ]);

    const query = buildGmailSearchQuery_();
    console.log("Gmail search query: " + query);

    const threads = GmailApp.search(query, 0, CONFIG.MAX_THREADS_PER_RUN);

    let importedCount = 0;
    let autoApprovedRuleCount = 0;
    let autoApprovedAiCount = 0;
    let reviewCount = 0;
    let ignoredCount = 0;
    let duplicateCount = 0;
    let noTransactionCount = 0;
    let aiUsedCount = 0;

    const reviewEmailItems = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      let threadHadProcessedMessage = false;
      let threadNeedsReview = false;

      for (const message of messages) {
        const messageId = message.getId();
        const subject = message.getSubject() || "";
        const from = message.getFrom() || "";
        const body = normalizeEmailBody_(message.getPlainBody() || "");

        const tx = parseTransactionEmail_(subject, from, body, messageId);
        if (!tx.is_transaction) {
          noTransactionCount++;
          continue;
        }

        tx.merchant = normalizeMerchantForDisplay_(tx.merchant);
        tx.duplicate_key = createDuplicateKey_(tx);

        if (existingKeys.has(tx.duplicate_key)) {
          duplicateCount++;
          threadHadProcessedMessage = true;
          logImport_("Duplicate Skipped", tx.merchant, tx.amount, "", tx.duplicate_key, "Duplicate key already exists.");
          continue;
        }

        if (shouldIgnoreTransaction_(tx.merchant, ignoredRules)) {
          ignoredCount++;
          threadHadProcessedMessage = true;
          existingKeys.add(tx.duplicate_key);
          logImport_("Ignored", tx.merchant, tx.amount, "", tx.duplicate_key, "Matched Ignored Merchants.");
          continue;
        }

        const ruleMatch = findMerchantRule_(tx.merchant, rules);
        const ruleCategory = ruleMatch ? sanitizeCategory_(ruleMatch.category) : "Misc";
        const ruleConfidence = ruleMatch ? clampNumber_(Number(ruleMatch.confidence || 0), 0, 1) : 0;
        const maxUncertainty = settings.MAX_UNCERTAINTY_FOR_AUTO_APPROVAL;
        const confidenceThreshold = 1 - maxUncertainty;

        tx.rule_category = ruleCategory;
        tx.rule_confidence = ruleConfidence;
        tx.selected_category = ruleCategory;
        tx.ai_category = "";
        tx.ai_confidence = 0;
        tx.ai_reason = "";

        const ruleCanAutoApprove =
          ruleMatch &&
          ruleMatch.autoApprove === true &&
          ruleCategory !== "Misc" &&
          ruleConfidence >= confidenceThreshold;

        if (ruleCanAutoApprove) {
          tx.category = ruleCategory;
          appendExpenseRow_(expenseSheet, tx, "Merchant Rule Auto");
          existingKeys.add(tx.duplicate_key);
          autoApprovedRuleCount++;
          importedCount++;
          threadHadProcessedMessage = true;
          logImport_("Auto Approved Rule", tx.merchant, tx.amount, ruleCategory, tx.duplicate_key, "Matched Merchant Rules.");
          continue;
        }

        let ai = {
          suggested_category: "",
          confidence: 0,
          reason: "",
          needs_review: true
        };

        const shouldUseAi = Boolean(settings.USE_AI_CATEGORIZER) && (
          !settings.AI_ONLY_FOR_UNCERTAIN ||
          !ruleMatch ||
          ruleCategory === "Misc" ||
          ruleConfidence < confidenceThreshold ||
          ruleMatch.autoApprove !== true
        );

        if (shouldUseAi) {
          ai = getAiCategorySuggestion_({
            merchant: tx.merchant,
            amount: tx.amount,
            transaction_date: tx.transaction_date,
            account: tx.account,
            category: ruleCategory
          });
          aiUsedCount++;
          tx.ai_category = sanitizeCategory_(ai.suggested_category);
          tx.ai_confidence = clampNumber_(Number(ai.confidence || 0), 0, 1);
          tx.ai_reason = String(ai.reason || "");
        }

        const aiCanAutoApprove =
          Boolean(settings.AUTO_APPROVE_AI) &&
          tx.ai_category &&
          tx.ai_category !== "Misc" &&
          tx.ai_confidence >= confidenceThreshold &&
          ai.needs_review !== true;

        if (aiCanAutoApprove) {
          tx.category = tx.ai_category;
          tx.selected_category = tx.ai_category;
          appendExpenseRow_(expenseSheet, tx, "AI Auto");
          existingKeys.add(tx.duplicate_key);
          autoApprovedAiCount++;
          importedCount++;
          threadHadProcessedMessage = true;
          logImport_("Auto Approved AI", tx.merchant, tx.amount, tx.ai_category, tx.duplicate_key, "AI confidence " + tx.ai_confidence);
          continue;
        }

        tx.category = tx.ai_category && tx.ai_category !== "Misc" ? tx.ai_category : ruleCategory;
        tx.selected_category = sanitizeCategory_(tx.category);
        tx.review_reason = buildReviewReason_(tx, ruleMatch, ai, settings, tx.parse_notes);
        appendPendingTransactionRow_(pendingSheet, tx);
        existingKeys.add(tx.duplicate_key);
        reviewCount++;
        importedCount++;
        threadHadProcessedMessage = true;
        threadNeedsReview = true;

        reviewEmailItems.push({
          merchant: tx.merchant,
          amount: tx.amount,
          selected_category: tx.selected_category,
          ai_category: tx.ai_category,
          ai_confidence: tx.ai_confidence,
          review_reason: tx.review_reason
        });

        logImport_("Sent To Review", tx.merchant, tx.amount, tx.selected_category, tx.duplicate_key, tx.review_reason);
      }

      if (threadHadProcessedMessage) processedLabel.addToThread(thread);
      if (threadNeedsReview) reviewLabel.addToThread(thread);
    }

    const summary =
      "Import complete.\n\n" +
      "Imported/processed transactions: " + importedCount + "\n" +
      "Auto-approved by rule: " + autoApprovedRuleCount + "\n" +
      "Auto-approved by AI: " + autoApprovedAiCount + "\n" +
      "Needs review: " + reviewCount + "\n" +
      "Ignored: " + ignoredCount + "\n" +
      "Skipped duplicates: " + duplicateCount + "\n" +
      "AI suggestions used: " + aiUsedCount + "\n" +
      "Not transactions: " + noTransactionCount;

    sendReviewNeededEmail_(reviewEmailItems);

    showMessage_(summary);

  } finally {
    lock.releaseLock();
  }
}

function buildReviewReason_(tx, ruleMatch, ai, settings, parseNotes) {
  const threshold = 1 - settings.MAX_UNCERTAINTY_FOR_AUTO_APPROVAL;

  if (parseNotes && /Partial/i.test(parseNotes)) return "Needs review because the email was only partially parsed.";
  if (!ruleMatch && (!tx.ai_category || tx.ai_category === "Misc")) return "Needs review because no confident rule or AI category was found.";
  if (ruleMatch && ruleMatch.learningAllowed === false) return "Needs review because this merchant is marked as ambiguous in Merchant Rules.";
  if (tx.ai_category && tx.ai_category !== "Misc" && tx.ai_confidence < threshold) return "Needs review because AI confidence is " + Math.round(tx.ai_confidence * 100) + "%.";
  if (ruleMatch && tx.rule_confidence < threshold) return "Needs review because rule confidence is " + Math.round(tx.rule_confidence * 100) + "%.";
  if (tx.ai_category === "Misc") return "Needs review because AI could not confidently categorize it.";
  return "Needs review because it did not meet the auto-approval rules.";
}

/********************************************************
 * OPENWEBUI AI CATEGORIZER
 ********************************************************/

function getAiCategorySuggestion_(tx) {
  if (!getBooleanSetting_("USE_AI_CATEGORIZER", true)) {
    return {
      suggested_category: "",
      confidence: 0,
      reason: "AI categorizer is disabled.",
      needs_review: true
    };
  }

  if (
    isPlaceholderConfigValue_(CONFIG.OPENWEBUI_API_URL) ||
    isPlaceholderConfigValue_(CONFIG.OPENWEBUI_API_KEY)
  ) {
    return {
      suggested_category: "",
      confidence: 0,
      reason: "OpenWebUI API URL or key is missing.",
      needs_review: true
    };
  }

  const systemPrompt =
    "You are a strict JSON-only household budget transaction categorizer. " +
    "Return only valid JSON. No markdown. No commentary. Do not invent facts.";

  const userPrompt = `
Categorize this transaction into exactly one allowed category.

Allowed categories:
${JSON.stringify(CONFIG.ALLOWED_CATEGORIES)}

User patterns:
- Gas stations and EV charging -> Fuel.
- Grocery stores -> Groceries.
- Restaurants, coffee shops, bakeries, ice cream, fast food -> Restaurants.
- Utilities, electric, water, sewer, internet, phone -> Utilities.
- Home improvement, car maintenance, car wash, hardware -> Home/Car Maintenance.
- Pet stores, vet, dog daycare/walker -> Pet.
- Hair, cosmetics, salon, nails -> Hair/Cosmetics.
- Pharmacy, clinic, doctor, dental, vision -> Medical.
- Clothing, movies, streaming, entertainment -> Entertainment/Clothing.
- Flights, hotels, airports, lodging, travel activities -> Vacation.
- My Fun and Spouse Fun are personal discretionary categories. Only use them when a merchant rule or clear user-specific clue supports that choice.
- My Benefits and Spouse Benefits are benefits/payroll/reimbursement categories. Only use them when a merchant rule or clear user-specific clue supports that choice.
- Amazon, Costco, Target, Walmart, PayPal, Venmo, Apple, Google are ambiguous unless the merchant text contains a specific clue.

Rules:
- suggested_category must exactly match one allowed category.
- If uncertain, choose Misc and needs_review=true.
- Confidence must be 0.0 to 1.0.
- Use confidence >= 0.90 only when the business type is clear.
- Return JSON only.

Transaction:
Merchant: ${tx.merchant || ""}
Amount: ${tx.amount || ""}
Date: ${tx.transaction_date || ""}
Account: ${tx.account || ""}
Rule category: ${tx.category || "Misc"}

Return exactly this JSON object:
{
  "suggested_category": "Misc",
  "confidence": 0.0,
  "reason": "brief reason",
  "needs_review": true
}
`;

  try {
    const modelCandidates = getOpenWebUiModelCandidates_();

    if (!modelCandidates.length) {
      return {
        suggested_category: "",
        confidence: 0,
        reason: "OpenWebUI model discovery failed. No usable models were returned by /api/models.",
        needs_review: true
      };
    }

    let lastError = "";

    for (const modelId of modelCandidates) {
      const result = callOpenWebUiChatCompletion_(modelId, systemPrompt, userPrompt);

      if (!result.ok) {
        lastError =
          "Model " +
          modelId +
          " failed with HTTP " +
          result.code +
          ": " +
          String(result.text || "").slice(0, 500);
        console.log(lastError);
        continue;
      }

      const content = extractOpenWebUiMessageContent_(result.text);

      if (!content) {
        lastError =
          "Model " +
          modelId +
          " returned no message content: " +
          String(result.text || "").slice(0, 500);
        console.log(lastError);
        continue;
      }

      const parsed = parseJsonObjectFromText_(content);
      const suggested = sanitizeCategory_(parsed.suggested_category);
      const confidence = clampNumber_(Number(parsed.confidence || 0), 0, 1);

      return {
        suggested_category: suggested,
        confidence: confidence,
        reason: String(parsed.reason || "").slice(0, 500) + " [OpenWebUI model: " + modelId + "]",
        needs_review: Boolean(parsed.needs_review)
      };
    }

    return {
      suggested_category: "",
      confidence: 0,
      reason: "OpenWebUI failed for all candidate models. Last error: " + lastError,
      needs_review: true
    };

  } catch (err) {
    return {
      suggested_category: "",
      confidence: 0,
      reason: "OpenWebUI categorizer failed: " + err.message,
      needs_review: true
    };
  }
}

/**
 * Calls OpenWebUI's OpenAI-compatible chat completion endpoint.
 *
 * Important:
 * Do NOT send a made-up chat_id here.
 * OpenWebUI can return 404 "Something went wrong :/" when chat_id is present
 * but does not correspond to an existing chat owned by the API-key user.
 *
 * parent_id: null tells OpenWebUI this is a new API-side completion rather than
 * an update to an existing WebUI chat.
 */
function callOpenWebUiChatCompletion_(modelId, systemPrompt, userPrompt) {
  const payload = {
    model: modelId,
    parent_id: null,
    messages: [
      {
        role: "system",
        content: systemPrompt
      },
      {
        role: "user",
        content: userPrompt
      }
    ],
    stream: false,
    temperature: 0
  };

  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + CONFIG.OPENWEBUI_API_KEY,
      Accept: "application/json"
    }
  };

  const response = UrlFetchApp.fetch(CONFIG.OPENWEBUI_API_URL, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  return {
    ok: code >= 200 && code < 300,
    code: code,
    text: text
  };
}

/**
 * Returns a prioritized list of model IDs that actually exist in OpenWebUI.
 * Set CONFIG.OPENWEBUI_MODEL to:
 * - "AUTO" to auto-pick the best available Gemma/local model
 * - or an exact model ID from /api/models
 */
function getOpenWebUiModelCandidates_() {
  const configuredModel = String(CONFIG.OPENWEBUI_MODEL || "AUTO").trim();
  const configuredIsAuto = !configuredModel || /^AUTO$/i.test(configuredModel);

  const models = fetchOpenWebUiModels_();
  const modelIds = models
    .map(model => String(model.id || model.name || "").trim())
    .filter(Boolean);

  const candidates = [];

  function addCandidate(id) {
    const clean = String(id || "").trim();
    if (!clean) return;
    if (candidates.indexOf(clean) !== -1) return;
    candidates.push(clean);
  }

  // If user supplied an exact real model ID, try it first.
  if (!configuredIsAuto) {
    const exact = modelIds.find(id => id === configuredModel);
    if (exact) addCandidate(exact);

    const caseInsensitive = modelIds.find(
      id => id.toLowerCase() === configuredModel.toLowerCase()
    );
    if (caseInsensitive) addCandidate(caseInsensitive);
  }

  // Prefer Gemma/E4B-type models.
  modelIds
    .filter(id => /gemma/i.test(id) && /(e4b|4b|4-e|gguf|it)/i.test(id))
    .forEach(addCandidate);

  modelIds
    .filter(id => /gemma/i.test(id))
    .forEach(addCandidate);

  // Then prefer normal local/chat models over web/tools/embed/rerank wrappers.
  modelIds
    .filter(id => !/(edge-web|web|tool|tools|embedding|embed|rerank|whisper|stt|tts)/i.test(id))
    .forEach(addCandidate);

  // Last resort: include everything returned by OpenWebUI.
  modelIds.forEach(addCandidate);

  // If /api/models failed or returned nothing, fall back to configured model.
  if (!candidates.length && !configuredIsAuto) {
    addCandidate(configuredModel);
  }

  console.log("OpenWebUI candidate model IDs: " + JSON.stringify(candidates));
  return candidates;
}

/**
 * Fetches OpenWebUI models from /api/models.
 */
function fetchOpenWebUiModels_() {
  const modelsUrl = getOpenWebUiModelsUrl_();

  const options = {
    method: "get",
    muteHttpExceptions: true,
    headers: {
      Authorization: "Bearer " + CONFIG.OPENWEBUI_API_KEY,
      Accept: "application/json"
    }
  };

  const response = UrlFetchApp.fetch(modelsUrl, options);
  const code = response.getResponseCode();
  const text = response.getContentText();

  console.log("OpenWebUI models URL: " + modelsUrl);
  console.log("OpenWebUI /api/models HTTP " + code);
  console.log(String(text || "").slice(0, 2000));

  if (code < 200 || code >= 300) {
    throw new Error("Could not list OpenWebUI models. HTTP " + code + ": " + text.slice(0, 500));
  }

  let parsed;
  try {
    parsed = JSON.parse(text);
  } catch (err) {
    throw new Error("OpenWebUI /api/models returned non-JSON: " + text.slice(0, 500));
  }

  let list = [];

  if (parsed && Array.isArray(parsed.data)) {
    list = parsed.data;
  } else if (parsed && Array.isArray(parsed.models)) {
    list = parsed.models;
  } else if (Array.isArray(parsed)) {
    list = parsed;
  }

  return list
    .map(model => {
      if (typeof model === "string") {
        return {
          id: model,
          name: model
        };
      }

      return {
        id: String(model.id || model.name || "").trim(),
        name: String(model.name || model.id || "").trim()
      };
    })
    .filter(model => model.id);
}

/**
 * Converts:
 * https://your-domain.com/api/chat/completions
 * into:
 * https://your-domain.com/api/models
 */
function getOpenWebUiModelsUrl_() {
  const chatUrl = String(CONFIG.OPENWEBUI_API_URL || "").trim();

  if (/\/api\/chat\/completions\/?$/i.test(chatUrl)) {
    return chatUrl.replace(/\/api\/chat\/completions\/?$/i, "/api/models");
  }

  if (/\/api\/?$/i.test(chatUrl)) {
    return chatUrl.replace(/\/api\/?$/i, "/api/models");
  }

  return chatUrl.replace(/\/+$/g, "") + "/api/models";
}

/**
 * Extracts assistant content from common OpenAI/OpenWebUI response shapes.
 */
function extractOpenWebUiMessageContent_(text) {
  let raw;

  try {
    raw = JSON.parse(text);
  } catch (err) {
    return "";
  }

  if (
    raw &&
    raw.choices &&
    raw.choices.length &&
    raw.choices[0].message &&
    raw.choices[0].message.content
  ) {
    return raw.choices[0].message.content;
  }

  if (
    raw &&
    raw.choices &&
    raw.choices.length &&
    raw.choices[0].delta &&
    raw.choices[0].delta.content
  ) {
    return raw.choices[0].delta.content;
  }

  if (raw && raw.message && raw.message.content) {
    return raw.message.content;
  }

  if (raw && raw.response) {
    return raw.response;
  }

  if (raw && raw.content) {
    return raw.content;
  }

  return "";
}

/**
 * Optional diagnostic function.
 * Run this manually if the categorizer still fails.
 */
function testOpenWebUiListModels() {
  const models = fetchOpenWebUiModels_();

  const modelIds = models
    .map(model => String(model.id || model.name || "").trim())
    .filter(Boolean);

  console.log("Available OpenWebUI model IDs:");
  modelIds.forEach(id => console.log(id));

  showMessage_(
    "OpenWebUI models found:\n\n" +
    modelIds.join("\n") +
    "\n\nUse one of these exact IDs for CONFIG.OPENWEBUI_MODEL, or keep it as AUTO."
  );
}

function parseJsonObjectFromText_(text) {
  let clean = String(text || "").trim();

  clean = clean
    .replace(/^```json\s*/i, "")
    .replace(/^```\s*/i, "")
    .replace(/\s*```$/i, "")
    .trim();

  try {
    return JSON.parse(clean);
  } catch (err) {
    // Continue to object extraction.
  }

  const match = clean.match(/\{[\s\S]*\}/);
  if (!match) {
    throw new Error("No JSON object found in model response: " + clean.slice(0, 200));
  }

  return JSON.parse(match[0]);
}



function parseTransactionEmail_(subject, from, body, messageId) {
  const chaseParsed = parseChaseTransactionEmail_(subject, from, body, messageId);
  if (chaseParsed.is_transaction) return chaseParsed;

  const amexParsed = parseAmericanExpressTransactionEmail_(subject, from, body, messageId);
  if (amexParsed.is_transaction) return amexParsed;

  return {
    is_transaction: false,
    source: "Unknown",
    account: "",
    card_last4: "",
    transaction_date: "",
    transaction_time: "",
    timezone: "",
    merchant: "",
    amount: null,
    category: "Misc",
    confidence: 0,
    needs_review: true,
    gmail_message_id: messageId,
    duplicate_key: "",
    parse_notes: "No supported transaction format matched."
  };
}

function parseChaseTransactionEmail_(subject, from, body, messageId) {
  const result = {
    is_transaction: false,
    source: "Chase",
    account: "",
    card_last4: "",
    transaction_date: "",
    transaction_time: "",
    timezone: "",
    merchant: "",
    amount: null,
    category: "Misc",
    confidence: 0,
    needs_review: true,
    gmail_message_id: messageId,
    duplicate_key: "",
    parse_notes: ""
  };

  const combined = subject + "\n" + from + "\n" + body;
  const hasChaseClues = /transaction alert/i.test(combined) && /you made a/i.test(combined) && /merchant/i.test(combined) && /amount/i.test(combined);

  if (!hasChaseClues) {
    result.parse_notes = "Missing Chase transaction alert clues.";
    return result;
  }

  const accountMatch = body.match(/Account\s+(.+?\(\.\.\.\d{4}\))/i);
  const dateMatch = body.match(/Date\s+([A-Z][a-z]{2}\s+\d{1,2},\s+\d{4})\s+at\s+([0-9]{1,2}:[0-9]{2}\s+[AP]M)\s+([A-Z]{2})/i);
  const merchantMatch = body.match(/Merchant\s+([^\n\r]+)/i);
  const amountMatch = body.match(/Amount\s+\$?([0-9,]+(?:\.[0-9]{2})?)/i) || body.match(/You made a\s+\$?([0-9,]+(?:\.[0-9]{2})?)\s+transaction/i);

  if (accountMatch) {
    result.account = cleanText_(accountMatch[1]);
    const last4Match = result.account.match(/\(\.\.\.(\d{4})\)/);
    if (last4Match) result.card_last4 = last4Match[1];
  }

  if (dateMatch) {
    result.transaction_date = normalizeDateToIso_(dateMatch[1]);
    result.transaction_time = cleanText_(dateMatch[2]);
    result.timezone = cleanText_(dateMatch[3]);
  }

  if (merchantMatch) result.merchant = cleanText_(merchantMatch[1]);
  if (amountMatch) result.amount = Number(amountMatch[1].replace(/,/g, ""));

  const missing = [];
  if (!result.account) missing.push("account");
  if (!result.card_last4) missing.push("card_last4");
  if (!result.transaction_date) missing.push("transaction_date");
  if (!result.transaction_time) missing.push("transaction_time");
  if (!result.merchant) missing.push("merchant");
  if (result.amount === null || isNaN(result.amount)) missing.push("amount");

  if (missing.length === 0) {
    result.is_transaction = true;
    result.confidence = 0.98;
    result.needs_review = false;
    result.parse_notes = "Parsed Chase transaction alert.";
  } else {
    result.is_transaction = result.amount !== null || result.merchant !== "";
    result.confidence = 0.60;
    result.needs_review = true;
    result.parse_notes = "Partial Chase parse. Missing: " + missing.join(", ");
  }

  return result;
}


function parseAmericanExpressTransactionEmail_(subject, from, body, messageId) {
  const result = {
    is_transaction: false,
    source: "American Express",
    account: "",
    card_last4: "",
    transaction_date: "",
    transaction_time: "",
    timezone: "",
    merchant: "",
    amount: null,
    category: "Misc",
    confidence: 0,
    needs_review: true,
    gmail_message_id: messageId,
    duplicate_key: "",
    parse_notes: ""
  };

  const combined = subject + "\n" + from + "\n" + body;

  const senderIsAmex = /AmericanExpress@welcome\.americanexpress\.com/i.test(from);

  // Explicitly reject Amex statement, account snapshot, balance, payment,
  // due-date, and summary emails. These contain dollar amounts but are not
  // individual purchases and should never enter Pending Transactions.
  const isAmexStatementOrSnapshot =
    /Stay up to date on your account/i.test(combined) ||
    /weekly account snapshot/i.test(combined) ||
    /See your account at a glance/i.test(combined) ||
    /Statement Icon/i.test(combined) ||
    /Statement balance\s*:/i.test(combined) ||
    /closing date/i.test(combined) ||
    /Recent payments\s*&\s*credits\s*:/i.test(combined) ||
    /Recent charges\s*:/i.test(combined) ||
    /Total balance\s*:/i.test(combined) ||
    /Payment due\s*:/i.test(combined) ||
    /due on\s+(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday)/i.test(combined);

  if (senderIsAmex && isAmexStatementOrSnapshot) {
    result.parse_notes = "Skipped American Express statement/account snapshot email.";
    return result;
  }

  const hasAmexPurchaseClues =
    /American Express/i.test(combined) &&
    /Account\s+Ending\s*:\s*\d{4,6}/i.test(combined) &&
    (/large purchase on your Card/i.test(combined) || /See the details about this purchase/i.test(combined)) &&
    /\$\s*[0-9,]+(?:\.[0-9]{2})?\*?/i.test(combined) &&
    !isAmexStatementOrSnapshot;

  // Require both the trusted sender and purchase-alert clues. Sender alone is
  // not enough because Amex also sends statements and weekly snapshots from
  // this address.
  if (!senderIsAmex || !hasAmexPurchaseClues) {
    result.parse_notes = "Missing American Express purchase alert clues.";
    return result;
  }

  const accountMatch = body.match(/Account\s+Ending\s*:\s*(\d{4,6})/i);
  if (accountMatch) {
    const ending = cleanText_(accountMatch[1]);
    result.account = "American Express Account Ending " + ending;

    // Amex purchase alerts may show a 5-digit account ending. Keep the full
    // value for duplicate protection even though the sheet column says last 4.
    result.card_last4 = ending;
  }

  const lines = String(body || "")
    .split("\n")
    .map(line => cleanText_(line))
    .filter(Boolean);

  let amountLineIndex = -1;
  let amountMatch = null;

  for (let i = 0; i < lines.length; i++) {
    const match = lines[i].match(/^\$\s*([0-9,]+(?:\.[0-9]{2})?)\*?$/i);
    if (match) {
      amountLineIndex = i;
      amountMatch = match;
      break;
    }
  }

  // Fallback for clients that flatten line breaks.
  if (!amountMatch) {
    amountMatch = body.match(/\$\s*([0-9,]+(?:\.[0-9]{2})?)\*?/i);
  }

  if (amountMatch) {
    result.amount = Number(amountMatch[1].replace(/,/g, ""));
  }

  if (amountLineIndex > 0) {
    result.merchant = cleanText_(lines[amountLineIndex - 1]);

    for (let j = amountLineIndex + 1; j < Math.min(lines.length, amountLineIndex + 5); j++) {
      if (/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun),?\s+[A-Z][a-z]+\s+\d{1,2},\s+\d{4}$/i.test(lines[j])) {
        result.transaction_date = normalizeDateToIso_(lines[j]);
        break;
      }
    }
  }

  if (!result.merchant || !result.transaction_date) {
    const detailRegex = new RegExp(
      "(?:^|\\n)([^\\n$]{2,100})\\n+\\$\\s*([0-9,]+(?:\\.[0-9]{2})?)\\*?\\n+" +
      "((?:Mon|Tue|Wed|Thu|Fri|Sat|Sun),?\\s+[A-Z][a-z]+\\s+\\d{1,2},\\s+\\d{4})",
      "i"
    );
    const detailMatch = body.match(detailRegex);
    if (detailMatch) {
      if (!result.merchant) result.merchant = cleanText_(detailMatch[1]);
      if (result.amount === null || isNaN(result.amount)) result.amount = Number(detailMatch[2].replace(/,/g, ""));
      if (!result.transaction_date) result.transaction_date = normalizeDateToIso_(detailMatch[3]);
    }
  }

  const missing = [];
  if (!result.account) missing.push("account");
  if (!result.card_last4) missing.push("card_last4");
  if (!result.transaction_date) missing.push("transaction_date");
  if (!result.merchant) missing.push("merchant");
  if (result.amount === null || isNaN(result.amount)) missing.push("amount");

  if (missing.length === 0) {
    result.is_transaction = true;
    result.confidence = 0.98;
    result.needs_review = false;
    result.parse_notes = "Parsed American Express purchase alert.";
  } else {
    result.is_transaction = result.amount !== null || result.merchant !== "";
    result.confidence = 0.60;
    result.needs_review = true;
    result.parse_notes = "Partial American Express parse. Missing: " + missing.join(", ");
  }

  return result;
}

/********************************************************
 * RULES
 ********************************************************/

function findMerchantRule_(merchant, rules) {
  const value = String(merchant || "");

  for (const rule of rules) {
    if (!rule.pattern || !rule.enabled) continue;

    try {
      const regex = new RegExp(rule.pattern, "i");
      if (regex.test(value)) return rule;
    } catch (err) {
      console.warn("Invalid merchant rule regex: " + rule.pattern);
    }
  }

  return null;
}

function loadMerchantRules_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, RULE_HEADERS.length).getValues();
  return values
    .map(row => ({
      pattern: String(row[0] || "").trim(),
      category: sanitizeCategory_(row[1]),
      confidence: clampNumber_(Number(row[2] || CONFIG.DEFAULT_RULE_CONFIDENCE), 0, 1),
      autoApprove: toBoolean_(row[3], true),
      learningAllowed: toBoolean_(row[4], true),
      notes: String(row[5] || ""),
      enabled: toBoolean_(row[6], true)
    }))
    .filter(rule => rule.pattern && rule.enabled);
}

function addMerchantRule_(merchant, category, options) {
  const ss = getActiveSpreadsheetOrThrow_();
  const sheet = getOrCreateSheet_(ss, CONFIG.MERCHANT_RULES_SHEET_NAME);
  setupMerchantRulesSheet_(sheet);

  const cleanMerchant = normalizeMerchantForDisplay_(merchant);
  const cleanCategory = validateCategory_(category, "Category");
  const escapedPattern = escapeRegex_(cleanMerchant);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const existingPatterns = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0] || "").trim().toUpperCase());
    if (existingPatterns.includes(escapedPattern.toUpperCase())) return;
  }

  sheet.appendRow([
    escapedPattern,
    cleanCategory,
    clampNumber_(Number(options && options.confidence || CONFIG.DEFAULT_RULE_CONFIDENCE), 0, 1),
    options && options.autoApprove !== undefined ? Boolean(options.autoApprove) : true,
    options && options.learningAllowed !== undefined ? Boolean(options.learningAllowed) : true,
    options && options.notes ? options.notes : "Learned from review app.",
    true
  ]);
}

function shouldIgnoreTransaction_(merchant, ignoredRules) {
  const value = String(merchant || "");

  for (const rule of ignoredRules) {
    try {
      const regex = new RegExp(rule.pattern, "i");
      if (regex.test(value)) return true;
    } catch (err) {
      console.warn("Invalid ignore rule regex: " + rule.pattern);
    }
  }

  return false;
}

function loadIgnoredRules_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  return sheet.getRange(2, 1, lastRow - 1, IGNORE_HEADERS.length).getValues()
    .map(row => ({
      pattern: String(row[0] || "").trim(),
      notes: String(row[1] || ""),
      enabled: toBoolean_(row[2], true)
    }))
    .filter(rule => rule.pattern && rule.enabled);
}

function addIgnoredMerchantRule_(merchant, notes) {
  const ss = getActiveSpreadsheetOrThrow_();
  const sheet = getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME);
  setupIgnoredMerchantsSheet_(sheet);

  const cleanMerchant = normalizeMerchantForDisplay_(merchant);
  const escapedPattern = escapeRegex_(cleanMerchant);
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const existingPatterns = sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(row => String(row[0] || "").trim().toUpperCase());
    if (existingPatterns.includes(escapedPattern.toUpperCase())) return;
  }

  sheet.appendRow([escapedPattern, notes || "Ignored from review app.", true]);
}

function removeIgnoredPendingTransactions() {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);
  const ignoredSheet = getOrCreateSheet_(ss, CONFIG.IGNORED_MERCHANTS_SHEET_NAME);
  setupPendingSheet_(pendingSheet);
  setupIgnoredMerchantsSheet_(ignoredSheet);

  const ignoredRules = loadIgnoredRules_(ignoredSheet);
  const lastRow = pendingSheet.getLastRow();
  let count = 0;

  if (lastRow >= 2) {
    for (let rowNumber = lastRow; rowNumber >= 2; rowNumber--) {
      const merchant = String(pendingSheet.getRange(rowNumber, 2).getValue() || "");
      const status = String(pendingSheet.getRange(rowNumber, 11).getValue() || "");
      if (status === "Pending Review" && shouldIgnoreTransaction_(merchant, ignoredRules)) {
        pendingSheet.getRange(rowNumber, 11).setValue("Ignored");
        count++;
      }
    }
  }

  showMessage_("Marked ignored pending transactions: " + count);
}

/********************************************************
 * SHEET SETUP
 ********************************************************/

function setupPendingSheet_(sheet) {
  migratePendingSheetIfNeeded_(sheet);
  ensureExactHeaders_(sheet, PENDING_HEADERS);
  sheet.setFrozenRows(1);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("C:C").setNumberFormat("$#,##0.00");
  sheet.getRange("F:F").setNumberFormat("0.00");
  sheet.getRange("H:H").setNumberFormat("0.00");
  sheet.getRange("M:M").setNumberFormat("yyyy-mm-dd hh:mm:ss");

  const categoryRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.ALLOWED_CATEGORIES, true).setAllowInvalid(false).build();
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(["Pending Review", "Moved", "Skipped", "Ignored", "Duplicate", "Error"], true).setAllowInvalid(false).build();

  sheet.getRange(2, 4, Math.max(lastRow - 1, 1), 1).setDataValidation(categoryRule);
  sheet.getRange(2, 5, Math.max(lastRow - 1, 1), 1).setDataValidation(categoryRule);
  sheet.getRange(2, 7, Math.max(lastRow - 1, 1), 1).setDataValidation(categoryRule);
  sheet.getRange(2, 11, Math.max(lastRow - 1, 1), 1).setDataValidation(statusRule);

  try {
    sheet.getRange(1, 1, 1, PENDING_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, PENDING_HEADERS.length);
    sheet.hideColumns(14, 4);
    sheet.setHiddenGridlines(true);
  } catch (err) {
    console.log("Pending style skipped: " + err.message);
  }
}

function migratePendingSheetIfNeeded_(sheet) {
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), PENDING_HEADERS.length)).getValues()[0].map(v => String(v || "").trim());
  const alreadyCompact = PENDING_HEADERS.every((h, i) => currentHeaders[i] === h);
  if (alreadyCompact) return;

  const rowIsBlank = currentHeaders.every(h => !h);
  if (rowIsBlank) return;

  const ss = getActiveSpreadsheetOrThrow_();
  const backup = sheet.copyTo(ss).setName("Pending Backup " + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyyMMddHHmmss"));
  backup.hideSheet();

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const oldValues = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const oldHeaderMap = {};
  currentHeaders.forEach((header, index) => { if (header) oldHeaderMap[header] = index; });

  const migratedRows = oldValues.map(row => {
    const date = valueByHeader_(row, oldHeaderMap, ["Date", "Transaction Date"]);
    const merchant = valueByHeader_(row, oldHeaderMap, ["Merchant"]);
    const amount = valueByHeader_(row, oldHeaderMap, ["Amount"]);
    const category = sanitizeCategory_(valueByHeader_(row, oldHeaderMap, ["Selected Category", "Category"]));
    const ruleCategory = sanitizeCategory_(valueByHeader_(row, oldHeaderMap, ["Rule Category", "Category"]));
    const ruleConfidence = Number(valueByHeader_(row, oldHeaderMap, ["Rule Confidence", "Confidence"]) || 0);
    const aiCategory = sanitizeCategory_(valueByHeader_(row, oldHeaderMap, ["AI Category", "AI Suggestion"]));
    const aiConfidence = Number(valueByHeader_(row, oldHeaderMap, ["AI Confidence"]) || 0);
    const aiReason = valueByHeader_(row, oldHeaderMap, ["AI Reason"]);
    const reviewReason = valueByHeader_(row, oldHeaderMap, ["Review Reason", "Parse Notes"]);
    const duplicateKey = valueByHeader_(row, oldHeaderMap, ["Duplicate Key"]);
    const importedAt = valueByHeader_(row, oldHeaderMap, ["Imported At"]);
    const account = valueByHeader_(row, oldHeaderMap, ["Account"]);
    const cardLast4 = valueByHeader_(row, oldHeaderMap, ["Card Last 4"]);
    const source = valueByHeader_(row, oldHeaderMap, ["Source"]);
    const gmailMessageId = valueByHeader_(row, oldHeaderMap, ["Gmail Message ID"]);

    const reviewed = valueByHeader_(row, oldHeaderMap, ["Reviewed"]);
    const movedAt = valueByHeader_(row, oldHeaderMap, ["Moved To Expenses At"]);
    const status = reviewed === true || movedAt ? "Moved" : "Pending Review";

    return [
      date,
      merchant,
      amount,
      category,
      ruleCategory,
      ruleConfidence,
      aiCategory,
      aiConfidence,
      aiReason,
      reviewReason || "Migrated from older pending layout.",
      status,
      duplicateKey,
      importedAt || new Date(),
      account,
      cardLast4,
      source,
      gmailMessageId
    ];
  });

  sheet.clear();
  sheet.getRange(1, 1, 1, PENDING_HEADERS.length).setValues([PENDING_HEADERS]);
  if (migratedRows.length) {
    sheet.getRange(2, 1, migratedRows.length, PENDING_HEADERS.length).setValues(migratedRows);
  }
}

function setupExpenseSheet_(sheet) {
  ensureExactHeaders_(sheet, EXPENSE_HEADERS);
  sheet.setFrozenRows(1);

  const lastRow = Math.max(sheet.getLastRow(), 2);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("C:C").setNumberFormat("$#,##0.00");
  sheet.getRange("I:I").setNumberFormat("yyyy-mm-dd hh:mm:ss");

  const categoryRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.ALLOWED_CATEGORIES, true).setAllowInvalid(false).build();
  sheet.getRange(2, 4, Math.max(lastRow - 1, 1), 1).setDataValidation(categoryRule);

  try {
    sheet.getRange(1, 1, 1, EXPENSE_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, EXPENSE_HEADERS.length);
    sheet.setHiddenGridlines(true);
  } catch (err) {
    console.log("Expense style skipped: " + err.message);
  }
}

function setupMerchantRulesSheet_(sheet) {
  migrateMerchantRulesIfNeeded_(sheet);
  ensureExactHeaders_(sheet, RULE_HEADERS);
  sheet.setFrozenRows(1);

  if (sheet.getLastRow() === 1) {
    const starterRules = getStarterMerchantRules_();
    sheet.getRange(2, 1, starterRules.length, RULE_HEADERS.length).setValues(starterRules);
  }

  const lastRow = Math.max(sheet.getLastRow(), 2);
  const categoryRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.ALLOWED_CATEGORIES, true).setAllowInvalid(false).build();
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();

  sheet.getRange(2, 2, Math.max(lastRow - 1, 1), 1).setDataValidation(categoryRule);
  sheet.getRange(2, 4, Math.max(lastRow - 1, 1), 1).setDataValidation(checkboxRule);
  sheet.getRange(2, 5, Math.max(lastRow - 1, 1), 1).setDataValidation(checkboxRule);
  sheet.getRange(2, 7, Math.max(lastRow - 1, 1), 1).setDataValidation(checkboxRule);
  sheet.getRange("C:C").setNumberFormat("0.00");

  try {
    sheet.getRange(1, 1, 1, RULE_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, RULE_HEADERS.length);
  } catch (err) {
    console.log("Rule style skipped: " + err.message);
  }
}

function migrateMerchantRulesIfNeeded_(sheet) {
  const currentHeaders = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), RULE_HEADERS.length)).getValues()[0].map(v => String(v || "").trim());
  const alreadyNew = RULE_HEADERS.every((h, i) => currentHeaders[i] === h);
  if (alreadyNew) return;

  const rowIsBlank = currentHeaders.every(h => !h);
  if (rowIsBlank) return;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  const oldValues = lastRow >= 2 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const oldHeaderMap = {};
  currentHeaders.forEach((header, index) => { if (header) oldHeaderMap[header] = index; });

  const migrated = oldValues.map(row => {
    const pattern = valueByHeader_(row, oldHeaderMap, ["Pattern"]);
    const category = sanitizeCategory_(valueByHeader_(row, oldHeaderMap, ["Category"]));
    const notes = String(valueByHeader_(row, oldHeaderMap, ["Notes"]) || "Migrated rule.");
    const enabled = toBoolean_(valueByHeader_(row, oldHeaderMap, ["Enabled"]), true);
    const isAmbiguous = /AMAZON|AMZN|TARGET|WALMART|COSTCO\$|PAYPAL|VENMO|APPLE|GOOGLE/i.test(String(pattern || "")) || category === "Misc";
    return [pattern, category, isAmbiguous ? 0.40 : CONFIG.DEFAULT_RULE_CONFIDENCE, !isAmbiguous, !isAmbiguous, notes, enabled];
  }).filter(row => row[0]);

  sheet.clear();
  sheet.getRange(1, 1, 1, RULE_HEADERS.length).setValues([RULE_HEADERS]);
  if (migrated.length) sheet.getRange(2, 1, migrated.length, RULE_HEADERS.length).setValues(migrated);
}

function setupIgnoredMerchantsSheet_(sheet) {
  ensureExactHeaders_(sheet, IGNORE_HEADERS);
  sheet.setFrozenRows(1);
  const lastRow = Math.max(sheet.getLastRow(), 2);
  const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
  sheet.getRange(2, 3, Math.max(lastRow - 1, 1), 1).setDataValidation(checkboxRule);

  try {
    sheet.getRange(1, 1, 1, IGNORE_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, IGNORE_HEADERS.length);
  } catch (err) {}
}

function setupReviewSettingsSheet_(sheet) {
  ensureExactHeaders_(sheet, SETTINGS_HEADERS);
  sheet.setFrozenRows(1);

  const defaults = [
    ["MAX_UNCERTAINTY_FOR_AUTO_APPROVAL", CONFIG.DEFAULT_AUTO_APPROVE_UNCERTAINTY, "0.10 means 90% confidence or higher can auto-approve."],
    ["USE_AI_CATEGORIZER", true, "Use OpenWebUI AI categorizer when rules are uncertain."],
    ["AI_ONLY_FOR_UNCERTAIN", true, "Only call AI when no confident auto-approve rule exists."],
    ["AUTO_APPROVE_AI", true, "Allow confident AI recommendations to go directly to Expenses."],
    ["REVIEW_MISC", true, "Misc should generally require review unless explicitly overridden by a rule."]
  ];

  const existing = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, 1).getValues().forEach(row => existing.add(String(row[0] || "")));
  }

  for (const setting of defaults) {
    if (!existing.has(setting[0])) sheet.appendRow(setting);
  }

  try {
    sheet.getRange(1, 1, 1, SETTINGS_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, SETTINGS_HEADERS.length);
  } catch (err) {}
}

function setupImportLogSheet_(sheet) {
  ensureExactHeaders_(sheet, IMPORT_LOG_HEADERS);
  sheet.setFrozenRows(1);
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd hh:mm:ss");
  try {
    sheet.getRange(1, 1, 1, IMPORT_LOG_HEADERS.length).setFontWeight("bold").setBackground("#1f4e79").setFontColor("#ffffff");
    sheet.autoResizeColumns(1, IMPORT_LOG_HEADERS.length);
  } catch (err) {}
}

function getStarterMerchantRules_() {
  return [
    ["STARBUCKS|DUNKIN|COFFEE|CAFE|CAFÉ|BAKERY|BAGEL|ICE CREAM|SMOOTHIE", "Restaurants", 0.90, true, true, "Coffee, bakery, dessert, and quick-service food", true],
    ["MCDONALD|MCDONALD'S|MCDONALDS|CULVER|CULVERS|BURGER KING|WENDY|SUBWAY|PANERA|CHIPOTLE|QDOBA|TACO BELL|PIZZA|RESTAURANT|GRILL|DINER|CAFE", "Restaurants", 0.95, true, true, "Restaurants and fast food", true],
    ["ALDI|HY-?VEE|HYVEE|FAREWAY|TRADER JOE|FRESH THYME|KROGER|JEWEL|WHOLE FOODS|SAFEWAY|PUBLIX|MEIJER|GROCERY|SUPERMARKET", "Groceries", 0.98, true, true, "Grocery stores", true],
    ["CASEY'?S|KWIK STAR|SHELL|BP\b|MOBIL|EXXON|PHILLIPS 66|MARATHON|KUM & GO|CIRCLE K|SPEEDWAY|COSTCO GAS|GAS COSTCO|FUEL", "Fuel", 0.98, true, true, "Gas and fuel", true],
    ["SUPERCHARGE|SUPERCHARGER|TESLA SUPERCHARGER|EV CHARGING|ELECTRIFY AMERICA|CHARGEPOINT", "Fuel", 0.98, true, true, "EV charging", true],
    ["ELECTRIC|WATER|SEWER|UTILITY|UTILITIES|GAS BILL|POWER|ENERGY|COMCAST|XFINITY|SPECTRUM|MEDIACOM|VERIZON|T-MOBILE|TMOBILE|AT&T|INTERNET|PHONE BILL", "Utilities", 0.95, true, true, "Utilities, internet, and phone", true],
    ["MENARDS|LOWE'?S|LOWES|HOME DEPOT|ACE HARDWARE|HARBOR FREIGHT|AUTO ?ZONE|O'?REILLY|ADVANCE AUTO|TIRE|JIFFY LUBE|VALVOLINE|CAR WASH|MECHANIC|AUTO REPAIR|SERVICE CENTER", "Home/Car Maintenance", 0.95, true, true, "Home improvement and car maintenance", true],
    ["PETSMART|PETCO|CHEWY|VET|VETERINARY|ANIMAL|PET SUPPLY|DOG|CAT|GROOMING", "Pet", 0.95, true, true, "Pet expenses", true],
    ["HAIRCUT|HAIR CUT|SALON|BARBER|NAILS?|ULTA|SEPHORA|COSMETIC|BEAUTY|MAKEUP", "Hair/Cosmetics", 0.95, true, true, "Hair and cosmetics", true],
    ["CVS|WALGREENS|PHARMACY|DOCTOR|CLINIC|HOSPITAL|MEDICAL|DENTAL|DENTIST|VISION|OPTICAL|GLASSES|COPAY|HEALTH", "Medical", 0.95, true, true, "Medical, dental, vision, and pharmacy", true],
    ["NETFLIX|HULU|DISNEY\+|SPOTIFY|YOUTUBE|THEATER|CINEMA|MOVIE|BOWLING|VON MAUR|KOHL'?S|OLD NAVY|GAP|NIKE|ADIDAS|DSW|SHOES|CLOTHING|APPAREL", "Entertainment/Clothing", 0.90, true, true, "Entertainment and clothing", true],
    ["HOTEL|HILTON|MARRIOTT|HYATT|HOLIDAY INN|AIRBNB|VRBO|LODGING|RESORT|DELTA|UNITED AIRLINES|AMERICAN AIRLINES|SOUTHWEST|FLIGHT|AIRPORT|TRAVEL|VACATION", "Vacation", 0.90, true, true, "Travel and vacation", true],
    ["HSA|FSA|BENEFITS|BENEFIT PROVIDER|INSURANCE PREMIUM|PAYROLL DEDUCTION", "Misc", 0.40, false, false, "Benefits-related merchants are user-specific. Manually map to My Benefits or Spouse Benefits if appropriate.", true],
    ["AMAZON|AMZN", "Misc", 0.40, false, false, "Ambiguous merchant; review unless specific context is available", true],
    ["COSTCO(?!.*GAS)", "Misc", 0.40, false, false, "Ambiguous merchant; review unless specific context is available", true],
    ["TARGET", "Misc", 0.40, false, false, "Ambiguous merchant; review unless specific context is available", true],
    ["WALMART", "Misc", 0.40, false, false, "Ambiguous merchant; review unless specific context is available", true],
    ["PAYPAL|VENMO|APPLE|GOOGLE", "Misc", 0.40, false, false, "Ambiguous payment/platform merchant", true]
  ];
}

/********************************************************
 * DATA WRITES / DUPLICATES
 ********************************************************/

function appendPendingTransactionRow_(sheet, tx) {
  const row = [
    tx.transaction_date,
    tx.merchant,
    tx.amount,
    sanitizeCategory_(tx.selected_category || tx.category),
    sanitizeCategory_(tx.rule_category),
    clampNumber_(Number(tx.rule_confidence || 0), 0, 1),
    sanitizeCategory_(tx.ai_category),
    clampNumber_(Number(tx.ai_confidence || 0), 0, 1),
    tx.ai_reason || "",
    tx.review_reason || "Needs review.",
    "Pending Review",
    tx.duplicate_key,
    new Date(),
    tx.account || "",
    tx.card_last4 || "",
    tx.source || "",
    tx.gmail_message_id || ""
  ];

  sheet.appendRow(row);
}

function appendExpenseRow_(sheet, tx, approvalSource) {
  const cleanCategory = sanitizeCategory_(tx.category);

  // Push to the Google Form first. If this fails, the caller should not
  // mark the pending transaction as moved.
  submitExpenseToGoogleForm_(tx.merchant, tx.amount, cleanCategory);

  const row = [
    tx.transaction_date,
    tx.merchant,
    tx.amount,
    cleanCategory,
    tx.account || "",
    tx.card_last4 || "",
    tx.source || "",
    tx.duplicate_key,
    new Date(),
    approvalSource || "Unknown"
  ];

  sheet.appendRow(row);
}

function findPendingRowByDuplicateKey_(sheet, duplicateKey) {
  const key = String(duplicateKey || "").trim();
  if (!key) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const finder = sheet.getRange(2, 12, lastRow - 1, 1).createTextFinder(key).matchEntireCell(true);
  const match = finder.findNext();
  return match ? match.getRow() : null;
}

function expenseDuplicateExists_(sheet, duplicateKey) {
  const key = String(duplicateKey || "").trim();
  if (!key) return false;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const match = sheet.getRange(2, 8, lastRow - 1, 1).createTextFinder(key).matchEntireCell(true).findNext();
  return Boolean(match);
}

function getExistingPendingDuplicateKeys_(sheet) {
  const duplicateKeys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return duplicateKeys;

  sheet.getRange(2, 12, lastRow - 1, 1).getValues().forEach(row => {
    const key = String(row[0] || "").trim();
    if (key) duplicateKeys.add(key);
  });

  return duplicateKeys;
}

function getExistingExpenseDuplicateKeys_(sheet) {
  const duplicateKeys = new Set();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return duplicateKeys;

  sheet.getRange(2, 8, lastRow - 1, 1).getValues().forEach(row => {
    const key = String(row[0] || "").trim();
    if (key) duplicateKeys.add(key);
  });

  return duplicateKeys;
}

function createDuplicateKey_(tx) {
  return [
    tx.transaction_date || "",
    normalizeForKey_(tx.merchant || ""),
    Number(tx.amount || 0).toFixed(2),
    tx.card_last4 || ""
  ].join("|");
}

function logImport_(action, merchant, amount, category, duplicateKey, notes) {
  try {
    const ss = getActiveSpreadsheetOrThrow_();
    const sheet = getOrCreateSheet_(ss, CONFIG.IMPORT_LOG_SHEET_NAME);
    setupImportLogSheet_(sheet);
    sheet.appendRow([new Date(), action || "", merchant || "", amount || "", category || "", duplicateKey || "", notes || ""]);
  } catch (err) {
    console.log("Import log skipped: " + err.message);
  }
}


/********************************************************
 * REVIEW EMAIL NOTIFICATIONS
 ********************************************************/

function sendReviewNeededEmail_(reviewItems) {
  if (!CONFIG.SEND_REVIEW_EMAILS) return;
  if (isPlaceholderConfigValue_(CONFIG.NOTIFICATION_EMAIL)) return;
  if (!Array.isArray(reviewItems) || reviewItems.length === 0) return;

  const subject =
    "Budget Review: " +
    reviewItems.length +
    " transaction" +
    (reviewItems.length === 1 ? "" : "s") +
    " need approval";

  const lines = reviewItems.map(item => {
    const merchant = String(item.merchant || "Unknown merchant");
    const amount = Number(item.amount || 0).toLocaleString("en-US", {
      style: "currency",
      currency: "USD"
    });
    const selectedCategory = sanitizeCategory_(item.selected_category || item.category || "Misc");
    const aiCategory = item.ai_category ? sanitizeCategory_(item.ai_category) : "";
    const aiConfidence = Number(item.ai_confidence || 0);
    const aiText = aiCategory
      ? "AI: " + aiCategory + (aiConfidence ? " (" + Math.round(aiConfidence * 100) + "%)" : "")
      : "AI: none";
    const reason = String(item.review_reason || "").trim();

    return [
      "• " + merchant + " — " + amount,
      "  Selected: " + selectedCategory + " | " + aiText,
      reason ? "  Reason: " + reason : ""
    ].filter(Boolean).join("\n");
  });

  const body =
    "The Budget Review app has " +
    reviewItems.length +
    " new transaction" +
    (reviewItems.length === 1 ? "" : "s") +
    " waiting for approval.\n\n" +
    lines.join("\n\n") +
    "\n\nOpen Budget Review:\n" +
    CONFIG.WEB_APP_URL;

  MailApp.sendEmail({
    to: CONFIG.NOTIFICATION_EMAIL,
    subject: subject,
    body: body
  });
}

function testReviewNeededEmail() {
  sendReviewNeededEmail_([
    {
      merchant: "TEST BUDGET REVIEW TRANSACTION",
      amount: 0.01,
      selected_category: "Misc",
      ai_category: "Restaurants",
      ai_confidence: 0.95,
      review_reason: "Test notification only."
    }
  ]);

  showMessage_("Sent test review-needed email to " + CONFIG.NOTIFICATION_EMAIL + ".");
}

/********************************************************
 * SETTINGS
 ********************************************************/

function loadReviewSettings_(sheet) {
  setupReviewSettingsSheet_(sheet);

  const settings = {
    MAX_UNCERTAINTY_FOR_AUTO_APPROVAL: CONFIG.DEFAULT_AUTO_APPROVE_UNCERTAINTY,
    USE_AI_CATEGORIZER: true,
    AI_ONLY_FOR_UNCERTAIN: true,
    AUTO_APPROVE_AI: true,
    REVIEW_MISC: true
  };

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();
    values.forEach(row => {
      const key = String(row[0] || "").trim();
      const value = row[1];
      if (!key) return;

      if (key === "MAX_UNCERTAINTY_FOR_AUTO_APPROVAL") {
        settings[key] = clampNumber_(Number(value || CONFIG.DEFAULT_AUTO_APPROVE_UNCERTAINTY), 0, 1);
      } else {
        settings[key] = toBoolean_(value, settings[key]);
      }
    });
  }

  return settings;
}

function getBooleanSetting_(key, defaultValue) {
  try {
    const ss = getActiveSpreadsheetOrThrow_();
    const sheet = getOrCreateSheet_(ss, CONFIG.REVIEW_SETTINGS_SHEET_NAME);
    const settings = loadReviewSettings_(sheet);
    return settings[key] !== undefined ? Boolean(settings[key]) : defaultValue;
  } catch (err) {
    return defaultValue;
  }
}

/********************************************************
 * TESTS + TRIGGERS
 ********************************************************/

function testParseExample() {
  const sampleBody = `
Chase Logo
Transaction alert
You made a $8.48 transaction

Account Chase Freedom Unlimited Visa (...1234)
Date Apr 29, 2026 at 2:14 PM ET
Merchant STARBUCKS
Amount $8.48
`;

  const parsed = parseTransactionEmail_("Transaction alert", "Chase <no.reply.alerts@chase.com>", normalizeEmailBody_(sampleBody), "TEST_MESSAGE_ID");
  parsed.merchant = normalizeMerchantForDisplay_(parsed.merchant);
  parsed.duplicate_key = createDuplicateKey_(parsed);
  showMessage_("Test parsed transaction:\n\n" + JSON.stringify(parsed, null, 2));
}


function testParseAmexExample() {
  const sampleBody = `
See the details about this purchase

American Express Logo CARDHOLDER NAME
Account Ending: 12345

Card Art

There was a large purchase on your Card

Dear CARDHOLDER NAME,

As you requested, we're letting you know that this purchase was more than $10.00.

WAL MART COM

$44.47*

Sat, May 9, 2026

*The amount above may not reflect the final amount as some merchants issue a pre-authorization charge
`;

  const parsed = parseTransactionEmail_(
    "See the details about this purchase",
    "American Express <AmericanExpress@welcome.americanexpress.com>",
    normalizeEmailBody_(sampleBody),
    "TEST_AMEX_MESSAGE_ID"
  );

  parsed.merchant = normalizeMerchantForDisplay_(parsed.merchant);
  parsed.duplicate_key = createDuplicateKey_(parsed);
  showMessage_("Test parsed Amex transaction:\n\n" + JSON.stringify(parsed, null, 2));
}


function testParseAmexStatementExclusion() {
  const sampleBody = `
Stay up to date on your account

American Express Logo CARDHOLDER NAME
Account ending: 123456

Card Art
Statement Icon
See your account at a glance

Here's your weekly account snapshot for your Example Card® from American Express

As of Thursday, April 23, 2026

Statement balance:

$176.06

closing date Thursday, April 16, 2026

Recent payments & credits:

$0.00

Recent charges:

$124.07

since Friday, April 17, 2026

Total balance:

$300.13

Payment due:

$40.00

due on Monday, May 11, 2026
`;

  const parsed = parseTransactionEmail_(
    "Stay up to date on your account",
    "American Express <AmericanExpress@welcome.americanexpress.com>",
    normalizeEmailBody_(sampleBody),
    "TEST_AMEX_STATEMENT_MESSAGE_ID"
  );

  showMessage_(
    "Amex statement exclusion test. is_transaction should be false:\n\n" +
    JSON.stringify(parsed, null, 2)
  );
}

function testOpenWebUiCategorizerFromAppsScript() {
  const tx = {
    merchant: "STARBUCKS",
    amount: 8.48,
    transaction_date: "2026-04-29",
    account: "Chase Freedom Unlimited Visa (...1234)",
    card_last4: "1234",
    category: "Misc"
  };

  const result = getAiCategorySuggestion_(tx);
  showMessage_(JSON.stringify(result, null, 2));
}

function createHourlyImportTrigger() {
  deleteTriggersForFunction_("importGmailTransactions");
  ScriptApp.newTrigger("importGmailTransactions").timeBased().everyHours(1).create();
  showMessage_("Hourly import trigger created.");
}

function deleteBudgetTriggers() {
  deleteTriggersForFunction_("importGmailTransactions");
  showMessage_("Budget triggers deleted.");
}

function deleteTriggersForFunction_(functionName) {
  const triggers = ScriptApp.getProjectTriggers();
  for (const trigger of triggers) {
    if (trigger.getHandlerFunction() === functionName) {
      ScriptApp.deleteTrigger(trigger);
    }
  }
}

/********************************************************
 * HELPERS
 ********************************************************/

function buildGmailSearchQuery_() {
  return [
    "newer_than:" + CONFIG.SEARCH_DAYS_BACK + "d",
    '-label:"' + CONFIG.PROCESSED_LABEL + '"',
    '(("Transaction alert" AND ("You made a" OR "Merchant" OR "Amount")) OR (from:AmericanExpress@welcome.americanexpress.com AND ("There was a large purchase on your Card" OR "See the details about this purchase")))',
    '-("weekly account snapshot" OR "Statement balance" OR "Payment due" OR "Total balance" OR "Recent charges" OR "Stay up to date on your account")'
  ].join(" ");
}

function getActiveSpreadsheetOrThrow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      "No active Google Sheet found. This script must be installed from the Google Sheet you want to use: open the sheet, then go to Extensions > Apps Script and paste this code into Code.gs."
    );
  }

  return ss;
}

function showMessage_(message) {
  console.log(message);
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    console.log("UI alert skipped: " + err.message);
  }
}

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function getOrCreateGmailLabel_(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);
  if (!label) label = GmailApp.createLabel(labelName);
  return label;
}

function ensureExactHeaders_(sheet, headers) {
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const rowIsBlank = current.every(cell => String(cell || "").trim() === "");

  if (rowIsBlank) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return;
  }

  for (let i = 0; i < headers.length; i++) {
    if (String(current[i] || "").trim() !== headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
}

function valueByHeader_(row, headerMap, possibleHeaders) {
  for (const header of possibleHeaders) {
    if (headerMap[header] !== undefined) return row[headerMap[header]];
  }
  return "";
}

function normalizeEmailBody_(body) {
  return String(body || "")
    .replace(/\u00A0/g, " ")
    .replace(/[ \t]+/g, " ")
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\n{2,}/g, "\n")
    .trim();
}

function normalizeMerchantForDisplay_(merchant) {
  let value = cleanText_(merchant);

  value = value.replace(/^SQ \*/i, "");
  value = value.replace(/^TST\*/i, "");
  value = value.replace(/^PAYPAL \*/i, "");
  value = value.replace(/\s+\d{3}[- ]?\d{3}[- ]?\d{4}\s*$/i, "");
  value = value.replace(/\s+\d{5}(?:-\d{4})?\s*$/i, "");
  value = value.replace(/AMZN MKTP.*|AMAZON MKTPLACE.*/i, "AMAZON");

  // American Express sometimes sends Walmart as "WAL MART COM".
  // Normalize it so your existing WALMART ambiguous merchant rule catches it.
  value = value.replace(/\bWAL[-\s]*MART\b/i, "WALMART");

  return cleanText_(value);
}

function cleanText_(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function normalizeDateToIso_(dateText) {
  const date = new Date(dateText);
  if (isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function normalizeForKey_(value) {
  return String(value || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function formatDateForWeb_(value) {
  if (!value) return "";
  try {
    const date = new Date(value);
    if (isNaN(date.getTime())) return String(value);
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } catch (err) {
    return String(value);
  }
}

function escapeRegex_(text) {
  return String(text || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function clampNumber_(value, min, max) {
  if (isNaN(value)) return min;
  return Math.min(Math.max(value, min), max);
}

function toBoolean_(value, defaultValue) {
  if (value === true) return true;
  if (value === false) return false;
  const text = String(value || "").trim().toUpperCase();
  if (text === "TRUE" || text === "YES" || text === "Y" || text === "1") return true;
  if (text === "FALSE" || text === "NO" || text === "N" || text === "0") return false;
  return Boolean(defaultValue);
}

function sanitizeCategory_(category) {
  const clean = String(category || "").trim();
  for (const allowed of CONFIG.ALLOWED_CATEGORIES) {
    if (allowed.toUpperCase() === clean.toUpperCase()) return allowed;
  }
  return "Misc";
}

function validateCategory_(category, label) {
  const raw = String(category || "").trim();
  const clean = sanitizeCategory_(raw);
  if (!raw) throw new Error((label || "Category") + " is blank.");
  if (clean === "Misc" && raw.toUpperCase() !== "MISC") {
    throw new Error((label || "Category") + " is not valid: " + raw);
  }
  return clean;
}

function submitExpenseToGoogleForm_(merchant, amount, category) {
  if (!CONFIG.PUSH_TO_GOOGLE_FORM) {
    return {
      ok: true,
      skipped: true,
      message: "Google Form push disabled."
    };
  }

  const cleanMerchant = String(merchant || "").trim();
  const cleanAmount = Number(amount || 0);
  const cleanCategory = mapCategoryForGoogleForm_(category);

  if (!cleanMerchant) {
    throw new Error("Cannot submit to Google Form: expense name is blank.");
  }

  if (isNaN(cleanAmount) || cleanAmount <= 0) {
    throw new Error("Cannot submit to Google Form: amount is invalid.");
  }

  if (isPlaceholderConfigValue_(CONFIG.GOOGLE_FORM_URL)) {
    throw new Error("Cannot submit to Google Form: set CONFIG.GOOGLE_FORM_URL or set PUSH_TO_GOOGLE_FORM to false.");
  }

  const form = FormApp.openByUrl(CONFIG.GOOGLE_FORM_URL);
  const response = form.createResponse();
  const items = form.getItems();

  const nameItem = findFormItemByTitle_(items, "What is Expense (Name)");
  const amountItem = findFormItemByTitle_(items, "Amount ($)");
  const categoryItem = findFormItemByTitle_(items, "Expense Category");

  response.withItemResponse(
    nameItem.asTextItem().createResponse(cleanMerchant)
  );

  response.withItemResponse(
    amountItem.asTextItem().createResponse(cleanAmount.toFixed(2))
  );

  response.withItemResponse(
    categoryItem.asMultipleChoiceItem().createResponse(cleanCategory)
  );

  response.submit();

  return {
    ok: true,
    skipped: false,
    message: "Submitted to Google Form."
  };
}

function findFormItemByTitle_(items, title) {
  const expected = normalizeFormTitle_(title);

  for (const item of items) {
    if (normalizeFormTitle_(item.getTitle()) === expected) {
      return item;
    }
  }

  throw new Error(
    "Could not find Google Form item titled '" +
    title +
    "'. Available titles: " +
    items.map(item => item.getTitle()).join(", ")
  );
}

function normalizeFormTitle_(title) {
  return String(title || "")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}


function isPlaceholderConfigValue_(value) {
  const text = String(value || "").trim();
  if (!text) return true;
  if (/^x$/i.test(text)) return true;
  if (/PASTE_YOUR_/i.test(text)) return true;
  if (/example\.com/i.test(text)) return true;
  return false;
}

function mapCategoryForGoogleForm_(category) {
  const cleanCategory = sanitizeCategory_(category);
  return CONFIG.FORM_CATEGORY_MAP[cleanCategory] || "Misc";
}

/**
 * Legacy compatibility stub.
 * Older versions used a 23-column Pending Transactions layout with Reviewed and Moved At columns.
 * The current reviewer uses Status + duplicate keys + processReviewActions(), so this function is intentionally disabled.
 */
function movePendingRowsToExpenses_(specificRowNumber) {
  throw new Error("movePendingRowsToExpenses_ is disabled. Use the Budget Review web app, which calls processReviewActions().");
}

function testGoogleFormExpenseSubmit() {
  const result = submitExpenseToGoogleForm_(
    "TEST IMPORTER EXPENSE",
    0.01,
    "Misc"
  );

  showMessage_(JSON.stringify(result, null, 2));
}

function addFakePendingReviewTransaction() {
  const ss = getActiveSpreadsheetOrThrow_();
  const pendingSheet = getOrCreateSheet_(ss, CONFIG.PENDING_SHEET_NAME);

  setupPendingSheet_(pendingSheet);

  const now = new Date();
  const testDate = Utilities.formatDate(
    now,
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const fakeTx = {
    transaction_date: testDate,
    transaction_time: Utilities.formatDate(
      now,
      Session.getScriptTimeZone(),
      "h:mm a"
    ),
    timezone: Session.getScriptTimeZone(),
    merchant: "TEST BUDGET REVIEW TRANSACTION",
    amount: 0.01,
    category: "Misc",
    ai_category: "Restaurants",
    ai_reason: "Test transaction only. This is used to verify the Budget Review web app approval flow.",
    needs_review: true,
    ai_needs_review: false,
    account: "Test Account (...0000)",
    card_last4: "0000",
    source: "Test",
    confidence: 0.50,
    duplicate_key: "TEST|" + now.getTime(),
    gmail_message_id: "TEST_MESSAGE_" + now.getTime(),
    parse_notes: "Fake test row created by addFakePendingReviewTransaction.",
    ai_confidence: 0.95
  };

  appendPendingTransactionRow_(
    pendingSheet,
    fakeTx,
    "TEST Transaction alert",
    "Test Sender <test@example.com>"
  );

  showMessage_(
    "Fake $0.01 pending transaction created.\n\n" +
    "Open the Budget Review web app and test the approval buttons."
  );
}

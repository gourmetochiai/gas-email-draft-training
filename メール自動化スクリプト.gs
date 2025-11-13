/**
 * スプレッドシートの data シートを読み込み、A〜Eのカテゴリに応じて
 * config シートのテンプレから Gmail の下書きを作成します。
 * 置換: {Name}, {Company}, {MeetingLink}
 */
const MEETING_LINK = "https://example.com/meet";

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("自動化")
    .addItem("下書き作成", "createDrafts")
    .addToUi();
}

function createDrafts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSh = ss.getSheetByName("data");
  const configSh = ss.getSheetByName("config");
  let logSh = ss.getSheetByName("log");
  if (!logSh) logSh = ss.insertSheet("log");

  const data = dataSh.getDataRange().getValues();
  const header = data.shift();
  const idx = Object.fromEntries(header.map((h,i)=>[h,i]));

  // 既存の列がなければ追加（Category, Processed）
  ensureColumn(dataSh, "Category", header);
  ensureColumn(dataSh, "Processed", header);

  const conf = readConfig(configSh); // {A:{subj,body}, ...}

  const lastRow = dataSh.getLastRow();
  const outCategoryCol = findColIndex(dataSh, "Category") + 1;
  const outProcessedCol = findColIndex(dataSh, "Processed") + 1;

  const resultLogs = [["Timestamp", "Row", "Email", "Category", "Status", "Message"]];

  for (let r = 2; r <= lastRow; r++) {
    const row = dataSh.getRange(r, 1, 1, dataSh.getLastColumn()).getValues()[0];
    const rowObj = Object.fromEntries(header.map((h,i)=>[h, row[i]]));

    // すでに処理済みならスキップ
    const processed = dataSh.getRange(r, outProcessedCol).getValue();
    if (processed === true || processed === "TRUE") continue;

    const email = rowObj["Email"];
    if (!email || typeof email !== "string" || email.indexOf("@") === -1) {
      resultLogs.push([new Date(), r, email || "", "", "SKIP", "Invalid email"]);
      continue;
    }

    // カテゴリ判定
    const category = decideCategory(rowObj);
    dataSh.getRange(r, outCategoryCol).setValue(category);

    // テンプレ取得
    const tpl = conf[category] || conf["E"];
    const subject = replaceVars(tpl.subject, rowObj);
    const body = replaceVars(replaceVars(tpl.body, rowObj), {MeetingLink: ""}); // safe chain
    const bodyFinal = body.replaceAll("{MeetingLink}", MEETING_LINK);

    try {
      GmailApp.createDraft(email, subject, bodyFinal);
      dataSh.getRange(r, outProcessedCol).setValue(true);
      resultLogs.push([new Date(), r, email, category, "OK", "Draft created"]);
    } catch (e) {
      resultLogs.push([new Date(), r, email, category, "ERROR", e.message]);
    }
  }

  // ログ書き込み
  const last = logSh.getLastRow();
  logSh.getRange(last+1, 1, resultLogs.length, resultLogs[0].length).setValues(resultLogs);
  SpreadsheetApp.getUi().alert("下書き作成が完了しました。");
}

function ensureColumn(sh, colName, header) {
  const idx = header.indexOf(colName);
  if (idx === -1) {
    sh.insertColumnAfter(header.length);
    sh.getRange(1, header.length+1).setValue(colName);
    header.push(colName);
  }
}

function findColIndex(sh, colName) {
  const h = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  return h.indexOf(colName);
}

function readConfig(configSh) {
  const vals = configSh.getDataRange().getValues();
  const head = vals.shift(); // Category, Subject, Body
  const idx = Object.fromEntries(head.map((h,i)=>[h,i]));
  const conf = {};
  vals.forEach(row => {
    const cat = row[idx["Category"]];
    conf[cat] = { subject: row[idx["Subject"]], body: row[idx["Body"]] };
  });
  return conf;
}

function decideCategory(row) {
  const title = String(row["Title"] || "");
  const tag = String(row["Tag"] || "");

  if (title.includes("代表") || title.includes("部長") || tag.includes("ホット")) return "A";
  if (title.includes("課長") || title.includes("マネージャ") || tag.includes("既存")) return "B";
  if (title.includes("担当") || tag.includes("要フォロー")) return "C";
  if (tag.includes("展示会")) return "D";
  return "E";
}

function replaceVars(text, row) {
  let s = String(text);
  Object.entries(row).forEach(([k,v])=>{
    s = s.replaceAll(`{${k}}`, String(v || ""));
  });
  // 互換: 明示キーにも対応
  s = s.replaceAll("{Name}", String(row["Name"] || ""));
  s = s.replaceAll("{Company}", String(row["Company"] || ""));
  s = s.replaceAll("{MeetingLink}", String(row["MeetingLink"] || ""));
  return s;
}

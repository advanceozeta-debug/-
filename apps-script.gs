// 잭팟 머신 뽑기 기록 — Google Apps Script
// 1) 아래 ADMIN_KEY를 나만 아는 비밀번호로 바꾸세요. 관리자 사이트에서 기록을 볼 때 이 비밀번호를 넣어요.
const ADMIN_KEY = "여기에-비밀번호";

const SHEET = "기록";

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET);
  if (!sh) {
    sh = ss.insertSheet(SHEET);
    sh.appendRow(["시간", "결과", "그림", "상품", "플레이어", "기기"]);
    sh.setFrozenRows(1);
  }
  return sh;
}

// 게임이 한 번 돌릴 때마다 한 줄씩 저장
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse((e.postData && e.postData.contents) || "{}");
    sheet_().appendRow([
      new Date(),
      d.r === "W" ? "당첨" : "꽝",
      String(d.f || "").replace(/[^1-7]/g, "").slice(0, 3),
      String(d.p || "").slice(0, 100),
      String(d.id || "").slice(0, 20),
      String(d.ua || "").slice(0, 10)
    ]);
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput("ok");
}

// 관리자 사이트가 기록을 읽어 감 (비밀번호가 맞을 때만)
function doGet(e) {
  if (!e.parameter.key || e.parameter.key !== ADMIN_KEY) return json_({ error: "key" });
  const sh = sheet_();
  const last = sh.getLastRow();
  const total = Math.max(0, last - 1);
  let wins = 0, today = 0;
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  if (total) {
    const all = sh.getRange(2, 1, total, 2).getValues();
    for (const [t, r] of all) {
      if (r === "당첨") wins++;
      if (Utilities.formatDate(new Date(t), tz, "yyyy-MM-dd") === todayStr) today++;
    }
  }
  const limit = Math.min(Number(e.parameter.limit) || 300, 2000);
  const n = Math.min(limit, total);
  const rows = n ? sh.getRange(last - n + 1, 1, n, 6).getValues().reverse().map(r => ({
    t: new Date(r[0]).toISOString(), r: r[1], f: String(r[2]), p: r[3], id: r[4], ua: r[5]
  })) : [];
  return json_({ total, wins, today, rows });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

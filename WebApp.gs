// =============================================================
// 영주 할매 예약 — 서버(API) v27  (색 겹침: 끝난 일정 포함 · 색만 바뀐 것도 충돌로 알아챔 · 날짜 시간대 통일)
//  · 화면은 GitHub 의 index.html. 여기는 데이터만 주고받음 (doPost)
//  · Apps Script 의 Index.html 은 이제 안 씀 → 지워도 됨
//  · DATA 시트 (5행부터): B 오는 날 | C 오는 시간 | D 가는 날 | E 색(1~6) | F 이름
//  · 예약: 끝난 일정도 시트에서 지우지 않고 그대로 둠 → 달력에서 2026. 11. 01 부터 볼 수 있음 (끝난 일정은 수정 · 삭제 불가)
//  · RED 시트 (5행부터): B 날짜 → 달력에서 빨간 날로 표시 (지난 날짜도 지우지 않음 → 지난달 달력에서도 빨간 날)
//  · KEYS 시트 (4행 머리글, 5행부터): B 이름 | C 열쇠 | D 사용(체크) | E 마지막 사용 | F 초대 링크
//      → 초대 링크(열쇠)가 있는 사람만 열림. 시트 메뉴 "영주 할매 → 새 초대 링크 만들기"로 만듦
//  · LOG 시트: 누가(어느 열쇠로) 언제 무엇을 저장했는지 자동으로 쌓임
//  · 배포: 실행 = 나 / 액세스 = 모든 사용자 (열쇠 검사는 doPost 맨 앞에서)
//  · API: 요청 본문(text/plain JSON) { k: 초대 링크 열쇠, fn: 함수명, args: [인자] }
//         응답 { ok: true, data } / { ok: false, error, auth: 열쇠 문제면 true }
//         부를 수 있는 함수: getAll, saveBooking (새 함수는 doPost 의 API 목록에 추가)
// Code.gs 의 getTodayZero(), DATA_SHEET, WEB_START_ROW 를 사용합니다.
// =============================================================
const WEB_TITLE = "영주 할매";
const RED_SHEET = "RED";
const WEB_NCOLS = 5;          // B ~ F
const WEB_COLORS = 6;         // 고를 수 있는 색 개수
const WEB_NAME_FORMAT = '"  "@';   // F열(이름) 표시 서식: 앞에 띄어쓰기 2칸 (값은 이름 그대로, 보이는 모양만)

// ▼ 앱 껍데기 주소 (GitHub Pages) — 초대 링크를 만들 때 씀. 저장소 이름이 다르면 여기만 바꾸세요
const SHELL_URL = "https://dix0314.github.io/sigol/";
const KEYS_SHEET = "KEYS";
const KEYS_HEAD_ROW = 4;      // KEYS 머리글 줄 (DATA 시트처럼 4행 머리글 · 5행부터 내용)
const KEYS_COL = 2;           // KEYS 첫 칸 = B열 (B 이름 · C 열쇠 · D 사용 · E 마지막 사용 · F 초대 링크)
const LOG_SHEET = "LOG";

// ★ 달력 시작점: 2026. 11. 01 (바꾸지 마세요)
//   · 가는 날이 이 날짜 이후인 예약을 모두 앱에 보냄 (끝난 일정 포함) — 앱 index.html 의 CAL_START 와 같은 날짜
//   · 앱은 처음 열 때 항상 오늘 날짜 · 이번 달부터 보여 주고, 이 달(2026년 11월)보다 앞으로는 넘어가지 않음
const WEB_CAL_START = "2026-11-01";

// 주소를 그냥 열면: 서버가 돌고 있는지만 알려 줌 (화면은 GitHub 에 있음)
function doGet() {
  return ContentService.createTextOutput(WEB_TITLE + " 서버 작동 중");
}

// 앱이 부르는 곳
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
    const who = webKeyName_(String(body.k || ""));
    if (!who) {
      Utilities.sleep(1000);                                   // 틀린 열쇠를 마구 넣어 보는 것을 느리게
      out = { ok: false, error: "초대 링크로 열어 주세요.", auth: true };
    } else {
      const API = {
        getAll: function () { return webAll_(); },
        saveBooking: function (req) { return webSave_(req || {}, who); }
      };
      const fn = API[body.fn];
      if (!fn) throw new Error("없는 기능이에요: " + body.fn);
      out = { ok: true, data: fn.apply(null, body.args || []) };
    }
  } catch (err) {
    out = { ok: false, error: err && err.message ? err.message : String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// 전체 데이터 (웹앱이 호출) — 가는 날이 달력 시작점(2026. 11. 01) 이후인 예약 전부 (끝난 일정 포함)
// -------------------------------------------------------------
function webAll_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = ss.getSpreadsheetTimeZone();
  const fromKey = WEB_CAL_START;                          // 달력 시작점부터 (끝난 일정도 보냄)

  const sheet = ss.getSheetByName(DATA_SHEET);
  const list = [];
  if (sheet) {
    webReadRows_(sheet).forEach(function (r) {
      const b = webRowToBooking_(r, tz);
      if (b && b.go >= fromKey) list.push(b);
    });
  }
  return { bookings: list, reds: webReadReds_(ss, tz) };
}

// -------------------------------------------------------------
// 예약 추가 / 수정 / 삭제
// req = { mode: "add"|"edit"|"delete", id, old:{come,time,go,text},
//         come:"yyyy-MM-dd", time(분|null), go:"yyyy-MM-dd", text, color(1~6) }
//  · 기간이 겹치는 예약끼리는 같은 색을 쓸 수 없음
//  · add: 오는 날이 오늘 이전이면 안 됨 / edit: 오늘 이전 날짜는 원래 날짜 그대로일 때만 됨
//  · edit · delete: 끝난 일정(가는 날이 오늘보다 앞)은 안 됨
// -------------------------------------------------------------
function webSave_(req, who) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = ss.getSpreadsheetTimeZone();
    const sheet = ss.getSheetByName(DATA_SHEET);
    if (!sheet) throw new Error("DATA 시트가 없어요.");

    const rows = webReadRows_(sheet);   // 시트는 한 번만 읽어서 색 검사 · 빈 줄 찾기 · 정리에 같이 씀 (저장 속도)
    let row = Number(req.id);
    let cur = null;
    if (req.mode !== "add") {
      if (!(row >= WEB_START_ROW)) throw new Error("잘못된 예약이에요.");
      const rng = sheet.getRange(row, 2, 1, WEB_NCOLS);
      cur = webRowToBooking_({ row: row, vals: rng.getValues()[0], disp: rng.getDisplayValues()[0] }, tz);
      const o = req.old || {};
      const same = cur && cur.come === o.come && cur.go === o.go && cur.text === o.text &&
        (cur.time == null ? null : cur.time) === (o.time == null ? null : Number(o.time)) &&
        (!("color" in o) || (cur.color == null ? null : cur.color) === (o.color == null ? null : Number(o.color)));   // 색만 바뀐 것도 알아챔
      if (!same) throw new Error("시트 내용이 바뀌었어요. 새로고침 후 다시 시도해 주세요.");
      // 끝난 일정(가는 날이 오늘보다 앞)은 수정 · 삭제 불가 — 앱에서도 막지만 서버에서 한 번 더
      if (cur.go < webKey_(getTodayZero(), tz)) throw new Error("끝난 일정은 수정 · 삭제할 수 없어요.");
    }

    if (req.mode === "delete") {
      sheet.getRange(row, 2, 1, WEB_NCOLS).clearContent();
    } else if (req.mode !== "add" && req.mode !== "edit") {
      throw new Error("없는 기능이에요: " + req.mode);
    } else {
      const text = String(req.text || "").trim();
      if (!text) throw new Error("이름을 입력해 주세요.");
      const come = webKeyToDate_(req.come, tz), go = webKeyToDate_(req.go, tz);
      if (!come || !go) throw new Error("날짜를 확인해 주세요.");
      if (req.go < req.come) throw new Error("가는 날을 오는 날 이후로 설정해 주세요.");
      // 지난 날짜: 새 예약은 안 됨 / 수정은 원래 날짜 그대로 둔 경우만 됨
      const todayK = webKey_(getTodayZero(), tz);
      if (req.mode === "add" && req.come < todayK) throw new Error("지난 날짜는 예약할 수 없어요.");
      if (req.mode === "edit" && ((req.come < todayK && req.come !== cur.come) || (req.go < todayK && req.go !== cur.go))) {
        throw new Error("지난 날짜로는 바꿀 수 없어요.");
      }
      const t = req.time == null || req.time === "" ? null : Number(req.time);
      if (t != null && !(t >= 0 && t < 1440)) throw new Error("오는 시간을 확인해 주세요.");

      const color = webColor_(req.color) || 1;
      if (webUsedColors_(rows, tz, req.come, req.go, req.mode === "edit" ? row : 0)[color]) {
        throw new Error("이미 쓰는 색이에요. 다른 색을 골라 주세요.");
      }

      if (req.mode === "add") row = webFreeRow_(sheet, rows);
      webWriteRow_(sheet, row, come, t, go, text, color);
    }
    webTidy_(ss, tz, rows);
    webLog_(ss, who, req, cur);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return webAll_();
}

// =============================================================
// 초대 링크 (열쇠)
// =============================================================

// 열쇠 → 이름 (사용 중인 열쇠만). KEYS 시트는 10분 동안 기억해 둠 (매번 시트를 읽지 않게)
function webKeyName_(k) {
  if (!/^[A-Za-z0-9_-]{16,}$/.test(k)) return null;
  const cache = CacheService.getScriptCache();
  let map = null;
  const c = cache.get("keys");
  if (c) map = JSON.parse(c);
  else {
    map = {};
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYS_SHEET);
    const first = KEYS_HEAD_ROW + 1;
    if (sh && sh.getLastRow() >= first) {
      sh.getRange(first, KEYS_COL, sh.getLastRow() - first + 1, 3).getValues().forEach(function (r, i) {   // 이름 · 열쇠 · 사용
        const key = String(r[1]).trim();
        if (key && r[2] === true) map[key] = { name: String(r[0]).trim() || "이름 없음", row: first + i };
      });
    }
    cache.put("keys", JSON.stringify(map), 600);   // 10분 (새 링크를 만들면 바로 지움)
  }
  const hit = map[k];
  if (!hit) return null;
  // 마지막 사용: 한 시간에 한 번만 적음
  if (!cache.get("used_" + k)) {
    cache.put("used_" + k, "1", 3600);
    try { SpreadsheetApp.getActiveSpreadsheet().getSheetByName(KEYS_SHEET).getRange(hit.row, KEYS_COL + 3).setValue(new Date()); } catch (e) {}
  }
  return hit.name;
}

// 저장 기록: 시간 | 누구(열쇠 이름) | 무엇을 | 예약 이름 | 오는 날 | 가는 날
function webLog_(ss, who, req, cur) {
  let sh = ss.getSheetByName(LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(LOG_SHEET);
    sh.appendRow(["시간", "누구", "무엇을", "예약 이름", "오는 날", "가는 날"]);
    sh.setFrozenRows(1);
  }
  const what = { add: "추가", edit: "수정", "delete": "삭제" }[req.mode] || req.mode;
  const b = req.mode === "delete" ? cur : req;
  sh.appendRow([new Date(), who, what, b.text, b.come, b.go]);
}

// KEYS 시트를 고치면(체크 빼기 등) 기억해 둔 열쇠 목록을 바로 지움 → 끊기가 바로 적용
function onEdit(e) {
  if (e && e.range && e.range.getSheet().getName() === KEYS_SHEET) CacheService.getScriptCache().remove("keys");
}

// 시트 메뉴
function onOpen() {
  SpreadsheetApp.getUi().createMenu(WEB_TITLE)
    .addItem("새 초대 링크 만들기", "webNewInvite")
    .addToUi();
}

// 새 초대 링크: 이름을 물어보고 → 긴 무작위 열쇠를 KEYS 시트에 추가 → 링크를 보여 줌
function webNewInvite() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt("새 초대 링크", "누구에게 줄 링크인가요? (예: 큰누나네)", ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;
  const name = res.getResponseText().trim() || "이름 없음";
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(KEYS_SHEET);
  if (!sh) {
    sh = ss.insertSheet(KEYS_SHEET);
    sh.getRange(KEYS_HEAD_ROW, KEYS_COL, 1, 5).setValues([["이름", "열쇠", "사용", "마지막 사용", "초대 링크"]]).setFontWeight("bold");
    sh.setFrozenRows(KEYS_HEAD_ROW);
    sh.setColumnWidth(KEYS_COL + 1, 280); sh.setColumnWidth(KEYS_COL + 4, 460);
  }
  const key = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, "").slice(0, 40);   // 40자 무작위
  // 넣을 줄: 5행부터 열쇠 칸(C열)이 빈 첫 줄 (중간에 지운 줄이 있으면 그 자리)
  const first = KEYS_HEAD_ROW + 1, last = Math.max(sh.getLastRow(), first - 1);
  let row = last + 1;
  if (last >= first) {
    const ks = sh.getRange(first, KEYS_COL + 1, last - first + 1, 1).getValues();
    for (let i = 0; i < ks.length; i++) if (String(ks[i][0]).trim() === "") { row = first + i; break; }
  }
  const link = SHELL_URL + "#k=" + key;
  sh.getRange(row, KEYS_COL, 1, 5).setValues([[name, key, true, "", link]]);
  sh.getRange(row, KEYS_COL + 2).insertCheckboxes().setValue(true);
  sh.getRange(row, KEYS_COL + 3).setNumberFormat("yyyy-MM-dd HH:mm");
  CacheService.getScriptCache().remove("keys");
  ui.alert("초대 링크 (" + name + ")", link + "\n\nKEYS 시트 F열에서도 복사할 수 있어요.\n끊으려면 D열 체크를 빼면 돼요 (바로 막힘).", ui.ButtonSet.OK);
}

// =============================================================
// 내부 헬퍼
// =============================================================

// 한 줄 쓰기 (서식 포함)
function webWriteRow_(sheet, row, come, t, go, text, color) {
  const rng = sheet.getRange(row, 2, 1, WEB_NCOLS);   // 서식 · 값을 한 번씩만 씀 (예전엔 칸마다 따로 5번)
  rng.setNumberFormats([["MMdd. ddd", "HH:mm", "MMdd. ddd", "0", WEB_NAME_FORMAT]]);   // B 오는 날 | C 시간 | D 가는 날 | E 색 | F 이름
  rng.setValues([[come, t == null ? "" : webHM_(t), go, color || "", text]]);
}

// 색 번호 1~6 (아니면 null)
function webColor_(v) {
  const n = parseInt(v, 10);
  return n >= 1 && n <= WEB_COLORS ? n : null;
}

// RED 시트 B열 날짜들 → ["yyyy-MM-dd", ...]
function webReadReds_(ss, tz) {
  const sh = ss.getSheetByName(RED_SHEET);
  if (!sh) return [];
  const last = sh.getLastRow();
  if (last < WEB_START_ROW) return [];
  const rng = sh.getRange(WEB_START_ROW, 2, last - WEB_START_ROW + 1, 1);
  const vals = rng.getValues(), disp = rng.getDisplayValues();
  const out = [];
  for (let i = 0; i < vals.length; i++) {
    const k = webDateKey_(vals[i][0], disp[i][0], tz);
    if (k) out.push(k);
  }
  return out;
}

// 예약의 색 번호: 고른 색, 비어 있으면 이름 첫 단어로 자동 (화면과 같은 계산)
function webColorNo_(b) {
  if (b.color) return b.color;
  const w = String(b.text).split(/[\s,]/)[0];
  let h = 0;
  for (let i = 0; i < w.length; i++) h = (h * 31 + w.charCodeAt(i)) >>> 0;
  return h % WEB_COLORS + 1;
}

// 기간 [come, go] 와 겹치는 다른 예약들이 쓰는 색 번호
function webUsedColors_(rows, tz, come, go, skipRow) {
  const used = {};
  rows.forEach(function (r) {
    if (r.row === skipRow) return;
    const b = webRowToBooking_(r, tz);
    // 끝난 일정도 포함 (달력 시작점 이후) — 앱 화면과 같은 기준: 지나간 날에도 같은 색 막대가 나란히 서지 않게
    if (b && b.go >= WEB_CAL_START && b.come <= go && b.go >= come) used[webColorNo_(b)] = true;
  });
  return Object.keys(used).length >= WEB_COLORS ? {} : used;   // 6색이 다 쓰이면 막지 않음
}

// 한 줄 → 예약 (날짜나 이름이 없으면 null)
function webRowToBooking_(r, tz) {
  const come = webDateKey_(r.vals[0], r.disp[0], tz);
  const text = String(r.disp[4]).trim();                     // F열: 이름
  if (!come || !text) return null;
  let go = webDateKey_(r.vals[2], r.disp[2], tz);
  if (!go || go < come) go = come;                           // 가는 날이 비면 당일로 봄
  return { id: r.row, come: come, time: webParseTime_(r.disp[1]), go: go, text: text, color: webColor_(r.disp[3]) };   // E열: 색
}

// 시작 행부터 마지막 행까지 읽기 (완전히 빈 줄은 건너뜀)
function webReadRows_(sheet) {
  const last = sheet.getLastRow();
  if (last < WEB_START_ROW) return [];
  const n = last - WEB_START_ROW + 1;
  const rng = sheet.getRange(WEB_START_ROW, 2, n, WEB_NCOLS);
  const vals = rng.getValues();
  const disp = rng.getDisplayValues();
  const out = [];
  for (let i = 0; i < n; i++) {
    if (disp[i].every(function (v) { return String(v).trim() === ""; })) continue;
    out.push({ row: WEB_START_ROW + i, vals: vals[i], disp: disp[i] });
  }
  return out;
}

// 정리: 예약 · 빨간 날 모두 지우지 않음 (지난 것도 그대로 남김) → 예약만 오는 날 · 시간 순으로 정렬
function webTidy_(ss, tz, rows) {
  const sheet = ss.getSheetByName(DATA_SHEET);
  if (!sheet) return;
  const last = sheet.getLastRow();
  if (last > WEB_START_ROW) {
    sheet.getRange(WEB_START_ROW, 2, last - WEB_START_ROW + 1, WEB_NCOLS)
      .sort([{ column: 2, ascending: true }, { column: 3, ascending: true }]);
  }
}

// 비어 있는 첫 줄 (중간에 지운 줄이 있으면 그 자리를 다시 씀) — rows(이미 읽어 둔 내용 있는 줄)로 판단
function webFreeRow_(sheet, rows) {
  const last = sheet.getLastRow();
  const filled = {};
  rows.forEach(function (r) { filled[r.row] = true; });
  for (let r = WEB_START_ROW; r <= last; r++) if (!filled[r]) return r;
  const row = Math.max(last + 1, WEB_START_ROW);
  const max = sheet.getMaxRows();
  if (row > max) sheet.insertRowsAfter(max, row - max);
  return row;
}

function webKey_(d, tz) {
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

// "yyyy-MM-dd" → 날짜 (시트 시간대 기준 그날 0시 — 읽을 때와 같은 기준이라 하루 밀리지 않음)
function webKeyToDate_(key, tz) {
  const m = String(key || "").match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return tz ? Utilities.parseDate(m[0], tz, "yyyy-MM-dd") : new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

// 날짜 칸 → "yyyy-MM-dd"
function webDateKey_(v, disp, tz) {
  if (v instanceof Date && !isNaN(v.getTime())) return webKey_(v, tz);
  const m = String(disp || "").match(/(\d{4})\D+(\d{1,2})\D+(\d{1,2})/);
  if (!m) return null;
  return m[1] + "-" + ("0" + m[2]).slice(-2) + "-" + ("0" + m[3]).slice(-2);
}

// "10:15", "오후 3:00", "3:00 PM" → 분 (비어 있으면 null)
function webParseTime_(s) {
  s = String(s || "").trim();
  if (!s) return null;
  const m = s.match(/(\d{1,2})(?::(\d{2}))?(?::\d{2})?/);
  if (!m) return null;
  let h = Number(m[1]);
  const mi = Number(m[2] || 0);
  if (/(PM|오후)/i.test(s) && h < 12) h += 12;
  if (/(AM|오전)/i.test(s) && h === 12) h = 0;
  if (h > 23 || mi > 59) return null;
  return h * 60 + mi;
}

function webHM_(min) {
  return ("0" + Math.floor(min / 60)).slice(-2) + ":" + ("0" + (min % 60)).slice(-2);
}
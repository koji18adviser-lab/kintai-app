/*************************************************************
 * 勤怠アプリ → スプレッドシート連携 (Google Apps Script)
 *
 * できること:
 *   - アプリの「シートに同期」ボタンから送られた打刻データを受け取り
 *   - 「明細」シートに全セッション（1日に何回でも・複数回勤務）を記録
 *   - 「月次（Stoke公式フォーマット）」シートに日別集計を自動入力
 *       出社時刻 = その日の最初の勤務開始
 *       退社時刻 = その日の最後の勤務終了
 *       休憩時間 = (退社 - 出社) - 実働時間  ← 空き時間・休憩を全部含む
 *       勤務時間 = 実働時間（その日の勤務セッションの合計）
 *       深夜労働 = 22時〜翌5時に重なった実働（設定で変更可）
 *       深夜休憩時間 = 22時〜翌5時に重なった休憩（深夜帯の空き時間 − 深夜労働）
 *     ※ 普通残業 / 40H超残業 / 休日出勤 の列は既存の計算式に任せます（触りません）
 *       → これらは「勤務時間」が入れば月次テンプレ側の式で自動計算されます
 *
 * 設定 (CONFIG):
 *   TZ            … タイムゾーン。日本なら "Asia/Tokyo"
 *   MONTHLY_SHEET … 月次フォーマットのシート名。"" なら先頭シートを使用
 *   DETAIL_SHEET  … 明細ログのシート名（無ければ自動作成）
 *   WRITE_NIGHT   … 深夜労働列をアプリの値で上書きするか（既存式を残すなら false）
 *   CLEAR_EMPTY   … その月に記録の無い日の 出社/退社/休憩 を空にするか
 *************************************************************/

var CONFIG = {
  TZ: "Asia/Tokyo",
  MONTHLY_SHEET: "",       // 例: "勤務時間表"。空なら先頭シート
  DETAIL_SHEET: "明細",
  WRITE_NIGHT: true,
  WRITE_WORK: true,        // 勤務時間・深夜休憩時間 をアプリの値で入力（テンプレに式があれば上書き）
  CLEAR_EMPTY: true
};

/* ---------- エンドポイント ---------- */
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var sessions = (payload.sessions || []).filter(function (s) {
      return s && s.start && !s.del;
    });
    var settings = payload.settings || {};
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    var detail = writeDetail_(ss, sessions, settings);
    var monthly = writeMonthly_(ss, sessions, settings);

    return json_({ ok: true, detail: detail, written: monthly.written, month: monthly.month });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  }
}

function doGet() {
  return json_({ ok: true, msg: "kintai sheet endpoint is alive" });
}

function json_(o) {
  return ContentService
    .createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ---------- 時間ユーティリティ ---------- */
function isBreak_(s) { return s.type === "break"; }
function durMin_(s) { return s.end ? (new Date(s.end) - new Date(s.start)) / 60000 : 0; }
function dayKey_(iso) { return Utilities.formatDate(new Date(iso), CONFIG.TZ, "yyyy-MM-dd"); }
function hhmm_(d) { return Utilities.formatDate(new Date(d), CONFIG.TZ, "HH:mm"); }
function fmtHM_(min) {
  min = Math.round(min);
  var sign = min < 0 ? "-" : "";
  min = Math.abs(min);
  return sign + Math.floor(min / 60) + ":" + ("0" + (min % 60)).slice(-2);
}

// 区間と深夜帯(既定22:00〜翌5:00)の重なり(分)。日またぎ対応。アプリ側ロジックと同一
function nightMin_(s, settings) {
  if (!s.end) return 0;
  var ns = settings.nightStart != null ? settings.nightStart : 22;
  var ne = settings.nightEnd != null ? settings.nightEnd : 5;
  var start = new Date(s.start), end = new Date(s.end), total = 0;
  var cursor = new Date(start); cursor.setHours(0, 0, 0, 0); cursor.setDate(cursor.getDate() - 1);
  var limit = new Date(end); limit.setHours(0, 0, 0, 0); limit.setDate(limit.getDate() + 1);
  while (cursor <= limit) {
    var w1 = new Date(cursor); w1.setHours(ns, 0, 0, 0);
    var w2 = new Date(cursor); w2.setDate(w2.getDate() + 1); w2.setHours(ne, 0, 0, 0);
    var a = Math.max(start, w1), b = Math.min(end, w2);
    if (b > a) total += (b - a) / 60000;
    cursor.setDate(cursor.getDate() + 1);
  }
  return total;
}

var WD_ = ["日", "月", "火", "水", "木", "金", "土"];
function weekday_(dk) {
  // dk = "yyyy-MM-dd"（その日の暦日）。正午UTCで曜日だけ取り出す（TZの影響を受けない）
  return WD_[new Date(dk + "T12:00:00Z").getUTCDay()];
}

/* ---------- 明細シート（全打刻ログ・公式シートと同じトーンに整形） ---------- */
function writeDetail_(ss, sessions, settings) {
  var sh = ss.getSheetByName(CONFIG.DETAIL_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.DETAIL_SHEET);

  var header = ["日付", "曜日", "種別", "開始", "終了", "実働(時:分)", "実働(分)", "深夜(時:分)", "深夜(分)", "休憩(時:分)"];
  var nCol = header.length;

  // データ行を構築（行ごとの日付キーも保持して日別グループ分けに使う）
  var body = [], dayOf = [];
  sessions.slice()
    .sort(function (a, b) { return new Date(a.start) - new Date(b.start); })
    .forEach(function (s) {
      var br = isBreak_(s), dk = dayKey_(s.start), dur = durMin_(s), nm = br ? 0 : nightMin_(s, settings);
      body.push([
        dk, weekday_(dk), br ? "休憩" : "勤務",
        hhmm_(s.start), s.end ? hhmm_(s.end) : "(勤務中)",
        br ? "" : (s.end ? fmtHM_(dur) : ""), br ? "" : (s.end ? Math.round(dur) : ""),
        nm > 0 ? fmtHM_(nm) : "", nm > 0 ? Math.round(nm) : "",
        (br && s.end) ? fmtHM_(dur) : ""
      ]);
      dayOf.push(dk);
    });

  // いったん中身も書式も全消去（前回の色・罫線が残らないように）
  sh.clear();
  sh.getRange(1, 1, 1, nCol).breakApart();

  // 1行目：タイトル帯（公式シートのグレー帯に合わせる）
  sh.getRange(1, 1, 1, nCol).merge();
  sh.getRange(1, 1)
    .setValue("勤怠 明細（全打刻ログ）　※1日に複数回の勤務もそのまま記録")
    .setFontWeight("bold").setFontColor("#ffffff").setBackground("#434343")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(1, 30);

  // 2行目：ヘッダー（公式シートの青ヘッダーに合わせる）
  sh.getRange(2, 1, 1, nCol).setValues([header])
    .setFontWeight("bold").setBackground("#c9daf8")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  sh.setRowHeight(2, 26);

  var lastRow = 2 + body.length;

  // データ書き込み＋日別の薄い背景でグループ分け
  if (body.length) {
    sh.getRange(3, 1, body.length, nCol).setValues(body)
      .setHorizontalAlignment("center").setVerticalAlignment("middle");

    var bg = [], shade = false, prevDay = null;
    for (var i = 0; i < body.length; i++) {
      if (dayOf[i] !== prevDay) { shade = !shade; prevDay = dayOf[i]; } // 日が変わるたび背景を交互に
      var rowBg = [];
      for (var c = 0; c < nCol; c++) rowBg.push(shade ? "#f3f6fc" : "#ffffff");
      if (body[i][2] === "休憩") rowBg[2] = "#fff2cc";                  // 休憩は種別セルを薄オレンジ
      bg.push(rowBg);
      if (body[i][4] === "(勤務中)") {                                  // 勤務中は緑で強調
        sh.getRange(3 + i, 5).setFontColor("#1e8e3e").setFontWeight("bold");
      }
    }
    sh.getRange(3, 1, body.length, nCol).setBackgrounds(bg);
  }

  // 全体に罫線（公式シートと同じ薄い線）
  sh.getRange(1, 1, Math.max(lastRow, 2), nCol)
    .setBorder(true, true, true, true, true, true, "#9aa6c2", SpreadsheetApp.BorderStyle.SOLID);

  // 列幅を内容に合わせて調整
  var widths = [96, 46, 54, 60, 74, 86, 70, 86, 70, 86];
  for (var w = 0; w < nCol; w++) sh.setColumnWidth(w + 1, widths[w]);

  sh.setFrozenRows(2);
  return body.length; // 件数
}

/* ---------- 月次シート（Stoke公式フォーマットに自動入力） ---------- */
function writeMonthly_(ss, sessions, settings) {
  var sh = CONFIG.MONTHLY_SHEET ? ss.getSheetByName(CONFIG.MONTHLY_SHEET) : ss.getSheets()[0];
  if (!sh) throw new Error("月次シートが見つかりません（CONFIG.MONTHLY_SHEET を確認）");

  var data = sh.getDataRange().getValues();

  // 1) ヘッダー行と列位置を特定（"日にち"・"出社時刻"・"退社時刻"・"休憩時間"・"深夜労働"）
  var hRow = -1, col = {};
  for (var r = 0; r < data.length; r++) {
    var txt = data[r].map(function (v) { return String(v).replace(/\s/g, ""); });
    if (txt.indexOf("日にち") >= 0 && txt.join("").indexOf("出社時刻") >= 0) {
      hRow = r;
      for (var c = 0; c < txt.length; c++) {
        var t = txt[c];
        if (t === "日にち" && col.day == null) col.day = c;
        else if (t === "出社時刻" && col.in == null) col.in = c;     // 最初の「出社時刻」
        else if (t === "退社時刻" && col.out == null) col.out = c;   // 最初の「退社時刻」
        else if (t === "休憩時間" && col.brk == null) col.brk = c;
        else if (t === "勤務時間" && col.work == null) col.work = c;
        else if (t === "深夜休憩時間" && col.nightBrk == null) col.nightBrk = c;
        else if (t === "深夜労働" && col.night == null) col.night = c;
      }
      break;
    }
  }
  if (hRow < 0 || col.day == null || col.in == null || col.out == null || col.brk == null) {
    throw new Error("月次フォーマットのヘッダー（日にち/出社時刻/退社時刻/休憩時間）を特定できません");
  }

  // 2) 年・月を特定（ヘッダー行より上で「月」ラベルの左隣＝月、2000-2100の数値＝年）
  var year = null, month = null;
  for (var r2 = 0; r2 < hRow && month == null; r2++) {
    for (var c2 = 0; c2 < data[r2].length; c2++) {
      var v = data[r2][c2];
      if (String(v).replace(/\s/g, "") === "月" && c2 > 0) {
        var mv = Number(data[r2][c2 - 1]);
        if (mv >= 1 && mv <= 12) month = mv;
      }
      var yv = Number(v);
      if (yv >= 2000 && yv <= 2100) year = yv;
    }
  }
  if (!year || !month) throw new Error("シートの年・月を特定できません（年と『月』欄を確認）");
  var mk = year + "-" + ("0" + month).slice(-2);

  // 3) 「日にち」→ 行番号(0始まり) のマップ
  var dayRow = {};
  for (var r3 = hRow + 1; r3 < data.length; r3++) {
    var dv = Number(data[r3][col.day]);
    if (dv >= 1 && dv <= 31) dayRow[dv] = r3;
  }

  // 4) 当月の勤務セッションを日別に集計
  var agg = {};
  sessions.forEach(function (s) {
    if (isBreak_(s) || !s.end) return;
    var dk = dayKey_(s.start);
    if (dk.substring(0, 7) !== mk) return;
    var d = Number(dk.substring(8, 10));
    var a = agg[d] || (agg[d] = { firstIn: null, lastOut: null, work: 0, night: 0 });
    var st = new Date(s.start), en = new Date(s.end);
    if (!a.firstIn || st < a.firstIn) a.firstIn = st;
    if (!a.lastOut || en > a.lastOut) a.lastOut = en;
    a.work += durMin_(s);
    a.night += nightMin_(s, settings);
  });

  // 5) 書き込み
  var written = 0;
  Object.keys(dayRow).forEach(function (dStr) {
    var d = Number(dStr);
    var row = dayRow[d] + 1; // 1始まり
    var a = agg[d];
    if (a) {
      sh.getRange(row, col.in + 1).setValue(hhmm_(a.firstIn));
      sh.getRange(row, col.out + 1).setValue(hhmm_(a.lastOut));
      var spanMin = (a.lastOut - a.firstIn) / 60000;
      var brkMin = Math.max(0, spanMin - a.work);
      sh.getRange(row, col.brk + 1).setValue(fmtHM_(brkMin));
      if (CONFIG.WRITE_WORK && col.work != null) {
        sh.getRange(row, col.work + 1).setValue(fmtHM_(a.work)); // 勤務時間 = 実働合計
      }
      if (CONFIG.WRITE_WORK && col.nightBrk != null) {
        // 深夜帯の空き時間 = 深夜帯に重なった全体時間 − 深夜労働
        var nightSpan = nightMin_({ start: a.firstIn, end: a.lastOut }, settings);
        var nightBrkMin = Math.max(0, nightSpan - a.night);
        sh.getRange(row, col.nightBrk + 1).setValue(nightBrkMin > 0 ? fmtHM_(nightBrkMin) : "");
      }
      if (CONFIG.WRITE_NIGHT && col.night != null) {
        sh.getRange(row, col.night + 1).setValue(a.night > 0 ? fmtHM_(a.night) : "");
      }
      written++;
    } else if (CONFIG.CLEAR_EMPTY) {
      sh.getRange(row, col.in + 1).clearContent();
      sh.getRange(row, col.out + 1).clearContent();
      sh.getRange(row, col.brk + 1).clearContent();
      if (CONFIG.WRITE_WORK && col.work != null) sh.getRange(row, col.work + 1).clearContent();
      if (CONFIG.WRITE_WORK && col.nightBrk != null) sh.getRange(row, col.nightBrk + 1).clearContent();
    }
  });

  return { written: written, month: mk };
}

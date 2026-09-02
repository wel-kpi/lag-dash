/**
 * LAGNET 獲得管理ツール - 共有データ同期用 Webアプリ (v3)
 *
 * v3の変更点: ツール側を「自動同期」から「受信・送信ボタンによる手動同期」に
 * 変更したのに合わせて、以下の2点を変更しました。
 *
 *   1. 送信（POST）を、ブラウザのfetch()ではなく「隠しiframe + フォーム送信」
 *      方式に変更したため、リクエストの中身が JSON の生ボディではなく
 *      フォームの1フィールド（payload）として届くようになりました。
 *      doPost側で両方の形式を受け付けるようにしています。
 *   2. 受信（GET）を、Apps Script経由ではなく、Googleスプレッドシート公式の
 *      gviz機能で直接読み取る方式に変更しました。そのため、
 *      このスプレッドシートを「リンクを知っている全員が閲覧可」に
 *      共有設定してください（データ本体は今まで通りこのApps Scriptの
 *      URLを知っている人にしか書き込めません。閲覧のみ誰でも可能になります）。
 *      ※共有設定変更手順: スプレッドシート右上の「共有」→
 *        「リンクを知っている全員」→ 閲覧者 に設定
 *
 * doGet はそのまま残していますが（後方互換のため）、ツール側は今後
 * doGetを使わず、直接gvizを読みに行きます。
 *
 * ============================
 * 移行手順（v1/v2から使っている場合）
 * ============================
 * 1. Apps Script エディタを開き、中身を全部このコードに置き換える
 *    （「AppData」シートの中身は消さなくてOKです。スキーマは変わっていません）
 * 2. 保存
 * 3. 「デプロイ」→「デプロイを管理」→ 既存のデプロイの鉛筆アイコン（編集）
 *    → バージョン「新しいバージョン」を選択 → デプロイ
 *    ※「新しいデプロイ」ではなく「デプロイを管理」から更新することで、
 *      Webアプリの URL は変わりません（ツール側の設定を変える必要なし）
 * 4. スプレッドシート右上の「共有」→「リンクを知っている全員」→
 *    閲覧者、に設定する（gviz受信のために必要）
 */

const SHEET_NAME = 'AppData';
const CHUNK_SIZE = 40000; // 50,000文字制限に対して安全マージンを確保

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['key', 'chunkIndex', 'value', 'updatedAt']);
  }
  return sheet;
}

function jsonOutput_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    const grouped = {}; // key -> [{idx, chunk}]
    for (let i = 1; i < data.length; i++) {
      const key = data[i][0];
      if (key === '' || key === undefined || key === null) continue;
      const idx = Number(data[i][1]) || 0;
      const chunk = data[i][2] != null ? String(data[i][2]) : '';
      if (!grouped[key]) grouped[key] = [];
      grouped[key].push({ idx: idx, chunk: chunk });
    }
    const result = {};
    Object.keys(grouped).forEach(function (key) {
      const sorted = grouped[key].sort(function (a, b) { return a.idx - b.idx; });
      const json = sorted.map(function (x) { return x.chunk; }).join('');
      try {
        result[key] = JSON.parse(json);
      } catch (err) {
        result[key] = null;
      }
    });
    return jsonOutput_(result);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    let payload;
    // ツール側は「隠しiframe + フォーム送信」方式（フォームの1フィールド
    // payload）で送ってくるが、後方互換のため生JSONボディも引き続き受け付ける。
    if (e.parameter && e.parameter.payload) {
      try {
        payload = JSON.parse(e.parameter.payload);
      } catch (err) {
        return jsonOutput_({ error: 'invalid JSON in payload field' });
      }
    } else if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        return jsonOutput_({ error: 'invalid JSON body' });
      }
    } else {
      return jsonOutput_({ error: 'no payload received' });
    }

    const key = payload.key;
    const value = payload.value;
    if (!key) {
      return jsonOutput_({ error: 'key is required' });
    }

    const json = JSON.stringify(value);
    const chunks = [];
    if (json.length === 0) {
      chunks.push('');
    } else {
      for (let i = 0; i < json.length; i += CHUNK_SIZE) {
        chunks.push(json.slice(i, i + CHUNK_SIZE));
      }
    }

    const sheet = getSheet_();
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    // remove all existing rows for this key first (bottom-to-top so row
    // indices don't shift while deleting), then append fresh chunk rows
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === key) {
        sheet.deleteRow(i + 1);
      }
    }

    const rows = chunks.map(function (chunk, idx) {
      return [key, idx, chunk, now];
    });
    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 4).setValues(rows);
    }

    return jsonOutput_({ ok: true, chunks: chunks.length });
  } catch (err) {
    return jsonOutput_({ error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

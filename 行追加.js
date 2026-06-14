/*************************************************
 * Google スプレッドシート用（アンダースコア区切り版）
 * A列のチェックボックスをONにすると、
 * B列の番号をもとに子行を下へ追加する
 *
 * 例:
 * 1
 * 1_1
 * 1_1_1
 *************************************************/

const SHEET_NAME = 'wbs';   // 対象シート名を固定したい場合は 'Sheet1' など。空ならアクティブシート
const START_ROW = 3;     // データ開始行
const CHECKBOX_COL = 1;  // A列
const NUMBER_COL = 2;    // B列
const SEPARATOR = '_';   // 区切り文字

/**
 * シートを開いたときに初期設定
 */
function onOpen() {
  setupSheet();
}

/**
 * 初期設定
 * - A2:A にチェックボックス
 * - B2:B をプレーンテキスト化
 */
function setupSheet() {
  const sheet = getTargetSheet_();
  const maxRows = sheet.getMaxRows();

  // A列にチェックボックス
  sheet.getRange(START_ROW, CHECKBOX_COL, maxRows - START_ROW + 1, 1).insertCheckboxes();

  // B列をプレーンテキストにする
  sheet.getRange(START_ROW, NUMBER_COL, maxRows - START_ROW + 1, 1).setNumberFormat('@');
}

/**
 * A列のチェックボックス変更時に実行
 */
function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (SHEET_NAME && sheet.getName() !== SHEET_NAME) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  const value = e.value;

  // A2以下、かつチェックがON(TRUE)のときのみ
  if (col !== CHECKBOX_COL || row < START_ROW || value !== 'TRUE') return;

  try {
    // B列は表示文字列を取得して正規化
    const parentCode = normalizeCode_(sheet.getRange(row, NUMBER_COL).getDisplayValue());

    if (!isValidCode_(parentCode)) {
      sheet.getRange(row, CHECKBOX_COL).setValue(false);
      SpreadsheetApp.getActive().toast(
        'B列の番号形式が不正です。例: 1 / 1_2 / 1_2_3'
      );
      return;
    }

    const result = getInsertPositionAndNewCode_(sheet, row, parentCode);

    // 親配下の一番下に1行追加
    sheet.insertRowsAfter(result.insertAfterRow, 1);
    const newRow = result.insertAfterRow + 1;

    const lastCol = Math.max(sheet.getLastColumn(), NUMBER_COL);

    // 元行の書式をコピー
    sheet.getRange(row, 1, 1, lastCol).copyTo(
      sheet.getRange(newRow, 1, 1, lastCol),
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );

    // 新しい行のB列をプレーンテキスト化して番号設定
    sheet.getRange(newRow, NUMBER_COL).setNumberFormat('@');
    sheet.getRange(newRow, NUMBER_COL).setValue(result.newCode);

    // 新しい行のA列にチェックボックスを付与
    sheet.getRange(newRow, CHECKBOX_COL).insertCheckboxes().setValue(false);

    // 元のチェックをOFFへ戻す
    sheet.getRange(row, CHECKBOX_COL).setValue(false);

  } catch (err) {
    sheet.getRange(row, CHECKBOX_COL).setValue(false);
    SpreadsheetApp.getActive().toast('エラー: ' + err.message);
  }
}

/**
 * 親コードの配下の一番下の行位置と、
 * 次に採番すべき直下の子コードを求める
 *
 * 例:
 * 親: 1
 * 既存: 1_1, 1_1_1, 1_2
 * => 挿入位置: 1_2 の下
 * => 新コード: 1_3
 */
function getInsertPositionAndNewCode_(sheet, parentRow, parentCode) {
  const lastRow = sheet.getLastRow();
  const parentDepth = parentCode.split(SEPARATOR).length;

  let insertAfterRow = parentRow;
  let maxDirectChildNo = 0;

  for (let r = parentRow + 1; r <= lastRow; r++) {
    const code = normalizeCode_(sheet.getRange(r, NUMBER_COL).getDisplayValue());

    // 空なら配下終了
    if (!code) break;

    // 親コードの子孫でなければ配下終了
    if (!code.startsWith(parentCode + SEPARATOR)) {
      break;
    }

    // 子孫である間は、この行が「配下の一番下」
    insertAfterRow = r;

    // 直下の子だけを数える
    const parts = code.split(SEPARATOR);
    if (parts.length === parentDepth + 1) {
      const prefix = parts.slice(0, parentDepth).join(SEPARATOR);
      if (prefix === parentCode) {
        const childNo = Number(parts[parentDepth]);
        if (Number.isInteger(childNo) && childNo > maxDirectChildNo) {
          maxDirectChildNo = childNo;
        }
      }
    }
  }

  const newCode = `${parentCode}${SEPARATOR}${maxDirectChildNo + 1}`;

  return {
    insertAfterRow,
    newCode
  };
}

/**
 * コード文字列を正規化
 * - 前後空白除去
 * - 全角アンダースコアを半角へ統一
 * - 旧データのハイフンもアンダースコアに統一
 * - 余分な空白除去
 */
function normalizeCode_(value) {
  return String(value || '')
    .trim()
    .replace(/＿/g, '_')                    // 全角アンダースコア → 半角
    .replace(/[‐‑‒–—―ー－-]/g, '_')         // ハイフン類 → アンダースコア
    .replace(/\s+/g, '');                  // 空白削除
}

/**
 * コード形式チェック
 * 例: 1 / 1_2 / 1_2_3
 */
function isValidCode_(code) {
  return /^\d+(?:_\d+)*$/.test(code);
}

/**
 * 対象シート取得
 */
function getTargetSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return SHEET_NAME ? ss.getSheetByName(SHEET_NAME) : ss.getActiveSheet();
}

/**
 * 既存のB列データをアンダースコア形式に統一したいとき用
 * 必要なときだけ手動実行してください
 */
function convertExistingNumbersToUnderscore() {
  const sheet = getTargetSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < START_ROW) return;

  const range = sheet.getRange(START_ROW, NUMBER_COL, lastRow - START_ROW + 1, 1);
  const values = range.getDisplayValues();

  const converted = values.map(([v]) => [normalizeCode_(v)]);
  range.setNumberFormat('@');
  range.setValues(converted);
}

/**
 * 設定用シートのentryIdセルが「entry.」プレフィックスなしで
 * 数字だけ入力されていても事前入力URLが機能するよう補完する
 */
function normalizeEntryId(entryId) {
  const idStr = String(entryId).trim();
  return idStr.startsWith('entry.') ? idStr : `entry.${idStr}`;
}

// 新しい回答の通知メール送信先
const NOTIFICATION_EMAIL = 'staff-global@tomonokai-corp.com';

/**
 * 作成したフォームに新しい回答が来るたびに、NOTIFICATION_EMAIL宛てに通知メールを送る。
 * createFormLink()内でフォームごとに自動でトリガー登録される。
 */
function notifyNewFormResponse(e) {
  const trigger = ScriptApp.getProjectTriggers().find(t => t.getUniqueId() === e.triggerUid);
  const formId = trigger ? trigger.getTriggerSourceId() : null;
  const form = formId ? FormApp.openById(formId) : null;
  const formTitle = form ? form.getTitle() : 'フォーム';
  const editUrl = form ? form.getEditUrl() : '';

  MailApp.sendEmail({
    to: NOTIFICATION_EMAIL,
    subject: `【新しい回答】${formTitle}`,
    body: `フォーム「${formTitle}」に新しい回答がありました。\n\n回答を確認する: ${editUrl}`
  });
}

/**
 * スプレッドシートのメニューを追加する関数
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フォーム自動生成')
    .addItem('リンクを一括作成', 'createAllForms')
    .addToUi();
}

/**
 * 設定用シートから、sectionLabel（「応募フォーム」「同意書」「詳細レス」）の
 * ブロックを検索して読み込む。
 *
 * ブロックの並び（sectionLabelの行を起点とした相対位置）:
 *   sectionLabel行
 *   sectionLabel行+1: フォルダID | テンプレート名1 | テンプレート名2 | ...
 *   sectionLabel行+2: <folderId> | <formId1> | <formId2> | ...
 *   sectionLabel行+3: (空欄)     | <entryId1> | <entryId2> | ...
 */
function loadFormTypeConfig(settingSheet, sectionLabel) {
  const data = settingSheet.getDataRange().getValues();
  let labelRow = -1;

  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === sectionLabel) {
      labelRow = r;
      break;
    }
  }
  if (labelRow === -1) {
    throw new Error(`設定用シートに「${sectionLabel}」のブロックが見つかりません。`);
  }

  const headerRow = data[labelRow + 1];
  const dataRow = data[labelRow + 2];
  const entryRow = data[labelRow + 3];

  const folderId = dataRow[0];
  const templates = {};

  for (let c = 1; c < headerRow.length; c++) {
    const templateName = headerRow[c];
    const formId = dataRow[c];
    const entryId = entryRow[c];
    if (templateName && formId && entryId) {
      templates[templateName] = { formId: formId, entryId: entryId };
    }
  }

  return { folderId: folderId, templates: templates };
}

/**
 * テンプレートフォームをコピーして公開設定・共有設定を行い、
 * 差し込み前URLと案件番号差し込み済みURLを返す
 */
function createFormLink(templateFormId, entryId, parentFolder, fileName, formTitle, caseNumber, rowLabel) {
  const newFormFile = DriveApp.getFileById(templateFormId).makeCopy(fileName, parentFolder);
  const newForm = FormApp.openById(newFormFile.getId());

  try {
    newForm.setPublished(true);
  } catch (err) {
    throw new Error(`${rowLabel} setPublished失敗: ${err.message}`);
  }

  try {
    newForm.setAcceptingResponses(true);
  } catch (err) {
    throw new Error(`${rowLabel} setAcceptingResponses失敗: ${err.message}`);
  }

  try {
    DriveApp.getFileById(newForm.getId()).setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.VIEW
    );
  } catch (err) {
    throw new Error(`${rowLabel} setSharing失敗: ${err.message}`);
  }

  if (formTitle) {
    newForm.setTitle(formTitle);
  }

  try {
    ScriptApp.newTrigger('notifyNewFormResponse')
      .forForm(newForm)
      .onFormSubmit()
      .create();
  } catch (err) {
    Logger.log(`${rowLabel} 通知トリガー設定失敗: ${err.message}`);
  }

  let viewUrl;
  try {
    if (newForm.supportsAdvancedResponderPermissions()) {
      viewUrl = newForm.getPublishedUrl();
    } else {
      viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
    }
  } catch (err) {
    Logger.log(`${rowLabel} getPublishedUrl失敗のためgetEditUrlにフォールバック: ${err.message}`);
    viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
  }

  let prefilledUrl = '';
  if (caseNumber) {
    prefilledUrl = `${viewUrl.replace('/viewform', '/formResponse')}?${normalizeEntryId(entryId)}=${encodeURIComponent(caseNumber)}`;
  }

  return { viewUrl: viewUrl, prefilledUrl: prefilledUrl };
}

// =================================================================
// 「フォーム作成シート」の1行ごとに、応募フォーム・同意書・詳細レスの
// リンクをまとめて作成する
//
// 列構成（フォーム作成シート、データは3行目から）:
//   A 案件番号
//   B 応募フォームのテンプレート種別
//   C 同意書のテンプレート種別
//   D 募集エリア（未使用）    E 実施日（VLOOKUP・未使用）   F 学校名（VLOOKUP・未使用）
//   G 応募フォーム 差し込み済みリンク（出力）
//   H 同意書 差し込み済みリンク（出力）
//   I 詳細レス 差し込み済みリンク（出力）
//   J （未使用・区切り）
//   K 応募フォームのファイル名（内部用）   L 応募フォームのタイトル（公開用）
//   M 同意書のファイル名（内部用）         N 同意書のタイトル（公開用）
//   O 詳細レスのファイル名（内部用）       P 詳細レスのタイトル（公開用）
//   Q 応募フォーム 差し込み前URL（出力）
//   R 同意書 差し込み前URL（出力）
//   S 詳細レス 差し込み前URL（出力）
// =================================================================
function createAllForms() {
  const settingSheetName = '設定用シート';
  const targetSheetName = 'フォーム作成シート';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingSheet = ss.getSheetByName(settingSheetName);
  const targetSheet = ss.getSheetByName(targetSheetName);
  const ui = SpreadsheetApp.getUi();

  if (!settingSheet) {
    ui.alert(`「${settingSheetName}」が見つかりません。`);
    return;
  }
  if (!targetSheet) {
    ui.alert(`「${targetSheetName}」が見つかりません。`);
    return;
  }

  const headerCheck = targetSheet.getRange(2, 1).getValue();
  if (headerCheck !== '案件番号') {
    ui.alert(`「${targetSheetName}」の2行目1列目が「案件番号」ではありません。シート構成が想定と異なる可能性があるため処理を中止しました。`);
    return;
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 3) {
    ui.alert('入力データがありません。');
    return;
  }

  const COL = {
    caseNumber: 1, appType: 2, consentType: 3, area: 4,
    appLink: 7, consentLink: 8, detailLink: 9,
    appFileName: 11, appTitle: 12,
    consentFileName: 13, consentTitle: 14,
    detailFileName: 15, detailTitle: 16,
    appPlainUrl: 17, consentPlainUrl: 18, detailPlainUrl: 19
  };

  try {
    const appConfig = loadFormTypeConfig(settingSheet, '応募フォーム');
    const consentConfig = loadFormTypeConfig(settingSheet, '同意書');
    const detailConfig = loadFormTypeConfig(settingSheet, '詳細レス');

    const appFolder = DriveApp.getFolderById(appConfig.folderId);
    const consentFolder = DriveApp.getFolderById(consentConfig.folderId);
    const detailFolder = DriveApp.getFolderById(detailConfig.folderId);

    // 詳細レスはテンプレートが1種類のみ
    const detailTemplateKeys = Object.keys(detailConfig.templates);
    const detailTemplate = detailTemplateKeys.length > 0 ? detailConfig.templates[detailTemplateKeys[0]] : null;

    const numRows = lastRow - 2; // 1〜2行目はヘッダー
    const values = targetSheet.getRange(3, 1, numRows, 19).getValues();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const caseNumber = row[COL.caseNumber - 1];
      const appType = row[COL.appType - 1];
      const consentType = row[COL.consentType - 1];
      const area = row[COL.area - 1];
      // A〜D列（案件番号・応募フォームのテンプレート種別・同意書のテンプレート種別・募集エリア）が
      // すべて入力されている行のみ対象とする
      if (!caseNumber || !appType || !consentType || !area) continue;

      const rowNum = i + 3;
      const rowLabel = `[${rowNum}行目 案件${caseNumber}]`;

      let appLink = row[COL.appLink - 1];
      let appPlainUrl = row[COL.appPlainUrl - 1];
      let consentLink = row[COL.consentLink - 1];
      let consentPlainUrl = row[COL.consentPlainUrl - 1];
      let detailLink = row[COL.detailLink - 1];
      let detailPlainUrl = row[COL.detailPlainUrl - 1];

      // ① 応募フォーム
      if (!appLink) {
        const templateType = row[COL.appType - 1];
        const template = appConfig.templates[templateType];
        const fileName = row[COL.appFileName - 1];
        const title = row[COL.appTitle - 1];
        if (template && fileName) {
          const result = createFormLink(
            template.formId, template.entryId, appFolder, fileName, title, caseNumber, `${rowLabel}(応募フォーム)`
          );
          appPlainUrl = result.viewUrl;
          appLink = result.prefilledUrl;
        }
      }

      // ② 同意書
      if (!consentLink) {
        const templateType = row[COL.consentType - 1];
        const template = consentConfig.templates[templateType];
        const fileName = row[COL.consentFileName - 1];
        const title = row[COL.consentTitle - 1];
        if (template && fileName) {
          const result = createFormLink(
            template.formId, template.entryId, consentFolder, fileName, title, caseNumber, `${rowLabel}(同意書)`
          );
          consentPlainUrl = result.viewUrl;
          consentLink = result.prefilledUrl;
        }
      }

      // ③ 詳細レス（テンプレートは1種類のみ、選択列なし）
      if (!detailLink && detailTemplate) {
        const fileName = row[COL.detailFileName - 1];
        const title = row[COL.detailTitle - 1];
        if (fileName) {
          const result = createFormLink(
            detailTemplate.formId, detailTemplate.entryId, detailFolder, fileName, title, caseNumber, `${rowLabel}(詳細レス)`
          );
          detailPlainUrl = result.viewUrl;
          detailLink = result.prefilledUrl;
        }
      }

      targetSheet.getRange(rowNum, COL.appLink, 1, 3).setValues([[appLink, consentLink, detailLink]]);
      targetSheet.getRange(rowNum, COL.appPlainUrl, 1, 3).setValues([[appPlainUrl, consentPlainUrl, detailPlainUrl]]);
    }

    ui.alert('フォームの一括作成処理が完了しました。');
  } catch (e) {
    ui.alert(`エラー: ${e.message}`);
  }
}

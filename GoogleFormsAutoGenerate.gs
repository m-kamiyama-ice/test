/**
 * 設定用シートのentryIdセルが「entry.」プレフィックスなしで
 * 数字だけ入力されていても事前入力URLが機能するよう補完する
 */
function normalizeEntryId(entryId) {
  const idStr = String(entryId).trim();
  return idStr.startsWith('entry.') ? idStr : `entry.${idStr}`;
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
 * 設定用シートから「テンプレート名 → {formId, entryId}」のマップを読み込む
 */
function loadTemplateMap(settingSheet, nameRow, nameStartCol, nameEndCol, formIdRow, entryIdRow) {
  const templateNames = settingSheet.getRange(nameRow, nameStartCol, 1, nameEndCol - nameStartCol + 1).getValues()[0];
  const templates = {};

  for (let i = 0; i < templateNames.length; i++) {
    const col = nameStartCol + i;
    const formId = settingSheet.getRange(formIdRow, col).getValue();
    const entryId = settingSheet.getRange(entryIdRow, col).getValue();
    if (formId && entryId && templateNames[i]) {
      templates[templateNames[i]] = { formId: formId, entryId: entryId };
    }
  }

  return templates;
}

/**
 * テンプレートフォームをコピーして公開設定・共有設定を行い、
 * 表示用URLと案件番号付きURLを返す
 */
function createFormLink(templateFormId, entryId, parentFolder, caseName, caseNumber, rowLabel) {
  const newFormFile = DriveApp.getFileById(templateFormId).makeCopy(caseName, parentFolder);
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

  newForm.setTitle(caseName);

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
// 「案件管理」シートの1行ごとに、応募フォーム・同意書・詳細レスの
// リンクをまとめて作成する
//
// 列構成:
//   A 案件管理番号
//   B 案件名
//   C 食事制限の有無（応募フォームのテンプレート選択）
//   D ブリーフィングの有無（同意書のテンプレート選択）
//   E 応募フォームURL       F 応募フォーム案件番号付きURL
//   G 同意書URL             H 同意書案件番号付きURL
//   I 詳細レスURL           J 詳細レス案件番号付きURL
// =================================================================
function createAllForms() {
  const settingSheetName = '設定用シート';
  const targetSheetName = '案件管理';
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

  const lastRow = targetSheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('入力データがありません。');
    return;
  }

  const COL = {
    caseNumber: 1, caseName: 2, mealType: 3, briefingType: 4,
    appFormUrl: 5, appFormPrefilled: 6,
    consentUrl: 7, consentPrefilled: 8,
    detailUrl: 9, detailPrefilled: 10
  };

  try {
    const appFolder = DriveApp.getFolderById(settingSheet.getRange('A3').getValue());
    const appTemplates = loadTemplateMap(settingSheet, 2, 2, 4, 3, 4); // B2:D2 / row3=formId / row4=entryId

    const consentFolder = DriveApp.getFolderById(settingSheet.getRange('A8').getValue());
    const consentTemplates = loadTemplateMap(settingSheet, 7, 2, 3, 8, 9); // B7:C7 / row8=formId / row9=entryId

    const detailFolder = DriveApp.getFolderById(settingSheet.getRange('A13').getValue());
    const detailFormId = settingSheet.getRange('B13').getValue();
    const detailEntryId = settingSheet.getRange('B14').getValue();

    const numRows = lastRow - 1;
    const values = targetSheet.getRange(2, 1, numRows, 10).getValues();

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const caseNumber = row[COL.caseNumber - 1];
      const caseName = row[COL.caseName - 1];
      if (!caseName) continue;

      const rowNum = i + 2;
      const rowLabel = `[${rowNum}行目 ${caseName}]`;

      let appFormUrl = row[COL.appFormUrl - 1];
      let appFormPrefilled = row[COL.appFormPrefilled - 1];
      let consentUrl = row[COL.consentUrl - 1];
      let consentPrefilled = row[COL.consentPrefilled - 1];
      let detailUrl = row[COL.detailUrl - 1];
      let detailPrefilled = row[COL.detailPrefilled - 1];

      // ① 応募フォーム（食事制限の有無でテンプレートを出し分け）
      if (!appFormUrl) {
        const mealType = row[COL.mealType - 1];
        const template = appTemplates[mealType];
        if (template) {
          const result = createFormLink(
            template.formId, template.entryId, appFolder, caseName, caseNumber, `${rowLabel}(応募フォーム)`
          );
          appFormUrl = result.viewUrl;
          appFormPrefilled = result.prefilledUrl;
        }
      }

      // ② 同意書（ブリーフィングの有無でテンプレートを出し分け）
      if (!consentUrl) {
        const briefingType = row[COL.briefingType - 1];
        const template = consentTemplates[briefingType];
        if (template) {
          const result = createFormLink(
            template.formId, template.entryId, consentFolder, caseName, caseNumber, `${rowLabel}(同意書)`
          );
          consentUrl = result.viewUrl;
          consentPrefilled = result.prefilledUrl;
        }
      }

      // ③ 詳細レス（テンプレートは1種類のみ、選択列なし）
      if (!detailUrl && detailFormId && detailEntryId) {
        const result = createFormLink(
          detailFormId, detailEntryId, detailFolder, caseName, caseNumber, `${rowLabel}(詳細レス)`
        );
        detailUrl = result.viewUrl;
        detailPrefilled = result.prefilledUrl;
      }

      targetSheet.getRange(rowNum, COL.appFormUrl, 1, 6).setValues([[
        appFormUrl, appFormPrefilled, consentUrl, consentPrefilled, detailUrl, detailPrefilled
      ]]);
    }

    ui.alert('フォームの一括作成処理が完了しました。');
  } catch (e) {
    ui.alert(`エラー: ${e.message}`);
  }
}

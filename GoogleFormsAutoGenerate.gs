/**
 * スプレッドシートのメニューに3つの項目を追加する関数
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('フォーム自動生成')
    .addItem('【応募フォーム】を作成', 'createApplicationForms')
    .addSeparator()
    .addItem('【同意書】を作成', 'createConsentForms')
    .addItem('【詳細レス】を作成', 'createDetailForms')
    .addToUi();
}

// =================================================================
// 1. 応募フォーム（3テンプレート）作成用スクリプト
// =================================================================
function createApplicationForms() {
  const settingSheetName = '設定用シート';
  const targetSheetName = '応募フォーム';
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

  try {
    const folderId = settingSheet.getRange('A3').getValue();
    const templateNames = settingSheet.getRange('B2:D2').getValues()[0];
    const templates = {};

    for (let i = 0; i < templateNames.length; i++) {
      const col = i + 2;
      const formId = settingSheet.getRange(3, col).getValue();
      const entryId = settingSheet.getRange(4, col).getValue();
      if (formId && entryId && templateNames[i]) {
        templates[templateNames[i]] = { formId: formId, entryId: entryId };
      }
    }

    const dataRange = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 8);
    const values = dataRange.getValues();
    const parentFolder = DriveApp.getFolderById(folderId);

    for (let i = 0; i < values.length; i++) {
      const caseNumber = values[i][0];
      const caseName = values[i][1];
      const formTitle = values[i][2];
      const templateType = values[i][3];
      let formUrl = values[i][6]; // G列（フォームURL）
      let prefilledUrlValue = values[i][7]; // H列（案件番号付きURL）

      const templateInfo = templates[templateType];

      if (caseName && templateInfo && !formUrl) {
        const newFormFile = DriveApp.getFileById(templateInfo.formId).makeCopy(caseName, parentFolder);
        const newForm = FormApp.openById(newFormFile.getId());

        newForm.setAcceptingResponses(true);
        newForm.setPublished(true);

        DriveApp.getFileById(newForm.getId()).setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );

        if (formTitle) {
          newForm.setTitle(formTitle);
        }

        let viewUrl;

        try {
          if (newForm.supportsAdvancedResponderPermissions()) {
            viewUrl = newForm.getPublishedUrl();
          } else {
            viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
          }
        } catch (err) {
          viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
        }

        values[i][6] = viewUrl;

        if (caseNumber) {
          const prefilledUrl = `${viewUrl.replace('/viewform', '/formResponse')}?${templateInfo.entryId}=${encodeURIComponent(caseNumber)}`;
          values[i][7] = prefilledUrl;
        }
      }
    }

    const outputValues = values.map(row => [row[6], row[7]]);
    targetSheet.getRange(2, 7, outputValues.length, 2).setValues(outputValues);

    ui.alert('【応募フォーム】の作成処理が完了しました。');
  } catch (e) {
    ui.alert(`エラー(応募フォーム): ${e.message}`);
  }
}

// =================================================================
// 2. 同意書（2テンプレート）作成用スクリプト
// =================================================================
function createConsentForms() {
  const settingSheetName = '設定用シート';
  const targetSheetName = '同意書';
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

  try {
    const folderId = settingSheet.getRange('A8').getValue();
    const templateNames = settingSheet.getRange('B7:C7').getValues()[0];
    const templates = {};

    for (let i = 0; i < templateNames.length; i++) {
      const col = i + 2;
      const formId = settingSheet.getRange(8, col).getValue();
      const entryId = settingSheet.getRange(9, col).getValue();
      if (formId && entryId && templateNames[i]) {
        templates[templateNames[i]] = { formId: formId, entryId: entryId };
      }
    }

    const dataRange = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 8);
    const values = dataRange.getValues();
    const parentFolder = DriveApp.getFolderById(folderId);

    for (let i = 0; i < values.length; i++) {
      const caseNumber = values[i][0];
      const caseName = values[i][1];
      const formTitle = values[i][2];
      const templateType = values[i][3];
      let formUrl = values[i][6]; // G列
      let prefilledUrlValue = values[i][7]; // H列

      const templateInfo = templates[templateType];

      if (caseName && templateInfo && !formUrl) {
        const newFormFile = DriveApp.getFileById(templateInfo.formId).makeCopy(caseName, parentFolder);
        const newForm = FormApp.openById(newFormFile.getId());

        newForm.setAcceptingResponses(true);
        newForm.setPublished(true);

        DriveApp.getFileById(newForm.getId()).setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );

        if (formTitle) {
          newForm.setTitle(formTitle);
        }

        let viewUrl;

        try {
          if (newForm.supportsAdvancedResponderPermissions()) {
            viewUrl = newForm.getPublishedUrl();
          } else {
            viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
          }
        } catch (err) {
          viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
        }

        values[i][6] = viewUrl;

        if (caseNumber) {
          const prefilledUrl = `${viewUrl.replace('/viewform', '/formResponse')}?${templateInfo.entryId}=${encodeURIComponent(caseNumber)}`;
          values[i][7] = prefilledUrl;
        }
      }
    }

    const outputValues = values.map(row => [row[6], row[7]]);
    targetSheet.getRange(2, 7, outputValues.length, 2).setValues(outputValues);

    ui.alert('【同意書】の作成処理が完了しました。');
  } catch (e) {
    ui.alert(`エラー(同意書): ${e.message}`);
  }
}

// =================================================================
// 3. 詳細レス（1テンプレート）作成用スクリプト
// =================================================================
function createDetailForms() {
  const settingSheetName = '設定用シート';
  const targetSheetName = '詳細レス';
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

  try {
    const folderId = settingSheet.getRange('A13').getValue();
    const templateFormId = settingSheet.getRange('B13').getValue();
    const templateEntryId = settingSheet.getRange('B14').getValue();

    const dataRange = targetSheet.getRange(2, 1, targetSheet.getLastRow() - 1, 7);
    const values = dataRange.getValues();
    const parentFolder = DriveApp.getFolderById(folderId);

    for (let i = 0; i < values.length; i++) {
      const caseNumber = values[i][0];
      const caseName = values[i][1];
      const formTitle = values[i][2];
      let formUrl = values[i][5]; // F列
      let prefilledUrlValue = values[i][6]; // G列

      if (caseName && !formUrl) {
        const newFormFile = DriveApp.getFileById(templateFormId).makeCopy(caseName, parentFolder);
        const newForm = FormApp.openById(newFormFile.getId());

        newForm.setAcceptingResponses(true);
        newForm.setPublished(true);

        DriveApp.getFileById(newForm.getId()).setSharing(
          DriveApp.Access.ANYONE_WITH_LINK,
          DriveApp.Permission.VIEW
        );

        if (formTitle) {
          newForm.setTitle(formTitle);
        }

        let viewUrl;

        try {
          if (newForm.supportsAdvancedResponderPermissions()) {
            viewUrl = newForm.getPublishedUrl();
          } else {
            viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
          }
        } catch (err) {
          viewUrl = newForm.getEditUrl().replace('/edit', '/viewform');
        }

        values[i][5] = viewUrl;

        if (caseNumber) {
          const prefilledUrl = `${viewUrl.replace('/viewform', '/formResponse')}?${templateEntryId}=${encodeURIComponent(caseNumber)}`;
          values[i][6] = prefilledUrl;
        }
      }
    }

    const outputValues = values.map(row => [row[5], row[6]]);
    targetSheet.getRange(2, 6, outputValues.length, 2).setValues(outputValues);

    ui.alert('【詳細レス】の作成処理が完了しました。');
  } catch (e) {
    ui.alert(`エラー(詳細レス): ${e.message}`);
  }
}

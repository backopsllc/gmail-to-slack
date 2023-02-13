const PROPERTY_PREFIX = 'gmail-to-slack.';

export interface SpreadSheetService {
  getUserProperty: (key: string) => string;
  setUserProperty: (key: string, value: string) => void;
  getSheetByName: (sheetName: string) => GoogleAppsScript.Spreadsheet.Sheet;
  clearSheet: (sheet: GoogleAppsScript.Spreadsheet.Sheet) => void;
  getRange: (
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    row?: number,
    column?: number,
    numRows?: number,
    numColms?: number
  ) => GoogleAppsScript.Spreadsheet.Range;
  getValues: (range: GoogleAppsScript.Spreadsheet.Range) => any[][];
  setValues: (
    range: GoogleAppsScript.Spreadsheet.Range,
    values: any[][]
  ) => GoogleAppsScript.Spreadsheet.Range;
  showMessage: (title: string, message: string) => void;
}

export class SpreadSheetServiceImpl implements SpreadSheetService {
  public getUserProperty(key: string): string {
    const value =
      PropertiesService.getUserProperties().getProperty(
        PROPERTY_PREFIX + key
      ) || '';
    return value;
  }

  public setUserProperty(key: string, value: string): void {
    PropertiesService.getUserProperties().setProperty(
      PROPERTY_PREFIX + key,
      value
    );
  }

  public getSheetByName(sheetName: string): GoogleAppsScript.Spreadsheet.Sheet {
    const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadSheet.getSheetByName(sheetName);
    if (sheet !== null) {
      return sheet;
    }
    sheet = spreadSheet.insertSheet();
    sheet.setName(sheetName);
    return sheet;
  }

  public clearSheet(sheet: GoogleAppsScript.Spreadsheet.Sheet): void {
    sheet.getDataRange().clearContent();
  }

  public getRange(
    sheet: GoogleAppsScript.Spreadsheet.Sheet,
    row?: number,
    column?: number,
    numRows?: number,
    numColms?: number
  ): GoogleAppsScript.Spreadsheet.Range {
    if (row && column && numRows && numColms) {
      return sheet.getRange(row, column, numRows, numColms);
    }
    if (row && column) {
      return sheet.getRange(row, column);
    }
    return sheet.getDataRange();
  }

  public getValues(range: GoogleAppsScript.Spreadsheet.Range): any[][] {
    return range.getValues();
  }

  public setValues(
    range: GoogleAppsScript.Spreadsheet.Range,
    values: any[][]
  ): GoogleAppsScript.Spreadsheet.Range {
    return range.setValues(values);
  }

  public showMessage(title: string, message: string): void {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, title);
  }
}

import {SpreadSheetService} from './SpreadSheetService';

const DefaultQuery = 'is:inbox';
const DefaultChannel = '#random';

export interface Gmail2SlackConfig {
  query: string;
  channel: string;
}

export const Gmail2SlackConfig = (
  query: string,
  channel: string
): Gmail2SlackConfig => ({
  query,
  channel,
});

export function getGmail2SlackConfig(
  spreadsheetService: SpreadSheetService,
  index: number
) {
  let _query = DefaultQuery;
  let _channel = DefaultChannel;
  // 設定sheetを取得する
  const _sheet = spreadsheetService.getSheetByName('設定');
  if (_sheet) {
    // sheetデータを読み込む
    const _range = spreadsheetService.getRange(_sheet);
    const _values = spreadsheetService.getValues(_range);
    if (_values.length > index && _values[index].length === 2) {
      // 設定値
      _query = _values[index][0] || _query;
      _channel = _values[index][1] || _channel;
    }
  }
  return Gmail2SlackConfig(_query, _channel);
}

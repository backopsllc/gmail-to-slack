/* eslint-disable @typescript-eslint/no-unused-vars */
import {getGmail2SlackConfig} from './Gmail2SlackConfig';
import {callWebApi} from './SlackService';
import {SpreadSheetServiceImpl} from './SpreadSheetService';
import {UserProperty} from './UserProperty';

const SpreadsheetService = new SpreadSheetServiceImpl();

function onOpen() {
  const _ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  _ui
    .createMenu('Gmail2Slack')
    .addItem('Step1:未読メール件数', 'menuItem1')
    .addItem('Step2:メール一覧取得', 'menuItem2')
    .addItem('Step3:Slack通知', 'menuItem3')
    .addItem('Step4:古いメール履歴を削除', 'menuItem4')
    .addSeparator()
    .addItem('Step2〜4実行', 'run_cron')
    .addItem('SlackToken設定', 'openDialog')
    .addToUi();
}

function menuItem1() {
  // 受信トレイ内の未読スレッドの数を取得します。
  SpreadsheetService.showMessage(
    'INFO',
    'Inbox Unread Count: ' + GmailApp.getInboxUnreadCount()
  );
}

function menuItem2() {
  const _config = getGmail2SlackConfig(SpreadsheetService, 1);
  SpreadsheetService.showMessage('START', 'Gmail Search: ' + _config.query);
  const _count = getGmailMessages(_config.query);
  SpreadsheetService.showMessage(
    'INFO',
    'Gmail Search: ' + _config.query + ', New Messages: ' + _count
  );
}

function menuItem3() {
  const _config = getGmail2SlackConfig(SpreadsheetService, 1);
  SpreadsheetService.showMessage(
    'START',
    'Post to Slack Channel ' + _config.channel
  );
  const _count = postMessages(_config.query, _config.channel);
  SpreadsheetService.showMessage(
    'INFO',
    'Post to Slack Channel ' + _config.channel + ', ' + _count + ' Messages'
  );
}

function menuItem4() {
  const _config = getGmail2SlackConfig(SpreadsheetService, 1);
  SpreadsheetService.showMessage('START', 'Delete 2 Days Ago Messages');
  const _count = deleteOldMessages(_config.query, 2);
  SpreadsheetService.showMessage('INFO', 'Spreadsheet Delete Rows: ' + _count);
}

function getGmailMessages(query: string) {
  Logger.log('getGmailMessages: ' + query);

  // 1日前よりも新しいメールを取得する
  const _gmailThreads = GmailApp.search('newer_than:1d ' + query);
  Logger.log('Threads count: ' + _gmailThreads.length);

  // sheetを取得する
  const _sheet = SpreadsheetService.getSheetByName(query);
  // sheetデータを読み込む
  const _range = SpreadsheetService.getRange(_sheet);
  const _values = SpreadsheetService.getValues(_range);
  // sheetデータ件数
  const _count = _values.length === 1 ? 0 : _values.length;
  Logger.log('Sheet Data count: ' + _count);

  // 取得済みmessageIdの配列
  const _messageIds = _count === 0 ? [] : _values.slice(1).map(v => v[0]);
  // 保存データ
  const _newValues = _count === 0 ? [] : _values.slice(1);

  // 新着のメールを取得する
  _gmailThreads.forEach(thread => {
    thread
      .getMessages()
      .filter(msg => {
        // sheetデータ内にメッセージIDがない場合、新着メール
        return !_messageIds.includes(msg.getId());
      })
      .forEach(msg => {
        // メッセージID
        const _id = msg.getId();
        Logger.log('New Message Id: ' + _id);
        // 保存データに追加
        _newValues.push([
          _id,
          msg.getFrom(),
          msg.getTo(),
          msg.getSubject(),
          Utilities.formatDate(
            msg.getDate(),
            'Asia/Tokyo',
            'yyyy-MM-dd HH:mm:ss'
          ),
          '',
        ]);
      });
  });

  // 日付でソート
  const _sorted = [
    ['メッセージID', '送信者', '宛先', '件名', '日付', 'slack送信日時'],
  ].concat(
    _newValues.sort(
      (a, b) => new Date(a[4]).getTime() - new Date(b[4]).getTime()
    )
  );

  // 新着メール件数
  const _newCount = _sorted.length - _count;
  Logger.log('New Messages count: ' + _newCount);

  // sheetにデータを書き込む
  if (_newCount > 0) {
    const _newRange = SpreadsheetService.getRange(
      _sheet,
      1,
      1,
      _sorted.length,
      _sorted[0].length
    ).setNumberFormat('@');
    Logger.log('Save Sheet: ' + _sorted.length);
    SpreadsheetService.setValues(_newRange, _sorted);
  }

  return _newCount;
}

function postMessages(query: string, channel: string) {
  Logger.log('postMessages: ' + query + ', ' + channel);

  // sheetを取得する
  const _sheet = SpreadsheetService.getSheetByName(query);
  // sheetデータを読み込む
  const _range = SpreadsheetService.getRange(_sheet);
  const _values = SpreadsheetService.getValues(_range);

  // 送信カウンタ
  let _sendCount = 0;

  // token
  const token = SpreadsheetService.getUserProperty('SLACK_BOT_TOKEN');
  if (token) {
    _values
      // slack未投稿のみ
      .filter(val => val.length === 6 && val[5] === '')
      .forEach(val => {
        Logger.log('Get MessageId: ' + val[0]);
        const _message = GmailApp.getMessageById(val[0]);

        // メッセージ送信する
        const _apiResponse = callWebApi(token, 'chat.postMessage', {
          channel: channel,
          blocks: JSON.stringify([
            {
              type: 'header',
              text: {
                type: 'plain_text',
                text: _message.getSubject(),
                emoji: true,
              },
            },
            {
              type: 'section',
              text: {
                type: 'mrkdwn',
                text:
                  '送信者: ' +
                  _message.getFrom() +
                  '\n宛先: ' +
                  _message.getTo() +
                  '\n日時: ' +
                  Utilities.formatDate(
                    _message.getDate(),
                    'Asia/Tokyo',
                    'yyyy-MM-dd HH:mm:ss'
                  ),
              },
            },
          ]),
        });

        // 送信結果
        if (_apiResponse.getResponseCode() === 200) {
          const _res = JSON.parse(_apiResponse.getContentText());
          if (_res['ok']) {
            // 送信成功
            _sendCount++;

            // slack送信済みにする
            val[5] = Utilities.formatDate(
              new Date(),
              'Asia/Tokyo',
              'yyyy-MM-dd HH:mm:ss'
            );
          }
        }
      });
  }
  Logger.log('Send Messages count: ' + _sendCount);

  // sheetにデータを書き込む
  if (_sendCount > 0) {
    Logger.log('Save Sheet: ' + _values.length);
    SpreadsheetService.setValues(_range, _values);
  }

  return _sendCount;
}

function deleteOldMessages(query: string, daysAgo: number) {
  Logger.log('deleteOldMessages: ' + query + ', ' + daysAgo);

  // x日前の日付を求める
  const _today = new Date();
  const _xDaysAgo = new Date().setDate(_today.getDate() - daysAgo);

  // sheetを取得する
  const _sheet = SpreadsheetService.getSheetByName(query);
  // sheetデータを読み込む
  const _range = SpreadsheetService.getRange(_sheet);
  const _values = SpreadsheetService.getValues(_range);

  // x日以前の行数
  const _removeCount = _values
    .slice(1)
    .filter(row => new Date(row[4]).getTime() < _xDaysAgo).length;
  Logger.log(
    Utilities.formatDate(new Date(_xDaysAgo), 'Asia/Tokyo', 'yyyy-MM-dd') +
      ' 以前の行数: ' +
      _removeCount
  );

  if (_removeCount > 0) {
    // 2行目からx日以前の行まで削除
    _sheet.deleteRows(2, _removeCount);
  }

  return _removeCount;
}

function openDialog() {
  const html = HtmlService.createTemplateFromFile('Index');
  html.mode = 'init';
  SpreadsheetApp.getUi().showModalDialog(html.evaluate(), '設定');
}

const getConfig = function (): UserProperty {
  Logger.log('getConfig');
  const slackBotToken = SpreadsheetService.getUserProperty('SLACK_BOT_TOKEN');
  return UserProperty(slackBotToken);
};

const init = function (property: UserProperty) {
  Logger.log('init: ' + property);
  SpreadsheetService.setUserProperty('SLACK_BOT_TOKEN', property.slackBotToken);
  SpreadsheetService.showMessage('Success', 'Save UserProperties.');
};

const run_cron = function () {
  Logger.log('run_cron');
  // 設定sheetを取得する
  const _sheet = SpreadsheetService.getSheetByName('設定');
  if (_sheet) {
    // sheetデータを読み込む
    const _range = SpreadsheetService.getRange(_sheet);
    const _values = SpreadsheetService.getValues(_range);
    _values
      // 1行目はスキップ
      .slice(1)
      // フィルター条件とチャンネルの設定値が取得できる場合
      .filter(value => value.length === 2)
      // 1行づつ実行
      .forEach(value => {
        const _query = value[0] || '';
        const _channel = value[1] || '#random';
        new Promise<number>(resolve => {
          // 新着メールを取得する
          const _count = getGmailMessages(_query);
          resolve(_count);
        })
          .then(count => {
            if (count > 0) {
              // slackにメッセージを投稿する
              postMessages(_query, _channel);
            }
          })
          .finally(() => {
            // 古いメッセージ履歴を削除する
            deleteOldMessages(_query, 2);
          });
      });
  }
};

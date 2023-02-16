/* eslint-disable @typescript-eslint/no-unused-vars */
import {callWebApi} from './SlackService';
import {SpreadSheetServiceImpl} from './SpreadSheetService';
import {UserProperty} from './UserProperty';

const SpreadsheetService = new SpreadSheetServiceImpl();
const DefaultQuery = 'is:inbox';
const DefaultChannel = '#random';

function onOpen() {
  const _ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  _ui
    .createMenu('Gmail2Slack')
    .addItem('未読メール件数', 'menuItem1')
    .addItem('メール一覧取得', 'menuItem2')
    .addItem('Slack通知', 'menuItem3')
    .addItem('古いメールIDを削除', 'menuItem4')
    .addItem('Token設定', 'openDialog')
    .addToUi();
}

function menuItem1() {
  // 受信トレイ内の未読スレッドの数を取得します。
  SpreadsheetService.showMessage(
    'Info',
    'Messages unread in inbox: ' + GmailApp.getInboxUnreadCount()
  );
}

function menuItem2() {
  SpreadsheetService.showMessage('Start', 'getGmailMessages: ' + DefaultQuery);
  const _count = getGmailMessages(DefaultQuery);
  SpreadsheetService.showMessage('End', 'New Messages count: ' + _count);
}

function menuItem3() {
  SpreadsheetService.showMessage(
    'Start',
    'Post Messages to Slack channel ' + DefaultChannel
  );
  const _count = postMessages(DefaultQuery, DefaultChannel);
  SpreadsheetService.showMessage('End', 'Post Messages count: ' + _count);
}

function menuItem4() {
  Logger.log('menuItem4');
}

function getGmailMessages(query: string) {
  Logger.log('getGmailMessages:' + query);

  // 1日前よりも新しいメールを取得する
  const _gmailThreads = GmailApp.search('newer_than:1d ' + query);
  Logger.log('Threads count: ' + _gmailThreads.length);

  // sheetを取得する
  const _sheet = SpreadsheetService.getSheetByName(query);

  // sheetデータを読み込む
  const _range = SpreadsheetService.getRange(_sheet);
  const _values = SpreadsheetService.getValues(_range);
  // sheetデータ件数
  const _count =
    _values.length === 1 && _values[0].length === 1 ? 0 : _values.length;
  // 保存データ
  const _newValues: any[][] =
    _count === 0
      ? [['メッセージID', '送信者', '宛先', '件名', '日付', 'slack送信日時']]
      : _values;
  Logger.log('Sheet Data count: ' + _count);

  // 新着のメールを追加する
  _gmailThreads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      // メッセージID
      const _id = msg.getId();
      // sheetデータ内にメッセージIDが存在するか
      const _filter = _values.filter(v => {
        return v[0] === _id;
      });
      // メッセージIDがない場合、新着メール
      if (_filter.length === 0) {
        Logger.log('New Message Id: ' + _id);
        // 保存データに追加
        _newValues.push([
          _id,
          msg.getFrom(),
          msg.getTo(),
          msg.getSubject(),
          msg.getDate(),
          '',
        ]);
      }
    });
  });

  // 新着メール件数
  const _newCount = _newValues.length - _count;
  Logger.log('New Messages count: ' + _newCount);

  // sheetにデータを書き込む
  if (_newCount > 0) {
    const _newRange = SpreadsheetService.getRange(
      _sheet,
      1,
      1,
      _newValues.length,
      _newValues[0].length
    ).setNumberFormat('@');
    Logger.log('Save Sheet: ' + _newValues.length);
    SpreadsheetService.setValues(_newRange, _newValues);
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
    _values.forEach(val => {
      if (val.length === 6 && val[5] === '') {
        Logger.log('Get MessageId: ' + val[0]);
        const messageById = GmailApp.getMessageById(val[0]);

        // メッセージ送信する
        const apiResponse = callWebApi(token, 'chat.postMessage', {
          channel: channel,
          blocks: JSON.stringify([
            {
              type: 'header',
              text: {
                type: 'plain_text',
                text: messageById.getSubject(),
                emoji: true,
              },
            },
            {
              type: 'section',
              text: {
                type: 'mrkdwn',
                text:
                  '送信者: ' +
                  messageById.getFrom() +
                  '\n宛先: ' +
                  messageById.getTo() +
                  '\n日時: ' +
                  Utilities.formatDate(
                    messageById.getDate(),
                    'Asia/Tokyo',
                    'yyyy-MM-dd HH:mm:ss'
                  ),
              },
            },
          ]),
        });

        // 送信結果
        if (apiResponse.getResponseCode() === 200) {
          const res = JSON.parse(apiResponse.getContentText());
          if (res['ok']) {
            // 送信成功
            _sendCount++;

            // slack送信済みにする
            val[5] = Utilities.formatDate(
              new Date(),
              'Asia/Tokyo',
              'yyyy-MM-dd HH:mm:ss'
            );
          } else {
            SpreadsheetService.showMessage('Error', `response: ${apiResponse}`);
          }
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
    // 1行づつ実行
    _values.forEach((value, index) => {
      if (index > 0 && value.length === 2) {
        const _query = value[0] || '';
        const _channel = value[1] || '#random';
        new Promise<number>(resolve => {
          const _count = getGmailMessages(_query);
          resolve(_count);
        })
          .then(count => {
            if (count > 0) {
              postMessages(_query, _channel);
            }
          })
          .finally(() => {
            Logger.log('TODO: finally, ' + _query);
          });
      }
    });
  }
};

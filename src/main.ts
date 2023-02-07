/* eslint-disable @typescript-eslint/no-unused-vars */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Custom Menu')
    .addItem('First item', 'menuItem1')
    .addSeparator()
    .addSubMenu(ui.createMenu('Sub-menu').addItem('Second item', 'menuItem2'))
    .addSeparator()
    .addItem('Show alert', 'showAlert')
    .addItem('Show prompt', 'showPrompt')
    .addItem('Show dialog', 'showDialog')
    .addItem('Show sidebar', 'showSidebar')
    .addItem('Open', 'openDialog')
    .addToUi();
}

function menuItem1() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('You clicked the first menu item!');
}

function menuItem2() {
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
    .alert('You clicked the second menu item!');
}

function showAlert() {
  const ui = SpreadsheetApp.getUi(); // Same variations.

  const result = ui.alert(
    'Please confirm',
    'Are you sure you want to continue?',
    ui.ButtonSet.YES_NO
  );

  // Process the user's response.
  if (result === ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }
}

function showPrompt() {
  const ui = SpreadsheetApp.getUi(); // Same variations.

  const result = ui.prompt(
    "Let's get to know each other!",
    'Please enter your name:',
    ui.ButtonSet.OK_CANCEL
  );

  // Process the user's response.
  const button = result.getSelectedButton();
  const text = result.getResponseText();
  if (button === ui.Button.OK) {
    // User clicked "OK".
    ui.alert('Your name is ' + text + '.');
  } else if (button === ui.Button.CANCEL) {
    // User clicked "Cancel".
    ui.alert("I didn't get your name.");
  } else if (button === ui.Button.CLOSE) {
    // User clicked X in the title bar.
    ui.alert('You closed the dialog.');
  }
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Page')
    .setWidth(400)
    .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showModalDialog(html, 'My custom dialog');
}

function showSidebar() {
  const html =
    HtmlService.createHtmlOutputFromFile('Page').setTitle('My custom sidebar');
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
    .showSidebar(html);
}

function openDialog() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate();
  SpreadsheetApp.getUi().showModalDialog(html, 'Dialog title');
}

function getUnreadEmails() {
  return GmailApp.getInboxUnreadCount();
}

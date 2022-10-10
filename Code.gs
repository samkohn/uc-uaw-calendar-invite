// To learn how to use this script, refer to the documentation:
// https://developers.google.com/apps-script/samples/automations/event-session-signup

/*
Copyright 2022 Google LLC

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    https://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
*/

/**
 * Inserts a custom menu when the spreadsheet opens.
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Invites')
      .addItem('Set up form & invitations', 'setUpConference_')
      .addItem('Update with new events', 'addNewEvents_')
      .addItem('Reset', 'resetProperties')
      .addToUi();
}

/**
 * Uses the conference data in the spreadsheet to create
 * Google Calendar events, a Google Form, and a trigger that allows the script
 * to react to form responses.
 */
function setUpConference_() {
  let scriptProperties = PropertiesService.getScriptProperties();
  if (scriptProperties.getProperty('calId')) {
    Browser.msgBox('Your form is already set up. Look in Google Drive for your'
                   + ' sign-up form!');
                   return;
  }
  let ss = SpreadsheetApp.getActive();
  let sheetName = 'Event Setup';
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet) {
    Browser.msgBox('Can\'t find a sheet named "' + sheetName + '". Aborting...');
    return;
  }
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Enter the question you want to use on the form for event selection. E.g. Which trainings will you attend?');
  if(response.getSelectedButton() != ui.Button.OK) {
      Browser.msgBox('Ok, aborting.');
      return;
  }
  let questionName = response.getResponseText();
  scriptProperties.setProperty('questionName', questionName);
  let range = sheet.getDataRange();
  let values = range.getValues();
  setUpCalendar_(values, range);
  setUpForm_(ss, values);
  ScriptApp.newTrigger('onFormSubmit').forSpreadsheet(ss).onFormSubmit()
      .create();
}

/**
 * Creates a Google Calendar with events for each conference session in the
 * spreadsheet, then writes the event IDs to the spreadsheet for future use.
 * @param {Array<string[]>} values Cell values for the spreadsheet range.
 * @param {Range} range A spreadsheet range that contains conference data.
 */
function setUpCalendar_(values, range) {  
  let cal = CalendarApp.createCalendar('Event Calendar');
  // Start at 1 to skip the header row.
  for (let i = 1; i < values.length; i++) {
    let session = values[i];
    let title = session[0];
    let start = joinDateAndTime_(session[1], session[2]);
    let end = joinDateAndTime_(session[1], session[3]);
    let options = {location: session[4], sendInvites: true};
    let event = cal.createEvent(title, start, end, options)
        .setGuestsCanSeeGuests(false);
    session[5] = event.getId();
  }
  range.setValues(values);

  // Stores the ID for the Calendar, which is needed to retrieve events by ID.
  let scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('calId', cal.getId());
}

function addNewEvents_() {
  let ss = SpreadsheetApp.getActive();
  let sheetName = 'Event Setup';
  let sheet = ss.getSheetByName(sheetName);
  if(!sheet) {
    Browser.msgBox('Can\'t find a sheet named "' + sheetName + '". Aborting...');
    return;
  }
  let range = sheet.getDataRange();
  let values = range.getValues();
  let cal = CalendarApp.getCalendarById(PropertiesService.getScriptProperties().getProperty('calId'));
  let formId = PropertiesService.getScriptProperties().getProperty('formId');
  let formArray = []
  if (!formId) {
    let ui = SpreadsheetApp.getUi();
    let response = ui.prompt("Enter the edit URL of the linked form:");
    if (response.getSelectedButton() != ui.Button.OK) {
      Browser.msgBox('Ok, aborting.');
      return;
    }
    formArray.push(response.getResponseText());
    formArray.push(FormApp.openByUrl);
  }
  else {
    formArray.push(formId);
    formArray.push(FormApp.openById);
  }
  Logger.log(formArray);
  let form = formArray[1](formArray[0]);
  let itemId = PropertiesService.getScriptProperties().getProperty('itemId');
  Logger.log(form);
  Logger.log(Number(itemId) + 1);
  let item = form.getItemById(Number(itemId)).asCheckboxItem();
  Logger.log(item);
  let choices = [];
  // Start at 1 to skip the header row.
  for (let i = 1; i < values.length; i++) {
    let session = values[i];
    if (session[5].length > 0) {
      let title = session[0];
      let start = joinDateAndTime_(session[1], session[2]);
      let end = joinDateAndTime_(session[1], session[3]);
      let options = {location: session[4], sendInvites: true};
      let event = cal.createEvent(title, start, end, options)
          .setGuestsCanSeeGuests(false);
      session[5] = event.getId();
      let day = session[1].toLocaleDateString();
      let time = session[2].toLocaleTimeString();
      choices.push(item.createChoice(session[0] + ' | ' + day + ' ' + time));
    }
  }
  range.setValues(values);
  item.setChoices(choices);

}

function test() {
  let properties = PropertiesService.getScriptProperties().getProperties();
  Logger.log(properties);
}

/**
 * Creates a single Date object from separate date and time cells.
 *
 * @param {Date} date A Date object from which to extract the date.
 * @param {Date} time A Date object from which to extract the time.
 * @return {Date} A Date object representing the combined date and time.
 */
function joinDateAndTime_(date, time) {
  date = new Date(date);
  date.setHours(time.getHours());
  date.setMinutes(time.getMinutes());
  return date;
}

/**
 * Creates a Google Form that allows respondents to select which conference
 * sessions they would like to attend, grouped by date and start time in the
 * caller's time zone.
 *
 * @param {Spreadsheet} ss The spreadsheet that contains the conference data.
 * @param {Array<String[]>} values Cell values for the spreadsheet range.
 */
function setUpForm_(ss, values) {
  // Group the sessions by date and time so that they can be passed to the form.
  let schedule = {};
  // Start at 1 to skip the header row.
  /*
  for (let i = 1; i < values.length; i++) {
    let session = values[i];
    let day = session[1].toLocaleDateString();
    let time = session[2].toLocaleTimeString();
    if (!schedule[day]) {
      schedule[day] = {};
    }
    if (!schedule[day][time]) {
      schedule[day][time] = [];
    }
    schedule[day][time].push(session[0]);
  }
    */

  // Creates the form and adds a multiple-choice question for each timeslot.
  let form = FormApp.create('Event Form');
  PropertiesService.getScriptProperties().setProperty('formId', form.getId());
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
  form.addTextItem().setTitle('Name').setRequired(true);
  form.addTextItem().setTitle('Email').setRequired(true);
  let item = form.addCheckboxItem();
  Logger.log(item.getId());
  PropertiesService.getScriptProperties().setProperty('itemId', item.getId());

  let questionName = PropertiesService.getScriptProperties().getProperty('questionName');
  item.setTitle(questionName);
  let choices = [];
  for (let i = 1; i < values.length; i++) {
    let session = values[i];
    let day = session[1].toLocaleDateString();
    let time = session[2].toLocaleTimeString();
    choices.push(session[0] + ' | ' + day + ' ' + time);
  }
  item.setChoiceValues(choices);
  /*
  Object.keys(schedule).forEach(function(day) {
    // let header = form.addSectionHeaderItem().setTitle('Events for ' + day);
    Object.keys(schedule[day]).forEach(function(time) {
      let item = form.addMultipleChoiceItem().setTitle(time + ' ' + day)
          .setChoiceValues(schedule[day][time]);
    });
  });
  */
}

/**
 * Sends out calendar invitations and a
 * personalized Google Docs itinerary after a user responds to the form.
 *
 * @param {Object} e The event parameter for form submission to a spreadsheet;
 *     see https://developers.google.com/apps-script/understanding_events
 */
function onFormSubmit(e) {
  let user = {name: e.namedValues['Name'][0], email: e.namedValues['Email'][0]};

  // Grab the session data again so that we can match it to the user's choices.
  let response = [];
  let values = SpreadsheetApp.getActive().getSheetByName('Event Setup')
      .getDataRange().getValues();
  let questionName = PropertiesService.getScriptProperties().getProperty('questionName');
  let responses = e.namedValues[questionName][0];
  Logger.log(questionName);
  Logger.log(responses);
  Logger.log(responses.length);
  for (let i = 1; i < values.length; i++) {
    let session = values[i];
    let title = session[0];
    let day = session[1].toLocaleDateString();
    let time = session[2].toLocaleTimeString();
    let timeslot = day + ' ' + time;
    let eventName = title + ' | ' + timeslot;
    Logger.log(eventName);
    // For every selection in the response, find the matching timeslot and title
    // in the spreadsheet and add the session data to the response array.
    if (responses.indexOf(eventName) >= 0) {
      response.push(session);
      Logger.log("Event response matches event name");
    }
    /*
    for (let j in responses) {
      let eventResponse = responses[j];
      Logger.log(eventResponse);
      if (eventResponse == eventName) {
        Logger.log("Event response matches event name");
        response.push(session);
      }
    }
    */
    /*
    if (e.namedValues[timeslot] && e.namedValues[timeslot] == title) {
      response.push(session);
    }
    */
  }
  sendInvites_(user, response);
  //sendDoc_(user, response);
}

/**
 * Add the user as a guest for every session he or she selected.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<String[]>} response An array of data for the user's session choices.
 */
function sendInvites_(user, response) {
  let id = ScriptProperties.getProperty('calId');
  let cal = CalendarApp.getCalendarById(id);
  for (let i = 0; i < response.length; i++) {
    cal.getEventSeriesById(response[i][5]).addGuest(user.email);
  }
}

/**
 * Creates and shares a personalized Google Doc that shows the user's itinerary.
 * @param {object} user An object that contains the user's name and email.
 * @param {Array<string[]>} response An array of data for the user's session choices.
 */
function sendDoc_(user, response) {
  let doc = DocumentApp.create('Conference Itinerary for ' + user.name)
      .addEditor(user.email);
  let body = doc.getBody();
  let table = [['Session', 'Date', 'Time', 'Location']];
  for (let i = 0; i < response.length; i++) {
    table.push([response[i][0], response[i][1].toLocaleDateString(),
      response[i][2].toLocaleTimeString(), response[i][4]]);
  }
  body.insertParagraph(0, doc.getName())
      .setHeading(DocumentApp.ParagraphHeading.HEADING1);
  table = body.appendTable(table);
  table.getRow(0).editAsText().setBold(true);
  doc.saveAndClose();

  // Emails a link to the Doc as well as a PDF copy.
  MailApp.sendEmail({
    to: user.email,
    subject: doc.getName(),
    body: 'Thanks for registering! Here\'s your itinerary: ' + doc.getUrl(),
    attachments: doc.getAs(MimeType.PDF),
  });
}

/**
 * Removes the calId script property so that the 'setUpConference_()' can be run again.
 */
function resetProperties(){
  let scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteAllProperties();
}


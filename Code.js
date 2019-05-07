/**
 * @NotOnlyCurrentDoc
 *
 * for - OnlyCurrentDoc The above comment directs Apps Script to limit the scope of file
 * access for this add-on. It specifies that this add-on will only
 * attempt to read or modify the files in which the add-on is used,
 * and not all of the user's files. The authorization request message
 * presented to users will reflect this limited scope.
 */

/**
 * A global constant String holding the title of the add-on. This is
 * used to identify the add-on in the notification emails.
 */
var ADDON_TITLE = 'Workflows';

/**
 * A global constant 'notice' text to include with each email
 * notification.
 */
var NOTICE = 'Workflows was created as an add-on, and is meant to' +
'work as both Form Publisher and Form Approvals. ' +
'The number of notifications this add-on produces are limited by the' +
'owner\'s available email quota; it will not send email notifications if the' +
'owner\'s daily email quota has been exceeded. Collaborators using this add-on on' +
'the same form will be able to adjust the notification settings, but will not be' +
'able to disable the notification triggers set by other collaborators.';

/**
 * Adds a custom menu to the active form to show the add-on sidebar.
 *
 * @param {object} e The event parameter for a simple onOpen trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode.
 */
function onOpen(e) {
  FormApp.getUi()
      .createAddonMenu()
      .addItem('Workflows', 'showSidebar')
      .addItem('About', 'showAbout')
      .addToUi();
}

/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  //gets the current form ID. Check the View-> Logs after running this method
  var form = FormApp.getActiveForm();
  Logger.log('Form id:',form.getId());
  Logger.log('Form name:',form.getTitle());
  onOpen(e);
}

/**
 * Opens a sidebar in the form containing the add-on's user interface for
 * configuring the notification this add-on will produce.
 */
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('Workflows');
  FormApp.getUi().showSidebar(ui);
  var folders = DriveApp.getFolders();
  var folderId;
  var folder;
  
  //Creats folder and template file with form fields
  while (folders.hasNext()) {
    folder = folders.next();
    if(folder.getName() == 'Form Publisher Template')
    {
      folderId = folder.getId()
      break;
    }
  }
  
  if(!folderId)
  {
    folder = DriveApp.createFolder('Form Publisher Template');
    folder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.EDIT);
  }
  
  DriveApp.addFolder(folder);
  var body = DocumentApp.create('Form Publisher Template').getBody();
  var header = body.appendParagraph("A Document");
  header.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  // Append a section header paragraph.
  var section = body.appendParagraph("Section 1");
  section.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  // Append a regular paragraph.
  body.appendParagraph("This is a typical paragraph.");
  
  var items = getFormFields();
  
  for (var i = 0; i < items.length; i++) {
      body.appendListItem(items[i].getTitle());
  }
  createSubmitTrigger();
}

/**
 * Opens a purely-informational dialog in the form explaining details about
 * this add-on.
 */
function showAbout() {
  var ui = HtmlService.createHtmlOutputFromFile('About')
      .setWidth(420)
      .setHeight(270);
  FormApp.getUi().showModalDialog(ui, 'About Logicwind Workflows');
}


/**
 * Get the form fields to populate in template
 *
 * @return {Object}
 */
function getFormFields() {

  var form = FormApp.getActiveForm();
  var items = form.getItems();
  return items;
}

/**
 * Create the Submit trigger.
 */
function createSubmitTrigger() {
  var form = FormApp.getActiveForm();
  var triggers = ScriptApp.getUserTriggers(form);
  var settings = PropertiesService.getDocumentProperties();
  Logger.log(settings);
  
  var existingTrigger = null;
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getEventType() == ScriptApp.EventType.ON_FORM_SUBMIT) {
      existingTrigger = triggers[i];
      break;
    }
  }
  if (!existingTrigger) {
    var trigger = ScriptApp.newTrigger('respondToFormSubmit')
        .forForm(form)
        .onFormSubmit()
        .create();
  } else if (existingTrigger) {
    ScriptApp.deleteTrigger(existingTrigger);
  }
}

/**
 * Responds to a form submission event if an onFormSubmit trigger has been
 * enabled.
 *
 * @param {Object} e The event parameter created by a form
 *      submission; see
 *      https://developers.google.com/apps-script/understanding_events
 */
function respondToFormSubmit(e) {
  
  //cal external URL to pass field names
  
  var form = FormApp.getActiveForm();
  var items = form.getItems();
  var formData = [];
  
  for (var i = 0; i < items.length; i++) {
    formData.push({
      title: items[i].getTitle(),
      id: items[i].getId()
    });
  }
  
  var options = {
  'method' : 'post',
  'payload' : JSON.stringify(formData)
  };
  var response = UrlFetchApp.fetch('http://ptsv2.com/t/Heli/post', options);
  Logger.log("response:",response);
  
  //limit form submission
  limitFormSubmissionByNumber();
  
    if (MailApp.getRemainingDailyQuota() > 0) {
      sendRespondentNotification(e.response);
    }
}

/**
 * Sends out creator notification email(s) if the current number
 * of form submission is increased by the given number in the form
 */
function limitFormSubmissionByNumber() {
  var form = FormApp.getActiveForm();
  var limit = 100;
  var address = 'heli.patel91@gmail.com';
  
  var formResponses = form.getResponses();
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      if(itemResponse.getItem().getTitle() == 'Form Submission Limit')
        limit = itemResponse.getResponse();
       if(itemResponse.getItem().getTitle() == 'Email Address')
        address = itemResponse.getResponse();
    }
  }
  
  if (form.getResponses().length == limit) {
    if (MailApp.getRemainingDailyQuota() > 0) {
      var template =
          HtmlService.createTemplateFromFile('CreatorNotification');
      template.sheet =
          DriveApp.getFileById(form.getDestinationId()).getUrl();
      template.summary = form.getSummaryUrl();
      template.responses = form.getResponses().length;
      template.title = form.getTitle();
      template.responseStep = limit;
      template.formUrl = form.getEditUrl();
      template.notice = NOTICE;
      var message = template.evaluate();
      MailApp.sendEmail(address,
          form.getTitle() + ': Form submissions increased detected',
          message.getContent(), {
            name: ADDON_TITLE,
            htmlBody: message.getContent()
          });
    }
    form.setAcceptingResponses(false);
  }
}


function limitFormSubmissionByDateTime()
{
}
/**
 * Sends out respondent notification emails.
 *
 * @param {FormResponse} response FormResponse object of the event
 *      that triggered this notification
 */
function sendRespondentNotification(response) {
  var form = FormApp.getActiveForm();
  var address = 'heli.patel91@gmail.com';
  
  var formResponses = form.getResponses();
  for (var i = 0; i < formResponses.length; i++) {
    var formResponse = formResponses[i];
    var itemResponses = formResponse.getItemResponses();
    for (var j = 0; j < itemResponses.length; j++) {
      var itemResponse = itemResponses[j];
      if(itemResponse.getItem().getTitle() == 'Email Address')
        address = itemResponse.getResponse();
    }
  }
  var settings = PropertiesService.getDocumentProperties();
  var respondentEmail = address;
  if (respondentEmail) {
    var template =
        HtmlService.createTemplateFromFile('RespondentNotification');
    template.paragraphs = settings.getProperty('responseText').split('\n');
    template.notice = NOTICE;
    var message = template.evaluate();
    MailApp.sendEmail(respondentEmail,
        settings.getProperty('responseSubject'),
        message.getContent(), {
          name: form.getTitle(),
            htmlBody: message.getContent()
        });
  }
}

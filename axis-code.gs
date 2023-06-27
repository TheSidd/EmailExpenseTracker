//REPLACE WITH YOUR GOOGLE SHEET URL
let your_spreadsheet_url="https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxx/edit?usp=sharing";

let axis_email_search="newer_than:1d AND in:inbox AND from:axisbank.com AND subject:Transaction alert AND -label:axis_processed";
let axis_regex= new RegExp(/Card no.\s(XX\d+)\sfor\s([A-Z]{3})\s(\d+(?:\.\d+)?)\s*at\s([\s\S]+?)\son\s*(\d{2}-\d+-\d+\s\d+:\d+:\d+)/);
let axis_sheet_name="Axis";
let axis_account_name="Axis";
let axis_gmail_label="axis_processed";

let hdfc_email_search="newer_than:1d AND from:(alerts@hdfcbank.net) subject:(Alert : Update on your HDFC Bank Credit Card) AND -label:hdfc_processed";
let hdfc_regex= new RegExp(/HDFC Bank Credit Card ending.([0-9]{4}).for.([A-Za-z]{2}).(\d+(?:\.\d+)?).at.([\s\S]+?)\son\s*(\d{2}-\d+-\d+\s\d+:\d+:\d+)/);
let hdfc_sheet_name="HDFC";
let hdfc_account_name="HDFC";
let hdfc_gmail_label="hdfc_processed";


function getRelevantMessages(email_search)
{
  var threads = GmailApp.search(email_search,0,100);
  var arrToConvert=[];
  for(var i = threads.length - 1; i >=0; i--) {
    arrToConvert.push(threads[i].getMessages());   
  }
  var messages = [];
  for(var i = 0; i < arrToConvert.length; i++) {
    messages = messages.concat(arrToConvert[i]);
  }
  return messages;
}

function parseMessageData(messages,regex_pattern,account_name)
{
  var records=[];
  if(!messages)
  {
    //messages is undefined or null or just empty
    return records;
  }
  for(var m=0;m<messages.length;m++)
  {
    var text = messages[m].getPlainBody();
    var matches = text.replace(/(\r\n|\n|\r)/gm, "");
    matches = matches.match(regex_pattern);
    
    if(!matches || matches.length < 6)
    {
      //No matches; couldn't parse continue with the next message
      continue;
    }
    var rec = {};
    rec.account = account_name;
    rec.currency = matches[2];
    rec.card = matches[1];
    rec.date= matches[5];
    rec.merchant = matches[4].replace(/\s+/g, ' ').trim();
    rec.amount = matches[3];

    records.push(rec);
  }
    return records;

}

function getMessagesDisplay(email_search)
{
  var templ = HtmlService.createTemplateFromFile('messages');
  templ.messages = getRelevantMessages(email_search);
  return templ.evaluate();  
}

function getParsedDataDisplay()
{
  var templ = HtmlService.createTemplateFromFile('parsed');
  var axis_records=parseMessageData(getRelevantMessages(axis_email_search),axis_regex,axis_account_name);
  var hdfc_records=parseMessageData(getRelevantMessages(hdfc_email_search),hdfc_regex,hdfc_account_name);
  templ.records=axis_records.concat(hdfc_records);
  return templ.evaluate();
}

function saveDataToSheet(records,sheet_name)
{
  var spreadsheet = SpreadsheetApp.openByUrl(your_spreadsheet_url);
  var sheet = spreadsheet.getSheetByName(sheet_name);
  for(var r=0;r<records.length;r++)
  {
    sheet.appendRow([records[r].date,records[r].card, records[r].merchant, records[r].amount, records[r].currency] );
  }
  
}

function processTransactionEmails()
{
  var axis_messages = getRelevantMessages(axis_email_search);
  var axis_records = parseMessageData(axis_messages,axis_regex,axis_account_name);
  saveDataToSheet(axis_records,axis_sheet_name);
  labelMessagesAsDone(axis_messages,axis_gmail_label);

  var hdfc_messages = getRelevantMessages(hdfc_email_search);
  var hdfc_records = parseMessageData(hdfc_messages,hdfc_regex,hdfc_account_name);
  saveDataToSheet(hdfc_records,hdfc_sheet_name);
  labelMessagesAsDone(hdfc_messages,hdfc_gmail_label);

  return true;
}

function labelMessagesAsDone(messages,email_label)
{
  var label_obj = GmailApp.getUserLabelByName(email_label);
  if(!label_obj)
  {
    label_obj = GmailApp.createLabel(label);
  }
  
  for(var m =0; m < messages.length; m++ )
  {
     label_obj.addToThread(messages[m].getThread() );  
  }
  
}

function doGet()
{
  return getParsedDataDisplay();

  //return getMessagesDisplay();
}

// Room Reservation System Video
// Nibir Das, 2023
// All rights reserved

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheets()[0];
var lastRow = sheet.getLastRow();
var lastColumn = sheet.getLastColumn();


// Calendars to output appointments to
var cal101 = CalendarApp.getCalendarById('6a422f0eb31f47fea5f3fad70058add58efc78e0b56e1a56fe17890d23fa632a@group.calendar.google.com');
var cal102 = CalendarApp.getCalendarById('7db6e74b7d892acfb9096389c1589544c9a8fb35cb2aa15eebfa640898422cf1@group.calendar.google.com');
var cal201 = CalendarApp.getCalendarById('5206eb09218c02c8057e58ae9db6293972db626a040bcdc4952beed8d3a79b58@group.calendar.google.com');
var cal202 = CalendarApp.getCalendarById('d3512fd781bb827b456d5acd7fbe72801609a64186992047a83af122d0645cc1@group.calendar.google.com');
var cal301 = CalendarApp.getCalendarById('32ffa8f0895324c1d1cdde825162e83e634dbd90dc9c614b66059c82dc899cef@group.calendar.google.com');


function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}



function formatTimeString(date) {
  var hours = date.getHours();
  var minutes = date.getMinutes();
  var Seconds = date.getSeconds();
  return `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}:${Seconds.toString().padStart(2, '0')}`;
}


// Create an object from user submission
// Create an object from user submission
function Submission(){
  var row = lastRow;
  this.timestamp = sheet.getRange(row, 1).getValue();
  this.name = sheet.getRange(row, 2).getValue();
  this.email = sheet.getRange(row, 3).getValue();
  this.id = sheet.getRange(row, 4).getValue();
  this.date = sheet.getRange(row, 5).getValue();
  this.time = sheet.getRange(row, 6).getValue();
  this.duration = sheet.getRange(row, 7).getValue();
  this.room = sheet.getRange(row, 8).getValue();
  // Info not from spreadsheet
  this.roomInt = this.room.replace(/\D+/g, '');
  this.status;
  this.dateString = (this.date.getMonth() + 1) + '/' + this.date.getDate() + '/' + this.date.getFullYear();
  this.timeString = formatTimeString(this.time);
  this.date.setHours(this.time.getHours());
  this.date.setMinutes(this.time.getMinutes());
  this.calendar = eval('cal' + String(this.roomInt));
  return this;
}


// Use duration to create endTime variable
function getEndTime(request){
  request.endTime = new Date(request.date);
  switch (request.duration){
    case "30 minutes":
      request.endTime.setMinutes(request.date.getMinutes() + 30);
      //request.endTimeString = formatTime(request.endTime);
      break;
    case "45 minutes":
      request.endTime.setMinutes(request.date.getMinutes() + 45);
      //request.endTimeString = formatTime(request.endTime);
      break;
    case "1 hour":
      request.endTime.setMinutes(request.date.getMinutes() + 60);
      //request.endTimeString = formatTime(request.endTime);
      break;
    case "2 hours":
      request.endTime.setMinutes(request.date.getMinutes() + 120);
      //request.endTimeString = formatTime(request.endTime);
      break;
  }
  request.endTimeString = formatTimeString(request.endTime);
}


// Check for appointment conflicts
function getConflicts(request){
  var conflicts = request.calendar.getEvents(request.date, request.endTime);
  if (conflicts.length < 1) {
    request.status = "Approve";
  } else {
    request.status = "Conflict";
  }
}

function draftEmail(request){
  request.buttonLink = "https://forms.gle/UF4s1hybx6YbfV5F9";
  request.buttonText = "New Request";
  switch (request.status) {
    case "Approve":
      request.subject = "Confirmation: " + request.room + " Reservation for " + request.dateString;
      request.header = "Confirmation";
      request.message = "Room Booked Successfully! Contact library for more Info.";
      break;
    case "Conflict":
      request.subject = "Conflict with " + request.room + "Reservation for " + request.dateString;
      request.header = "Conflict";
      request.message = "There is a scheduling conflict. Please pick another room or time. Contact library for more info!"
      request.buttonText = "Reschedule";
      break;
  }
}

function updateCalendar(request){
  var event = request.calendar.createEvent(
    request.name,
    request.date,
    request.endTime
    )
}

function sendEmail(request){
  MailApp.sendEmail({
    to: request.email,
    subject: request.header,
    htmlBody: makeEmail(request)
  })
  sheet.getRange(lastRow, lastColumn).setValue("Sent: " + request.status);
}

// --------------- main --------------------

function main(){
  var request = new Submission();
  getEndTime(request);
  getConflicts(request);
  draftEmail(request);
  if (request.status == "Approve") updateCalendar(request);
  sendEmail(request);
}



















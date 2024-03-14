function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Feedback')
      .addItem('Calculate Feedback', 'calculateAverageFeedbacks')
      .addItem('Feedback Report', 'generateFeedbackDoc')
      .addToUi();
}

function calculateAverageFeedbacks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName("Sheet1"); // Change "Sheet1" to the name of your source sheet
  var targetSheet = ss.getSheetByName("Sheet2"); // Change "Sheet2" to the name of your target sheet
  
  
  // Check if the "Feedback" sheet exists
  var targetSheet = ss.getSheetByName("Feedback");
  
  // If the "Feedback" sheet exists, delete it
  if (targetSheet) {
    ss.deleteSheet(targetSheet);
  }
  
  // Create a new "Feedback" sheet
  targetSheet = ss.insertSheet("Feedback");


  var numRows = sourceSheet.getLastRow();
  var numColumns = sourceSheet.getLastColumn();
  
  // Assuming there are 10 feedbacks for each subject
  var numFeedbacksPerSubject = 10;
  
  // Loop through each subject
  for (var i = 3; i <= numColumns; i += numFeedbacksPerSubject + 1) {
    var subjectName = sourceSheet.getRange(1, i).getValue(); // Get subject name from the first row
    var feedbacksRange = sourceSheet.getRange(2, i + 1, numRows - 1, numFeedbacksPerSubject); // Get range of feedbacks for this subject
   
    var averageFeedbacks = [];
    var sub = sourceSheet.getRange(2, i).getValue();
    Logger.log(sub)
    // Calculate average for each feedback
    for (var j = 1; j <= numFeedbacksPerSubject; j++) {
      
      var feedbackColumn = i + j;
    
      var feedbackValues = sourceSheet.getRange(2, feedbackColumn, numRows - 1, 1).getValues();
      
      
      var sum = 0;
      for (var k = 0; k < feedbackValues.length; k++) {
        sum += feedbackValues[k][0];
      }
     
      var average = sum / (numRows - 1);
      
      averageFeedbacks.push([average]);
    }
    
    // Write the averages to the target sheet
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1).setValue(sub);
 
var constantValues = [
  ["Lectures are well prepared?"],
  ["Effectiveness of delivery of lectures and communication"],
  ["Knowledgeable and makes the course interesting"],
  ["Encourages discussion, creative / critical thinking and clears doubts"],
  ["Regular & Punctual to the class"],
  ["Cycle test papers are evaluated and distributed in time"],
  ["Transparency and fairness of Evaluation system / Internal Examinations"],
  ["Availability of the Faculty after class hours for guidance"],
  ["The prescribed syllabi is entirely completed"],
  ["The source material (books etc) for each portion covered in the syllabi is clearly informed."]
];


var lastrow=targetSheet.getLastRow();

   targetSheet.getRange(lastrow+1, 1, numFeedbacksPerSubject, 1).setValues(constantValues);
   targetSheet.getRange(lastrow+1, 2, numFeedbacksPerSubject, 1).setValues(averageFeedbacks);

  
  var sum1 = 0;
for (var ii = 0; ii < averageFeedbacks.length; ii++) {
  sum1 += averageFeedbacks[ii][0]; // Access the value inside each inner array
}

var average = sum1 / averageFeedbacks.length;
var lastRow1 = targetSheet.getLastRow() + 1;
targetSheet.getRange(lastRow1, 1).setValue("Overall Score");
targetSheet.getRange(lastRow1, 2).setValue(average);

  }
}

function generateFeedbackDoc() {
/*  subjectLine = Browser.inputBox("Enter the semester Details",Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine === "cancel" || subjectLine == ""){ 
    // If no subject line, finishes up
    return;
    }

     subjectLine1 = Browser.inputBox("Enter the semester Details",Browser.Buttons.OK_CANCEL);
                                      
    if (subjectLine1 === "cancel" || subjectLine1 == ""){ 
    // If no subject line, finishes up
    return;
    }*/
     
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var targetSheet = ss.getSheetByName("Feedback");
  var numRows = targetSheet.getLastRow();
  
  // Create a new document without any content
var existingDocs = DriveApp.getFilesByName('Feedback Report');
  
 
  // Delete the existing document if found
  while (existingDocs.hasNext()) {
    var file = existingDocs.next();
    file.setTrashed(true);
  }

  // Create a new document
  var doc = DocumentApp.create('Feedback Report');
  
  // Get the header of the document
  var header = doc.addHeader();
  
  // Create a table within the header to align logo and text
  var table = header.appendTable();
  var row = table.appendTableRow();
 
  // Loop through each row in the "Feedback" sheet
  for (var i = 1; i <= numRows; i += 12) {
    var subject = targetSheet.getRange(i, 1).getValue();
    var feedbacks = targetSheet.getRange(i + 1, 1, 10, 2).getValues();
    
    // Add subject name as heading
    var body = doc.getBody();
    
  //  var subjectHeading = body.appendParagraph(subject);
  //  subjectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var parts = subject.split('/');

var courseCodeName = parts[0].trim()+" / "+parts[1].trim(); // First part contains the course code/name
var facultyName = parts[2].trim(); 
/*
body.appendParagraph(subjectLine).setAlignment(DocumentApp.HorizontalAlignment.CENTER);;
body.appendParagraph('');
body.appendParagraph(subjectLine1).setAlignment(DocumentApp.HorizontalAlignment.CENTER);;
*/
    body.appendParagraph('Course code / Name: ' + courseCodeName);
    body.appendParagraph('');
body.appendParagraph('Faculty Name: ' + facultyName);
body.appendParagraph('');
    // Create a table for feedback
    var table = body.appendTable();
    var tableRow = table.appendTableRow();
var headerCell1=tableRow.appendTableCell("S.No").setWidth(50); // Set width for SNO
var headerCell2=tableRow.appendTableCell("Details").setWidth(350); // Set width for Details
var headerCell3=tableRow.appendTableCell("Score /10").setWidth(80);
headerCell1.setBold(true);
headerCell2.setBold(true);
headerCell3.setBold(true);
// Initialize the SNO counter
var sno = 1;

// Loop through each feedback and add to the table
for (var j = 0; j < feedbacks.length; j++) {
  var row = table.appendTableRow();
  
  // Add SNO
  var a=String(j+1).split('.')[0];
  
  var cell1=row.appendTableCell(a);
  
  // Add Details and Score
  var cell2=row.appendTableCell(feedbacks[j][0]);
  var cell3=row.appendTableCell(feedbacks[j][1]);
  cell1.setBold(false);
  cell2.setBold(false);
  cell3.setBold(false);
}
    
    // Add Total
    var total = targetSheet.getRange(i + 11, 2).getValue();
    var overallScoreParagraph=body.appendParagraph('Overall Score: ' + total);
    overallScoreParagraph.setAlignment(DocumentApp.HorizontalAlignment.RIGHT);
    body.appendPageBreak();
  }

   var table = body.appendTable();
   var tableRow = table.appendTableRow();
   var headerCell1=tableRow.appendTableCell("S.No").setWidth(50); // Set width for SNO
var headerCell2=tableRow.appendTableCell("Subject Name ").setWidth(200); // Set width for Details

var headerCell3=tableRow.appendTableCell("Faculty Name").setWidth(120);  
var headerCell4=tableRow.appendTableCell("Score /10").setWidth(80); 
headerCell1.setBold(true);
headerCell2.setBold(true);
headerCell3.setBold(true);
headerCell4.setBold(true);
var count=1;
  for (var i = 1; i <= numRows; i += 12) {
    var subject = targetSheet.getRange(i, 1).getValue();
    var feedbacks = targetSheet.getRange(i + 1, 1, 10, 2).getValues();
    
    // Add subject name as heading
    var body = doc.getBody();
    
  //  var subjectHeading = body.appendParagraph(subject);
  //  subjectHeading.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  var parts = subject.split('/');

var courseCodeName = parts[0].trim()+" / "+parts[1].trim(); // First part contains the course code/name
var facultyName = parts[2].trim(); 
var total = targetSheet.getRange(i + 11, 2).getValue();
 var row = table.appendTableRow();
 var a=String(count).split('.')[0];
  
  var cell1=row.appendTableCell(a);
  var cell2=row.appendTableCell(courseCodeName);
  var cell3=row.appendTableCell(facultyName);
  var cell4=row.appendTableCell(total);
  cell1.setBold(false);
  cell2.setBold(false);
  cell3.setBold(false);
  cell4.setBold(false);
  count++;

  }
  
  // Save and close the document
  doc.saveAndClose();
}


# GSFMS
Google Sheets Feedback Management Script

# Google Sheets Feedback Management Script Documentation

## Overview
This script is designed to automate feedback management tasks in Google Sheets. It provides functions to calculate average feedback scores and generate a comprehensive feedback report document.

## Functions

### onOpen():
- This function creates a custom menu named 'Feedback' in the Google Sheets UI.
- Menu items:
  - 'Calculate Feedback': Triggers the calculation of average feedback scores.
  - 'Feedback Report': Triggers the generation of a feedback report document.

### calculateAverageFeedbacks():
- Calculates average feedback scores for various subjects based on feedback data stored in a source sheet.
- It loops through each subject, calculates the average score for each feedback criterion, and writes the results to a target sheet.
- Finally, it calculates the overall score for each subject and updates the target sheet accordingly.

### generateFeedbackDoc():
- Generates a feedback report document based on the calculated average scores stored in the target sheet.
- It creates a new Google Document named 'Feedback Report' and constructs the document with feedback data.
- Feedback for each subject is organized in tables, including details such as course code/name, faculty name, feedback criteria, and scores.
- Additionally, it provides an overview table summarizing subjects, faculty names, and overall scores.

## Usage
1. Open your Google Sheets document containing feedback data.
2. Navigate to 'Extensions' > 'Custom menu' > 'Feedback'.
3. Choose either 'Calculate Feedback' to compute average feedback scores or 'Feedback Report' to generate a detailed feedback report document.

## Notes
- Ensure that your Google Sheets document contains feedback data structured appropriately, with subjects listed horizontally and feedback criteria listed vertically.
- Customize sheet names and other parameters as per your spreadsheet structure.
- Review the generated document to ensure accuracy and completeness of the feedback report.

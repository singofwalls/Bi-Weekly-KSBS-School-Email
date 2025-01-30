function send_welcome_email(e) {
  const column_names = {
    watched: "Welcome Letter Sent",
    delegate_fname: "Preferred First Name",
    delegate_lname: "Delegate Last Name",
    delegate_email: "Delegate Primary Email Address",
    parent: "Name",
    parent_email: "Email"
  }
  const sheet_name = "2025 Current Delegate List";
  const letter_doc_id = "1PJIU6Gd0InsZByySZvaB6SYyDo55MJKsJUIXPJZPkcs";
  const letter_dest_id = '10QOTNhneAKKMPWq6ovBJ5wEOPzUc3qj0';
  const email_subject = "You have been accepted into the Boys State of Kansas";
  let delegate_email_body = "[delegate_fname],<br><br>On behalf of the American Legion Boys State of Kansas, I'd like to inform you that your application has been accepted! Please see the attached acceptance letter! We are very excited to welcome you to the ranks of thousands of other young men from Kansas who have been a delegate at Kansas Boys State.<br><br>Thank you,<br><br>Kyle Wheatley<br>Executive Director<br>American Legion Boys State of Kansas";
  let parent_email_body = "To the parent/guardian of [delegate_fname] [delegate_lname],<br><br>On behalf of the American Legion Boys State of Kansas, I'd like to inform you that his application has been accepted! Please see the attached acceptance letter! We are very excited to welcome him to the ranks of thousands of other young men from Kansas who have been a delegate at Kansas Boys State.<br><br>Thank you,<br><br>Kyle Wheatley<br>Executive Director<br>American Legion Boys State of Kansas";

  if (e === undefined) {
    // e is undefined e.g. when ran from the editor instead of as a triggered event.
    // Popluate e with test data.
    e = { range: { getColumn: () => 4, getRow: () => 60 }, oldValue: 'N/A', value: 'Yes', source: { getSheetName: () => "2025 Current Delegate List"} };
  }

  // console.log(e.range.getColumn(), e.range.getRow(), e.oldValue, e.value, e.source.getSheetName());
  if (e.source.getSheetName() !== sheet_name) {
    // Ignore edits to other sheets
    console.log("Ignoring edit to wrong sheet: " + e.source.getName());
    return;
  }

  let sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheet_name);
  let rows = sheet.getDataRange().getValues();

  let headers = rows[0];

  let edited_column_index = e.range.getColumn() - 1;
  let edited_row_index = e.range.getRow() - 1;

  // Get column indices
  let column_indices = {};
  Object.keys(column_names).forEach(k => {
    column_indices[k] = headers.indexOf(column_names[k]);
  })

  // get column values
  let column_values = {};
  Object.keys(column_names).forEach(k => {
    column_values[k] = rows[edited_row_index][column_indices[k]];
  })

  if (edited_column_index !== column_indices['watched']) {
    // Ignore edits to other columns
    console.log("Ignoring edit to wrong column: " + headers[e.range.getColumn()]);
    return;
  }

  if (e.oldValue.toLowerCase() === "yes") {
    // Ignore rows that the email has already been sent for
    console.log("Ignoring edit from value that was already 'yes'");
    return;
  }

  if (e.value.toLowerCase() !== "yes") {
    // Ignore rows that are not now marked for send
    console.log("Ignoring edit to non-yes value: " + e.value);
    return;
  }

  // Column was switched to yes for this stater. Prepare and send letter

  let delegate_username = column_values['delegate_fname'] + " " + column_values['delegate_lname'];


  console.log("Row switched to yes for: " + delegate_username);

  // Create the acceptance letter copy
  let letter_doc_template = DriveApp.getFileById(letter_doc_id);
  let letter_dest = DriveApp.getFolderById(letter_dest_id);
  let letter_doc = letter_doc_template.makeCopy("2025 Acceptance Letter - " + delegate_username, letter_dest);
  console.log("Created acceptance letter for : " + delegate_username);

  // Edit for specific stater
  let letter_doc_open = DocumentApp.openById(letter_doc.getId());
  let letter_doc_body = letter_doc_open.getBody();
  letter_doc_body.replaceText("\\[date\\]", new Date().toLocaleDateString());  // have to double escape the brackets cause the pattern is passed as a string instead of a regex object
  letter_doc_body.replaceText("\\[preferred name\\]", column_values['delegate_fname']);
  letter_doc_open.saveAndClose();

  // save as pdf
  let letter_pdf = letter_doc.getAs('application/pdf');
  letter_pdf.setName(letter_doc_open.getName() + ".pdf");
  let letter_pdf_saved = DriveApp.createFile(letter_pdf);
  letter_pdf_saved.moveTo(letter_dest);

  // delete doc now that we have pdf
  letter_doc.setTrashed(true);

  console.log("Created pdf: " + letter_pdf_saved.getName());

  // Send the email
  console.log("Remaining emails we can send today: " + MailApp.getRemainingDailyQuota());

  // Send delegate email
  // replace variables
  Object.keys(column_names).forEach(k => {
    delegate_email_body = delegate_email_body.replace("[" + k + "]", column_values[k]);
  });
  console.log("Sending delegate email to " + column_values['delegate_email'], email_subject, delegate_email_body);
  MailApp.sendEmail({
    to: column_values['delegate_email'],
    subject: email_subject,
    htmlBody: delegate_email_body,  // non-html body inserts line break every 80 characters for no reason
    attachments: [letter_pdf_saved.getAs(MimeType.PDF)]
  });

  // Send parent email
  // replace variables
  Object.keys(column_names).forEach(k => {
    parent_email_body = parent_email_body.replace("[" + k + "]", column_values[k]);
  });
  console.log("Sending parent email to " + column_values['parent_email'], email_subject, parent_email_body);
  MailApp.sendEmail({
    to: column_values['parent_email'],
    subject: email_subject,
    htmlBody: parent_email_body,  // non-html body inserts line break every 80 characters for no reason
    attachments: [letter_pdf_saved.getAs(MimeType.PDF)]
  });
}
function sendEmailsFromSheets() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = sheet.getSheetByName('2025 Master Contact List');
  var delegateSheet = sheet.getSheetByName('2025 Current Delegate List');

  var masterData = masterSheet.getDataRange().getValues();
  var delegateData = delegateSheet.getDataRange().getValues();

  for (var i = 1; i < masterData.length; i++)  {
    var row = masterData[i];
    var schoolName = row[11]; // Column 'High School'
    var contactEmails = row[16].split(";"); // Column 'Contact eMail'
    var contactNames = row[14].split(';'); // Column 'Contact'

    let contactFirstNames = [];
    for (let contact of contactNames) {
      contactFirstNames.push(contact.trim(" ").split(" ")[0]);
    }

    if (contactEmails.length == 0) {
      Logger.log('Skipping school ' + schoolName + ' due to missing contact email.');
      continue;
    }

    var delegates2025 = parseInt(row[6], 10); // Column 'Delegates for 2025'
    var delegates2024 = parseInt(row[5], 10); // Column 'Delegates from 2024' (note the space in column name)
    var delegates2023 = parseInt(row[4], 10); // Column 'Delegates from 2023'

    var schoolDelegates = delegateData.filter(d => d[0] === schoolName); // Match school name in '2025 Delegates' sheet
    var delegateNames = schoolDelegates.map(d => d[1] + ', Grade ' + d[2]).join('\n');


    for (let contact_num = 0; contact_num < contactEmails.length; contact_num++) {
      let contactEmail = contactEmails[contact_num].trim();
      if (contactEmail.length == 0) {
        // if there is a semi colon with nothing after it in the email field
        continue;
      }
      if (!contactEmail.includes("@")) {
        continue;
      }

      let contactFirstName;
      if (contact_num >= contactFirstNames.length) {
        contactFirstName = contactFirstNames[0];
      } else {
        // if all emails are using one name (e.g. "ConselingSecretary"), use that
        contactFirstName = contactFirstNames[contact_num];
      }
      var subject, body;
      if (delegates2025 > 0) {
        subject = 'Follow-up on Boys State Participation for ' + schoolName;
        body = `Greetings ${contactFirstName}!

This is Kyle Wheatley, Executive Director of The American Legion Boys State of Kansas, checking in to see what we can do to have more students from your school attend the 2025 Session. Your school should have received a packet of information regarding our program in late 2024, please let me know if you did not receive it.

We currently have ${delegates2025} signed up to attend. The following students have applied and been accepted to attend:

${delegateNames}

[Program Details]

Thank you for all you do for your students!

Kyle Wheatley
Executive Director
The American Legion Boys State of Kansas`;
      } else if (delegates2024 > 0 || delegates2023 > 0) {
        subject = 'Encouraging Boys State Participation for ' + schoolName;
        body = `Greetings ${contactFirstName}!

This is Kyle Wheatley, Executive Director of The American Legion Boys State of Kansas, checking in to see what we can do to have students from your school attend the 2025 Session of Boys State. Your school should have received a packet of information regarding our program in late 2024, please let me know if you did not receive it.

We currently have no delegates signed up to attend but would love to have students attend to represent your school and community.

[Program Details]

Thank you for all you do for your students!

Kyle Wheatley
Executive Director
The American Legion Boys State of Kansas`;
      } else {
        subject = 'Introducing Boys State to ' + schoolName;
        body = `Greetings ${contactFirstName}!

This is Kyle Wheatley, Executive Director of The American Legion Boys State of Kansas, checking in to see what we can do to have students from your school attend the 2025 Session of Boys State. Your school should have received a packet of information regarding our program in late 2024, please let me know if you did not receive it.

We have not had any students attend in the past two years and currently have none signed up to attend this year.

[Program Details]

Thank you for all you do for your students!

Kyle Wheatley
Executive Director
The American Legion Boys State of Kansas`;
      }


      // MailApp.sendEmail(contactEmail, subject, body);
      Logger.log('Email sent to: ' + contactFirstName + ' <' + contactEmail + '> for school: ' + schoolName);
    }

}
}
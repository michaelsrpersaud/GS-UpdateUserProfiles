// References:
// https://stackoverflow.com/questions/42632622/use-google-apps-script-to-loop-through-the-whole-column
// https://stackoverflow.com/questions/23364870/update-phone-number-on-google-apps-user


function update_user_profile() {
  try {

    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    var lastSourceRow = sourceSheet.getLastRow();
    var lastSourceCol = sourceSheet.getLastColumn();

    var sourceRange = sourceSheet.getRange(2, 1, lastSourceRow,lastSourceCol);
    //     var sourceRange = sourceSheet.getRange(2, 1, lastSourceRow-1, lastSourceCol-1);
    var sourceData = sourceRange.getValues();

    var activeRow = 0;

    //Loop through every retrieved row from the Source
    for (row in sourceData) {

      var userPrimaryEmail = sourceData[row][0];
      var phoneValue = sourceData[row][1];
      var jobtitle = sourceData[row][2]

       Logger.log('User info: %s ', userPrimaryEmail,' & title:', jobtitle);

      var user = AdminDirectory.Users.get(userPrimaryEmail);
        //Take a look at what data is returned
       // Logger.log('User data:\n %s', JSON.stringify(user, null, 2));

      // blank out the phones
      user.phones = [];
      user.phones.push(
        {
          value: phoneValue,
          type: "mobile" // Could be 'home' or 'work' of whatever is allowed
        }
      )

      try {
        // remove title
        user.organizations[0].title =[];
        //Logger.log('titles variable emptied!');

        // add the new title
        user.organizations[0].title =jobtitle;
      }
      catch (err) {
        Logger.log('ERROR in organization[0]: ',  err );

        // remove title
        user.organizations.title =[];
        // add the new title
        user.organizations.title =jobtitle;
      }

      AdminDirectory.Users.update(user, userPrimaryEmail);
      Logger.log('User update complete for: ',  userPrimaryEmail );

    }

  }
  catch (err) {
    Logger.log('ERROR: ',  err );
  }
}



function reset_user_phones(user,userPrimaryEmail,phoneValue){

  try {

    var sourceSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
    var lastSourceRow = sourceSheet.getLastRow();
    var lastSourceCol = sourceSheet.getLastColumn();

    var sourceRange = sourceSheet.getRange(2, 1, lastSourceRow-1, lastSourceCol-1);
    var sourceData = sourceRange.getValues();

    var activeRow = 0;

    //Loop through every retrieved row from the Source
    for (row in sourceData) {

      var userPrimaryEmail = sourceData[row][0];
      var phoneValue = sourceData[row][1];
      var jobtitle = sourceData[row][2]

      // Logger.log('User info: %s ', userPrimaryEmail,' & title:', jobtitle);

      var user = AdminDirectory.Users.get(userPrimaryEmail);
      // Logger.log('User data:\n %s', JSON.stringify(user, null, 2));
      // If user has no phones add a 'phones' empty list to the user resource
      // if there is an existing phone remove it
      if (user.phones){
        try {
          Logger.log('resetting exising phones');

          user.phones = [];
          user.phones.value == phoneValue;
          user.phones.type == "mobile" ;
        }
        catch (errs) {
          Logger.log('No exising phines: ', errs);
        }
      }

      if (! user.phones ){
        Logger.log('pushing phones');

        user.phones = [];
        user.phones.push(
          {
            value: phoneValue,
            type: "mobile" // Could be 'home' or 'work' of whatever is allowed
          }
        )
      }

      Logger.log('User updating phones ',  userPrimaryEmail );

    }

  }
  catch (err) {
    Logger.log('ERROR: ',  err );
  }

  AdminDirectory.Users.update(user, userPrimaryEmail);
}

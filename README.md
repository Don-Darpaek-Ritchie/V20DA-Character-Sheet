You can use the version at https://don-darpaek-ritchie.github.io/V20DA-Character-Sheet/ whenever you want (unless Github says otherwise).  The information is all saved locally on your computer and you and your group can download and upload JSON's to each other every session to share updates.  No one else using the https://don-darpaek-ritchie.github.io/V20DA-Character-Sheet/ can see your characters and you cannot see anyone else's (except for my sample character).

However, you can also host your own personal version of the V20DA-Character-Sheet where you and your group can see each other's characters (or not, without a password) and update them in real time.  This is great if your group plays online and I promise it will only take you a few minutes to set up.

I know this looks big and complicated, but I promise it is not.  You will get a scary pop-up about Google Access and Unapproved Websites that you will have to click "Advanced" and then "Approve".  Worry not.  You can see the code below that you will paste.  I do not care what you keep in your Google Drive and have no interest to nose around in there.

1.  Create a new, blank Google Sheet.  Uploaded images will be saved in a folder in your Google Drive called "V20DA Interactive Character Sheet Images" (no quotes).  You might want to save your Google Sheet in a folder with that name to keep everything tidy in your Google Drive.
2.  In the menu of the Google Sheet you just created, go to Extensions > Apps Script.
3.  Delete any default code in the Code.gs editor and paste the script below and save.
4.  Click Deploy > New deployment.
5.  Deployment Options: ensure Type is Web app, Execute As is Me, and Who Has Access is Anyone.
6.  Click Deploy and save the Web Apps URL somewhere handy.  The URL should end with /exec/.
7.  Open your copy of index.html in any text editor and search for "GOOGLE" (no quotes).
8.  Replace the YOUR_DEPLOYED_URL_HERE (keep the quotes around your link) in [const API_URL = "YOUR_DEPLOYED_URL_HERE";] with the your /exec/ Google Sheet Web App URL.
9.  Host your modified index.html anywhere, I think.  I built it and have good success hosting it in a Github repository.  The web site images are hosted through Github's partner CDN so I don't think they will notice us using them and character images are hosted in your Google Drive.

// --- SCRIPT TO PASTE INTO GOOGLE SHEET CODE.GS ---

function getActiveCharSheet(ss) {
  var sheet = ss.getSheetByName("Active");
  if (!sheet) {
    var firstSheet = ss.getSheets()[0];
    if (firstSheet && firstSheet.getName() !== "Deleted") {
      firstSheet.setName("Active");
      sheet = firstSheet;
    } else {
      sheet = ss.insertSheet("Active");
      sheet.appendRow(["ID", "Name", "Clan", "Road", "Full JSON", "Is NPC", "NPC Password"]);
    }
  }
  return sheet;
}

function getDeletedCharSheet(ss) {
  var sheet = ss.getSheetByName("Deleted");
  if (!sheet) {
    sheet = ss.insertSheet("Deleted");
    sheet.appendRow(["ID", "Name", "Clan", "Road", "Full JSON", "Is NPC", "NPC Password", "Deleted Date"]);
  }
  return sheet;
}

function getOrCreateDriveFolder() {
  var folderName = "V20DA Interactive Character Sheet Images";
  var folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    return DriveApp.createFolder(folderName);
  }
}

// --- GET REQUEST HANDLER (FETCH DATA) ---

function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = getActiveCharSheet(ss);
  var data = activeSheet.getDataRange().getValues();
  
  if (e && e.parameter && e.parameter.id) {
    var searchId = e.parameter.id;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == searchId) {
        var jsonStr = data[i][4]; // Full JSON in Column E
        return ContentService.createTextOutput(jsonStr)
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ error: "Not found" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  var list = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] && data[i][0] !== "ID") {
      list.push({
        id: data[i][0],
        name: data[i][1],
        clan: data[i][2],
        road: data[i][3],
        isNpc: data[i][5] === true || data[i][5] === "true",
        npcPassword: data[i][6] || "",
        lastSaved: new Date().getTime()
      });
    }
  }
  
  return ContentService.createTextOutput(JSON.stringify(list))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- POST REQUEST HANDLER (SAVE / UPLOAD / DELETE) ---

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var activeSheet = getActiveCharSheet(ss);
  var deletedSheet = getDeletedCharSheet(ss);
  
  var payload = JSON.parse(e.postData.contents);

// --- DELETE PORTRAIT FROM GOOGLE DRIVE ---
  if (payload.action === "deleteImage" && payload.charId) {
    try {
      var folder = getOrCreateDriveFolder();
      var prefix = payload.charId + "_";
      var files = folder.getFiles();
      while (files.hasNext()) {
        var file = files.next();
        if (file.getName().indexOf(prefix) === 0) {
          file.setTrashed(true); // Trashes the portrait file in Google Drive
        }
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "image_deleted" }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- UPLOAD PORTRAIT TO GOOGLE DRIVE ---
  if (payload.action === "uploadImage") {
    try {
      var folder = getOrCreateDriveFolder();
      var charId = payload.charId || "char";
      var rawFileName = payload.imageName || "portrait.png";
      
      // Prefix filename with charId so each character has their own file
      var fileName = charId + "_" + rawFileName;

      // Delete any duplicate files with this exact name in the folder first
      var existingFiles = folder.getFilesByName(fileName);
      while (existingFiles.hasNext()) {
        var fileToDelete = existingFiles.next();
        fileToDelete.setTrashed(true);
      }

      // Extract image content type and raw base64 data
      var base64Data = payload.imageData;
      var contentType = base64Data.substring(base64Data.indexOf(":") + 1, base64Data.indexOf(";"));
      var base64Image = base64Data.substring(base64Data.indexOf(",") + 1);
      var blob = Utilities.newBlob(Utilities.base64Decode(base64Image), contentType, fileName);

      // Create file in host's Google Drive folder
      var newFile = folder.createFile(blob);
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      
      var fileId = newFile.getId();
      var imageUrl = "https://lh3.googleusercontent.com/d/" + fileId;

      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        imageUrl: imageUrl,
        imageName: rawFileName
      })).setMimeType(ContentService.MimeType.JSON);

    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "error",
        message: err.toString()
      })).setMimeType(ContentService.MimeType.JSON);
    }
  }

  // --- HANDLE CHARACTER DELETION ---
  if (payload.action === "delete" && payload.id) {
    var data = activeSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == payload.id) {
        var rowToMove = data[i];
        
        deletedSheet.appendRow([
          rowToMove[0],
          rowToMove[1],
          rowToMove[2],
          rowToMove[3],
          rowToMove[4],
          rowToMove[5],
          rowToMove[6],
          new Date().toISOString()
        ]);

        activeSheet.deleteRow(i + 1);

        return ContentService.createTextOutput(JSON.stringify({ status: "archived_to_deleted" }))
          .setMimeType(ContentService.MimeType.JSON);
      }
    }
    return ContentService.createTextOutput(JSON.stringify({ status: "not_found_in_active" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // --- SAVE / UPDATE CHARACTER DATA ---
  var charId = payload.id;
  var name = payload.meta ? payload.meta.name : "";
  var clan = payload.meta ? payload.meta.clan : "";
  var road = payload.mechanics ? payload.mechanics.roadName : "";
  var fullJson = JSON.stringify(payload);
  var isNpc = payload.isNpc || false;
  var npcPassword = payload.npcPassword || "";

  var data = activeSheet.getDataRange().getValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == charId) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex > 0) {
    activeSheet.getRange(rowIndex, 1, 1, 7).setValues([[charId, name, clan, road, fullJson, isNpc, npcPassword]]);
  } else {
    activeSheet.appendRow([charId, name, clan, road, fullJson, isNpc, npcPassword]);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- STOP COPYING HERE TO PASTE INTO GOOGLE SHEET CODE.GS ---



This Web App requires some recent version of Java, so if you buy drugs or slaves off of the Dark Web, you are screwed.  You should be supporting local entrepreneurs anyways.

How this works:

1.  Creation & Syncing: Every time a user creates or edits a character, it is automatically written to the Active sheet.
2.  Deletion: When a user clicks Delete on the Dashboard, doPost finds the character in the Active sheet, copies its entire JSON and metadata over to the Deleted sheet along with a deletion timestamp, and then deletes the row from Active.
3.  Global Consistency: Every player opening the Web App reads directly from Active, meaning all users immediately see the exact same list of active characters with the same up to date stats (give or take 1-3 seconds on Google Sheets updating).

Note on Deployments: In Google Apps Script, whenever you make changes to your spreadsheet code, you must select Deploy → Manage Deployments → Edit and choose New Version to ensure the live URL updates to run your latest edits.  You may have to update the line in your index.html where your Google Sheet /exec/ URL is located.

Limited Use:  Yes, I used Google AI Studio to put this together.  Did you think I actually typed all this myself?  If you have a problem with this, go throw a wrench in a gig mill.  For the rest of you, feel free to modify away if you are hosting a copy locally for your group.  Please do not modify anything in this and repost it somewhere else claiming it as your own.  I storytell a V20DA chronicle right now, I am not involved in a V20 modern game right now, and I do not play V5 (or 6?), but I am interested in modifying this for V20 modern if there is interest.  I am considering trying to redo it in V5 because that seems to be the most popular version now (until 6, but then the 5-ers will refuse to play 6 just like I refuse to play 5) but I am not terribly familiar with the rules and I would want to do the work well if I chose to.  

Either way, I do not own the rights to online character generators or particular versions of Vampire, so if you want to make a similar application - go for it!  Do not use my work as your code base, please.  As far as I can tell at the time of writing this, Mr. Gone's PDFs are still the gold standard for V20.  There look like there are a couple modern V20 online character generators out there (ShreckNet, trechkalov), but they look incomplete.  There are already a bunch of V5 online generators out there, including a semi-official one?  Anyways, I will be ticked if a bunch of V20 modern character generators pop-up in the near future using some very familiar code.

I know, my UI sucks.  Print mode sucks.  This thing is practically unreadable on phones.  If there are people out there that could give me some pointers on these issues, I would enjoy the assist.  Also, I have played a lot of Vampire back in the day, but I am sure I got some information wrong somewhere in here.  You will probably notice from Darpaek that we houserule languages because the merits are clunky.  Typos!  I do that a lot, too.  Email me at ritchie.don@gmail.com if you have suggestions or find bugs.

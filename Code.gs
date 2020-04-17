// Gmail2GDrive
// https://github.com/ahochsteger/gmail2gdrive

/**
 * Recursive function to create and return a complete folder path.
 */
function getOrCreateSubFolder(baseFolder,folderArray) {
  if (folderArray.length == 0) {
    return baseFolder;
  }
  var nextFolderName = folderArray.shift();
  var nextFolder = null;
  var folders = baseFolder.getFolders();
  while (folders.hasNext()) {
    var folder = folders.next();
    if (folder.getName() == nextFolderName) {
      nextFolder = folder;
      break;
    }
  }
  if (nextFolder == null) {
    // Folder does not exist - create it.
    nextFolder = baseFolder.createFolder(nextFolderName);
  }
  return getOrCreateSubFolder(nextFolder,folderArray);
}

/**
 * Returns the GDrive folder with the given path.
 */
function getFolderByPath(path) {
  var parts = path.split("/");

  if (parts[0] == '') parts.shift(); // Did path start at root, '/'?

  var folder = DriveApp.getRootFolder();
  for (var i = 0; i < parts.length; i++) {
    var result = folder.getFoldersByName(parts[i]);
    if (result.hasNext()) {
      folder = result.next();
    } else {
      throw new Error( "folder not found." );
    }
  }
  return folder;
}

/**
 * Returns the GDrive folder with the given name or creates it if not existing.
 */
function getOrCreateFolder(folderName) {
  var folder;
  try {
    folder = getFolderByPath(folderName);
  } catch(e) {
    var folderArray = folderName.split("/");
    folder = getOrCreateSubFolder(DriveApp.getRootFolder(), folderArray);
  }
  return folder;
}

/**
 * Sanitize filenames
 */
function getSanitizedFilename(name) {
  return name.replace(/[^a-zA-Z0-9 .@]/g, "");
}

/**
 * Unzip file
 */
function unZip(blob) {
  blob.setContentType("application/zip");
  var unzipped = Utilities.unzip(blob);
  var filename = unzipped[0].getName();
  return {
    filename: filename,
    content: unzipped[0],
  };
}

/**
 * Processes a message
 */
function processMessage(message, rule, config) {
  Logger.log("INFO:       Processing message: " + message.getSubject() + " (" + message.getId()  + ") from " + message.getFrom());
  var messageDate = message.getDate();

  // Adjust message date by optional offset in days
  if (typeof rule.dateOffset !== 'undefined') {
    messageDate.setDate( messageDate.getDate() - rule.dateOffset );
  }

  var attachments = message.getAttachments();
  for (var attIdx=0; attIdx<attachments.length; attIdx++) {
    var attachment = attachments[attIdx];

    try {
      var folder = getOrCreateFolder(Utilities.formatDate(messageDate, config.timezone, rule.driveFolder));
      var attachmentFilename = attachment.getName();
      var driveFilename = attachmentFilename;

      Logger.log("INFO:         Processing attachment: " + attachmentFilename);

      // Stop processing if `filenameFilter` is defined and its regex doesn't match attachment name
      if ( (rule.filenameFilter) && (!attachmentFilename.match(RegExp(rule.filenameFilter))) ) {
        Logger.log("INFO:           Rejecting file '" + attachmentFilename + "' not matching " + rule.filenameFilter);
        continue;
      }

      // Unzip if defined
      if (rule.unzip == true) {
        var unzipped = unZip(attachment);
        attachment = unzipped.content;
        driveFilename = unzipped.filename;
        Logger.log("INFO:           Unzipping file '" + attachmentFilename + "' as '" + driveFilename + "'");
      }

      // Set new filename if `driveFilename` is defined
      if (rule.driveFilename) {
        driveFilename = Utilities.formatDate(messageDate, config.timezone, rule.driveFilename
                                             .replace('%s',getSanitizedFilename(message.getSubject()))
                                             .replace('%f',getSanitizedFilename(message.getFrom()))
                                             .replace('%n',attachmentFilename));
      }

      // Detect if file already exists, if so, append index
      var index = 0;
      var indexedDriveFilename = driveFilename;
      while (folder.getFilesByName(indexedDriveFilename).hasNext()) {
        indexedDriveFilename = driveFilename.replace(/\.(?=[^.]*$)/, '-' + ++index + '.');
      }
      driveFilename = indexedDriveFilename;

      // Create file
      var file = folder.createFile(attachment);
      file.setName(driveFilename);
      file.setDescription("Mail title: " + message.getSubject() + "\nMail date: " + message.getDate() + "\nMail link: https://mail.google.com/mail/u/0/#inbox/" + message.getId());
      Logger.log("INFO:           Attachment '" + attachmentFilename + "' saved as '" + driveFilename + "'");

      Utilities.sleep(config.sleepTime);
    } catch (e) {
      Logger.log(e);
    }
  }
}

/**
 * Generate HTML code for one message of a thread.
 */
function processThreadToHtml(thread) {
  Logger.log("INFO:   Generating HTML code of thread '" + thread.getFirstMessageSubject() + "'");
  var messages = thread.getMessages();
  var html = "";
  for (var msgIdx=0; msgIdx<messages.length; msgIdx++) {
    var message = messages[msgIdx];
    html += "From: " + message.getFrom() + "<br />\n";
    html += "To: " + message.getTo() + "<br />\n";
    html += "Date: " + message.getDate() + "<br />\n";
    html += "Subject: " + message.getSubject() + "<br />\n";
    html += "<hr />\n";
    html += message.getBody() + "\n";
    html += "<hr />\n";
  }
  return html;
}

/**
* Generate a PDF document for the whole thread using HTML form.
 */
function processThreadToPdf(thread, rule) {
  Logger.log("INFO: Saving PDF copy of thread '" + thread.getFirstMessageSubject() + "'");
  var folder = getOrCreateFolder(rule.driveFolder);
  var html = processThreadToHtml(thread);
  var blob = Utilities.newBlob(html, 'text/html');
  var pdf = folder.createFile(blob.getAs('application/pdf')).setName(thread.getFirstMessageSubject() + ".pdf");
  return pdf;
}

/**
 * Main function that processes Gmail attachments and stores them in Google Drive.
 * Use this as trigger function for periodic execution.
 */
function Gmail2GDrive() {
  if (!GmailApp) return; // Skip script execution if GMail is currently not available (yes this happens from time to time and triggers spam emails!)
  var config = getGmail2GDriveConfig();
  var label = GmailApp.createLabel(config.processedLabel);
  var end, start, runTime;
  start = new Date(); // Start timer

  Logger.log("INFO: Starting mail attachment processing.");
  if (config.globalFilter===undefined) {
    config.globalFilter = "has:attachment -in:trash -in:drafts -in:spam";
  }

  // Iterate over all rules:
  for (var ruleIdx=0; ruleIdx<config.rules.length; ruleIdx++) {
    var rule = config.rules[ruleIdx];
    var gSearchExp  = config.globalFilter + " " + rule.filter + " -label:" + config.processedLabel;
    if (config.newerThan != "") {
      gSearchExp += " newer_than:" + config.newerThan;
    }
    var doArchive = rule.archive == true;
    var doPDF = rule.saveThreadPDF == true;

    // Process all threads matching the search expression:
    var threads = GmailApp.search(gSearchExp);
    Logger.log("INFO:   Processing rule: " + gSearchExp);
    for (var threadIdx=0; threadIdx<threads.length; threadIdx++) {
      var thread = threads[threadIdx];
      end = new Date();
      runTime = (end.getTime() - start.getTime())/1000;
      Logger.log("INFO:     Processing thread: " + thread.getFirstMessageSubject() + " (runtime: " + runTime + "s/" + config.maxRuntime + "s)");
      if (runTime >= config.maxRuntime) {
        Logger.log("WARNING: Self terminating script after " + runTime + "s");
        return;
      }

      // Process all messages of a thread:
      var messages = thread.getMessages();
      for (var msgIdx=0; msgIdx<messages.length; msgIdx++) {
        var message = messages[msgIdx];
        processMessage(message, rule, config);
      }
      if (doPDF) { // Generate a PDF document of a thread:
        processThreadToPdf(thread, rule);
      }

      // Mark a thread as processed:
      thread.addLabel(label);

      if (doArchive) { // Archive a thread if required
        Logger.log("INFO:     Archiving thread '" + thread.getFirstMessageSubject() + "' ...");
        thread.moveToArchive();
      }
    }
  }
  end = new Date(); // Stop timer
  runTime = (end.getTime() - start.getTime())/1000;
  Logger.log("INFO: Finished mail attachment processing after " + runTime + "s");
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Comment Extractor')
    .addItem('Import Comments from Doc', 'promptForDocId')
    .addToUi();
}

/**
 * Prompts the user for a Google Doc URL or ID.
 */
function promptForDocId() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Enter Google Doc URL or ID',
    'Please paste the full URL or the ID of the Google Doc you want to extract comments from:',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() == ui.Button.OK) {
    const input = response.getResponseText().trim();
    const docId = extractIdFromUrl(input);
    
    if (docId) {
      extractComments(docId);
    } else {
      ui.alert('Invalid URL or ID provided.');
    }
  }
}

function extractComments(fileId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const ui = SpreadsheetApp.getUi();
  
  // Set sheet headers
  const headers = [
    'Date Created', 
    'Author', 
    'Status', 
    'Highlighted Text (Context)', 
    'Comment Thread'
  ];
  
  // Check if headers exist, if not, add them
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  try {    
    let pageToken = null;
    let commentsList = [];

    do {
      const response = Drive.Comments.list(fileId, {
        fields: 'nextPageToken, comments(author(displayName), createdTime, content, resolved, quotedFileContent, replies(author(displayName), content, createdTime))',
        pageToken: pageToken,
        pageSize: 100
      });
      
      if (response.comments) {
        commentsList = commentsList.concat(response.comments);
      }
      pageToken = response.nextPageToken;
    } while (pageToken);

    if (commentsList.length === 0) {
      ui.alert('No comments found in this document.');
      return;
    }

    const rows = [];
    
    // Process each comment
    commentsList.forEach(function(comment) {
      const author = comment.author ? comment.author.displayName : 'Unknown';
      const date = new Date(comment.createdTime);
      const status = comment.resolved ? 'resolved' : 'open';
      const content = comment.content || '';
      
      // Get the text the comment is attached to (Anchor text)
      const context = (comment.quotedFileContent && comment.quotedFileContent.value) ? comment.quotedFileContent.value : '[General Comment]';
      
      // Format the full thread (Main comment + Replies)
      let fullThread = `${content}`; // add author name here if want to show in comments
      
      if (comment.replies && comment.replies.length > 0) {
        comment.replies.forEach(function(reply) {
          const rAuthor = reply.author ? reply.author.displayName : 'Unknown';
          fullThread += `\n\n${reply.content}`; // add author name here if you want that to populate in row
        });
      }

      rows.push([date, author, status, context, fullThread]);
    });

    // Write all rows to the sheet at once
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    
    ui.alert(`Success! Extracted ${rows.length} comment threads.`);

  } catch (e) {
    Logger.log(e);
    ui.alert('Error: ' + e.message + '\n\nMake sure you enabled the "Drive API" service (v3) in the script editor.');
  }
}

/**
 * Helper to extract ID from a full URL or return the input if it's already an ID.
 */
function extractIdFromUrl(url) {
  // Regex to catch ID in standard Docs URLs
  const match = url.match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

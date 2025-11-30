# **Google Docs Comment Extractor**

A Google Apps Script tool that extracts all comments (open and resolved) from a Google Doc and exports them into a structured Google Sheet. It captures the full conversation thread, author details, timestamps, and the specific text context the comment refers to.

## **Features**

* **Full History**: Extracts both **Open** and **Resolved** comments.  
* **Context Aware**: Captures the specific **highlighted text** (anchor) the comment is attached to.  
* **Threaded View**: Formats the main comment and all subsequent replies into a single, readable cell.  
* **Modern API**: Built using Google Drive API v3.

## **Prerequisites**

* A Google account.  
* A Google Sheet to host the script and data.  
* Edit access to the Google Doc you wish to analyze.

## **Installation**

### **1\. Create the Script**

1. Open a new or existing Google Sheet.  
2. Go to Extensions \> Apps Script.  
3. Copy the code from extract\_comments.gs and paste it into the script editor, replacing any existing code.  
4. Save the project (Disk icon).

### **2\. Enable Drive API Service (Critical)**

**Note:** This script uses the Advanced Drive Service. It will not work without this step.

1. In the Apps Script editor, look at the left sidebar.  
2. Click the **\+** button next to **Services**.  
3. Scroll down and select **Drive API**.  
4. Ensure the version is set to **v3**.  
5. Click **Add**.

## **Usage**

1. Refresh your Google Sheet.  
2. Wait a few seconds for the custom menu to appear in the toolbar.  
3. Click **Comment Extractor** \> **Import Comments from Doc**.  
4. A prompt will appear. Paste the **URL** or **ID** of the Google Doc.  
5. Click **OK**.  
6. The script will ask for permission to run the first time. Authorize it.

The script will automatically create headers (if the sheet is empty) and populate the rows with the comment data.

## **Troubleshooting**

| Error Message | Cause and Solution |
| ----- | ----- |
| **"Drive is not defined"** | **Cause:** The Drive API is not enabled in the Services menu. **Solution:** See Installation step 2 to enable the API. |
| **"Invalid field selection items"** | **Cause:** You are using an outdated version of the script. **Solution:** Update your script to the version in this repository, which uses the newer comments (API v3) instead of items (API v2). |
| **"No comments found"** | **Cause:** The document either contains no comments, or you lack the necessary permission to view them. **Solution:** Verify the document contains comments and that you have appropriate viewing access. |

## **Output Format**

The script generates the following columns:

1. **Date Created**: Timestamp of the initial comment.  
2. **Author**: Display name of the comment creator.  
3. **Status**: open or resolved.  
4. **Highlighted Text (Context)**: The text in the document that was highlighted.  
5. **Comment Thread**: The full conversation history.

## **License**

This project is licensed under the MIT License \- see the LICENSE file for details.

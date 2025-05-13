Here’s the same solution presented in **Markdown** format, suitable for documentation or a README file:

---

# 📧 Extract Email Addresses from Outlook Using Batch and VBScript

This guide shows how to extract sender email addresses from your Outlook inbox using a batch script and VBScript.

---

## ✅ Overview

A batch file on its own can't access Outlook data directly, but we can use it to **call a VBScript** that interacts with Outlook via COM automation. The VBScript extracts sender email addresses from the inbox and writes them to a text file.

---

## 🧩 Files Needed

### 🔹 `extractEmails.vbs`

This script connects to Outlook, iterates over inbox emails, and writes sender email addresses to a file:

```vbscript
Dim olApp, olNs, inbox, item, fso, outputFile, emailAddress
Set olApp = CreateObject("Outlook.Application")
Set olNs = olApp.GetNamespace("MAPI")
Set inbox = olNs.GetDefaultFolder(6) ' 6 = olFolderInbox

Set fso = CreateObject("Scripting.FileSystemObject")
Set outputFile = fso.CreateTextFile("email_addresses.txt", True)

For Each item In inbox.Items
    If item.Class = 43 Then ' 43 = MailItem
        emailAddress = item.SenderEmailAddress
        outputFile.WriteLine emailAddress
    End If
Next

outputFile.Close
MsgBox "Extraction complete. Saved to email_addresses.txt."
```

---

### 🔹 `run_extract.bat`

This batch file executes the VBScript:

```bat
@echo off
cscript //nologo extractEmails.vbs
pause
```

---

## 🚀 How to Use

1. Save the VBScript as `extractEmails.vbs`.
2. Save the batch file as `run_extract.bat` in the same folder.
3. Double-click `run_extract.bat`.
4. The script will create `email_addresses.txt` containing all sender email addresses from your Outlook inbox.

---

## 📁 Output

* File: `email_addresses.txt`
* Content: One email address per line from the **Inbox** sender field.

---

## ⚙️ Notes

* Requires Microsoft Outlook to be installed and configured.
* Works with classic Outlook, **not** the new "Outlook (New)" app.
* Extracts **Inbox** only. For other folders like **Sent Items**, the script can be modified.
* To include **recipients** (To/CC), ask for an enhanced version of the script.

---

Let me know if you’d like a version that:

* Extracts from another folder (like Sent Items),
* Includes recipients (To/CC),
* Removes duplicates, or
* Saves the output as `.csv`.

Happy scripting! 💻

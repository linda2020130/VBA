# VBA Collections
A place to collect all the excel/Outlook VBA works written so far, making it easier for future look up.
<br>
## Table of Contents
* [VBA Syntax Notes](#Notes)
* [Copy Sheets](#CopySheets)
* [Create Folder](#CreateFolder)
* [Generate Email](#GenerateEmail)
* [Save Htm File](#SaveHtmFile)
* [Send Drafts](#SendDrafts)
<br><br>

<h2 id="Notes">VBA Syntax Notes</h2>

[Link](/VBA_SyntaxNotes.ipynb)
* **Usages** : Quickly look up common use syntax.
* **Example** : Copy and paste, delete, modify, autofill, filter, sort, find, replace cells/range. Conditions and loop...etc
<br>

<h2 id="CopySheets">Copy Sheets</h2>

[Link](/Copy_Sheets.bas)
* **Usage** : Copy worksheets in current excel workbook to another.
* **Example** : Copy sheets("Summary", "table1", "table2") from Test.xlsm to new workwork and save as "Summary_YYYYMMDD.xlsx".
<br>

<h2 id="CreateFolder">Create Folder</h2>

[Link](/Create_Folder.bas)
* **Usage** : Create a folder(directory) in desired path using excel VBA.
* **Example** : Create a folder at Desktop and name it as current year and month, and create a folder inside of it as current date. Do not create if already existed.
<br>

<h2 id="GenerateEmail">Generate Email</h2>

[Link](/Generate_Email.bas)
1. **Get Signature (function)**
    * **Usage** : Get outlook signature to be inserted to email content.
    * **Example** : Go to `C:\Users\lindac\AppData\Roaming\Microsoft\Signatures\` to get signature named as `Linda Chou.htm`.
2. **Get Htm (function)**
    * **Usage** : Get htm file to be inserted to email content. Put htm path including file name as input for GetHtm function.
    * **Example** : Go to certain path(input parameter) to get htm file.
3. **HtmlBody (Function)**
    * **Usage** : Generate email body in html format. Put receiver's name and time of downloading data as input for HtmlBody function.
    * **Example** : Generate two paragraphs with different format style(font-family, font-size, ...) as desired.
4. **Generate Email (Process)**
    * **Usage** : Generate drafts in outlook according to DataTable. Each email will be attached a certain file(input parameter) and be inserted a certain htm table(input parameter). Call out above functions if needed.
    * **Example** : Generate drafts in outlook according to worksheet "Receiver List". Attach file in the email (`attachment = path + file`). Insert htm table in the email content (`htm = path + file`).
<br>

<h2 id="SaveHtmFile">Save Htm File</h2>

[Link](/Save_htm_File.bas)
* **Usage** : Save table in Acurrent workbook as html file(htm file can be inserted to email content).
* **Example** : Save table in worksheet "Summary" as `Summary Table_YYYYMMDD` to folder with current year, month and date at Desktop.
<br>
 
<h2 id="SendDrafts">Send Draft</h2>

[Link](/Send_Drafts.bas)
1. **Send All Your Mail Box Drafts**
    * **Usage** : Outlook VBA for sending a batch of emails and input the mail box name.
    * **Example** : Call out SendAllDrafts and put mail box name `"linda2020130"` as input parameter.
2. **Send All Drafts**
    * **Usage** : Outlook VBA for sending a batch of emails
    * **Example** : Pop up notification for user to make sure sending out correct mailbox and sending out all drafts after clicking `Yes`.
<br> 

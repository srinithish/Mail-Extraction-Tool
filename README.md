# Mail-Extraction-Tool

#To install

1. Download the Zip of the repository

2. Import the "ExtractIncExp" file to Outlook under Developer 

3. Add references in tools:
   Visual Basic for Applicaitons,
   Microsoft Outlook 14.0,
   OLE Automation,
   Microsoft Office 14.0,
   Microsoft Excel 14.0,
   Microsoft Internet Controls,
   Microsoft Scripting Runtime
  
 4. Add database file in the required shared folder and add the same to trgWb (Line 36 in outlook ExtractInc module)
 
 5. Show the location of Temp folder (Line 91 in outlook ExtractInc module) where the macro will save and keep deleting the PMOs        (preferably Local path)
 
 6. Add folder in  Outlook under 'Inbox' as 'Testing'
 
 7. Add shortcut button for Macro "forUpdateIncExp" from  "UpdateIncExp.xlam" 
 
 
 #To Use
 
 1. Copy Emails to be read received to GES Mailbox and transfer them to testing folder 
 2. In Outlook > Developer > Macros> Run 'Mail.Items'
 3. After the mails are read a "Done" message appears
 4. Go to the Tracker input file and click on the 'forUpdateIncExp' button previously added 
 5. Show it the Database.xlsx file.
 
 
 ##Musts
 1. Keep the Database.xlsx file closed before running the Macro
 

###How to Add GES Mail Box is at 
\\pngsfsdg04\AnalyticsCOE\@CATModeling\@CATAccountModeling\@Teams\@International\@APAC\Archives\Mail Archive for APAC
# Mail-Extraction-Tool

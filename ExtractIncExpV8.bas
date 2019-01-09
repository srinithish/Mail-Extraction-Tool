Attribute VB_Name = "ExtractIncExp1"



Dim srcWb As Workbook
Dim trgWb As Workbook
Dim currentRow As Integer
Dim TransactionButtons() As String
'Dim testWb As Workbook

Sub MailItems()
'    On Error Resume Next
 TransactionButtons = Split("NewBusiness,Renewal,Endorsement", ",")
 Dim count As Integer
 count = 0
'    Dim wb As Workbook
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   Dim xl As Object
'    Set xl = New Excel.Application 'create session
'    'xl.Workbooks.Open FileName:="\\pngscitrix01\sk\Desktop\Automation\Test and Supporting Files" & "\" & "Database.xlsx"  'open wb in the new session
'    xl.Visible = True 'this is what you need, show it up!

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
  Excel.Application.ScreenUpdating = False
    Dim olNamespace As Outlook.NameSpace
    Dim olFolder  As Outlook.MAPIFolder
    Dim olItem As Outlook.MailItem
    Dim allItems As Object
    Set olNamespace = Application.GetNamespace("MAPI")
    Set olFolder = olNamespace.GetDefaultFolder(olFolderInbox).Folders("Testing") '''''''insert the folder name in which to run this script
'
'    On Error GoTo ErrHandler1
    
    Set olItem = Application.CreateItem(olMailItem) ' Creat EMail+
      Excel.Application.AskToUpdateLinks = False
    Excel.Application.DisplayAlerts = False
    Set trgWb = Workbooks.Open("\\pngsfsdg04\AnalyticsCOE\@CATModeling\@CATAccountModeling\@Teams\@International\@APAC\Back end Work Tracker\E-task reports\2017\Database\Database.xlsx") '''path of the workbook where the data is to be dumped
    
'    myItems = olFolder.Items
'    myItems.Sort "[ReceivedTime]"
   
      olFolder.Items.Sort "[ReceivedTime]", True
    With olItem
        For Each allItems In olFolder.Items
            If allItems.Class = olMail Then
            
                count = count + 1
            
                If count <= 200 Then
                
                    Debug.Print ("")
                  Debug.Print ("Reading  " & allItems.Subject)
                  
                
                    Call saveAttachtoDisk(allItems)
                    
                Else
                    MsgBox ("first " & (count - 1) & " delete them")
                    Exit For
                End If
                
            End If
            
        Next
        
    End With
    trgWb.Save
    trgWb.Close
    MsgBox ("Done")
    
    Excel.Application.DisplayAlerts = True
    Excel.Application.AskToUpdateLinks = True
    Excel.Application.ScreenUpdating = True
    Exit Sub

ErrHandler1:
    trgWb.Close
    MsgBox ("An unknown error occured please delete the files already extracted and rerun the macro,contact nithish")
    Excel.Application.DisplayAlerts = True
    Excel.Application.AskToUpdateLinks = True
    Excel.Application.ScreenUpdating = True
End Sub


Public Sub saveAttachtoDisk(itm As Outlook.MailItem)
Dim tempFileName As String
tempFileName = "Temp.xls"
'    On Error Resume Next
    Dim objAtt As Outlook.Attachment
    Dim saveFolder As String
    
    saveFolder = "\\pngscitrix01\sk\Desktop\Automation\Test and Supporting Files" 'temp folder to save workbooks
'    dateFormat = Format(Now, "yyyy-mm-dd H-mm")
'    Debug.Print (itm.ReceivedTime)
    
    For Each objAtt In itm.Attachments
    
    
        If objAtt.DisplayName Like ("*.xls") Or objAtt.DisplayName Like ("*.XLS") Then ''''check for PMO and Excel File
                
                objAtt.SaveAsFile saveFolder & "\" & tempFileName ''' saving the file
    
            '''''open and extract the fields'''''''''''''''''''''''''''
    
                Set srcWb = Workbooks.Open(saveFolder & "\" & tempFileName, UpdateLinks:=False)
                Debug.Print ("name of the workbook is " & objAtt.DisplayName)
                flag = CheckIfSheetExists(" Policy Model Options") ''2nd Check for PMO file
                Debug.Print ("Its a " & flag & " PMO")
                    On Error GoTo ErrHandler:
                    If flag = True Then   ' check for PMO sheet
                                
                                myFileName = objAtt.DisplayName
                                srcWb.UnProtect Password1
'
                                srcWb.Sheets(" Policy Model Options").UnProtect Password1
'                                Debug.Print ("password unprotected")
                                
                                a = getNextRow(currentRow)
                            
                                ''''''extract required fields
                                trgWb.Sheets("Sheet1").Cells(currentRow, 1).Value = srcWb.Sheets(" Policy Model Options").Range("AccountName").Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 2).Value = srcWb.Sheets(" Policy Model Options").Range("EffectiveDate").Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 3).Value = srcWb.Sheets(" Policy Model Options").Range("ExpirationDate").Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 4).Value = srcWb.Sheets(" Policy Model Options").OLEObjects("ML").Object.Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 5).Value = srcWb.Sheets(" Policy Model Options").OLEObjects("MinorL").Object.Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 6).Value = itm.Subject
                                trgWb.Sheets("Sheet1").Cells(currentRow, 7).Value = itm.ReceivedTime
                                trgWb.Sheets("Sheet1").Cells(currentRow, 8).Value = srcWb.Sheets(" Policy Model Options").Range("C18").Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 9).Value = srcWb.Sheets(" Policy Model Options").OLEObjects("ComboBox3").Object.Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 10).Value = srcWb.Sheets(" Policy Model Options").Range("TargetDate").Value
                                trgWb.Sheets("Sheet1").Cells(currentRow, 11).Value = getSelectedOption(TransactionButtons)
                                trgWb.Save
                                Debug.Print ("written to row " & currentRow)
                                
                    End If
                    srcWb.Close False
                    b = KillFile(saveFolder & "\" & tempFileName)
                ''''''''''''''check for errors'''''''''
                
                    If b = False Then
'                    b = KillFile(saveFolder & "\" & srcWb.Name)
                    Debug.Print ("Couldnt  kill the file " & myFileName)
                    End If
                    
        End If
        
    Set objAtt = Nothing

    Next objAtt
    
    Exit Sub
    
ErrHandler:
   
    trgWb.Save
    srcWb.Close False
    
    b = KillFile(saveFolder & "\" & tempFileName)
    Set objAtt = Nothing
'    MsgBox ("error2")
    Debug.Print ("error in files in  " & myFileName & "  in email  " & itm.Subject)
'
End Sub



Public Function Password1() As String
    Password1 = "gespmo"
End Function

Private Sub saveDetailsToExcel()

End Sub


Private Function getNextRow(ByRef currentRow As Integer)

currentRow = Excel.Application.WorksheetFunction.CountA(trgWb.Sheets("Sheet1").Range("A:A"))
currentRow = currentRow + 1
End Function

Private Function KillFile(filePath As String) As Boolean
    If Len(Dir(filePath)) > 0 Then
        SetAttr filePath, vbNormal
        Kill filePath
        KillFile = True
        Exit Function
    End If
    KillFile = False
End Function


Private Function CheckIfSheetExists(SheetName As String) As Boolean
      CheckIfSheetExists = False
      For Each WS In srcWb.Worksheets
        If SheetName = WS.Name Then
          CheckIfSheetExists = True
          Exit Function
        End If
      Next WS
End Function

'Sub simple()
'Set testWb = Workbooks.Open("\\pngscitrix01\sk\Desktop\Automation\Test and Supporting Files\Locations List with SI (Updated on 06 April 2017) - for PMO.xlsx")
' If CheckIfSheetExists(" Policy Model Options") = True Then
' Debug.Print ("here")
'
' Else: Debug.Print ("Else")
' End If
'
'End Sub

Private Function getSelectedOption(groupName() As String) As String
    
    For Each element In groupName
'    Debug.Print (element)
       Value = srcWb.Sheets(" Policy Model Options").OLEObjects(element).Object.Value
       If Value = True Then
            Debug.Print (element)
            getSelectedOption = element
            Exit Function
        End If
    Next element
    
    getSelectedOption = "None"
End Function

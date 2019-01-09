Attribute VB_Name = "UpdateInceExpiry"
'version V7
'Updates : Extract and convert limits to USD,target date

Sub forUpdateIncepExpiry()
Attribute forUpdateIncepExpiry.VB_ProcData.VB_Invoke_Func = " \n14"
'
' forUpdateIncepExpiry Macro

    Dim CP_EER As String
    Dim TR_CAR As String
    Dim Limit As Double
    Dim LimitInUSD As Double
    Dim PMOCurrency As String
  Dim Ws As Worksheet
  Set Ws = Application.ActiveSheet
  
  Dim numRows  As Integer
  numRows = Application.WorksheetFunction.CountA(Range("B:B"))
  
  SourceFilePath = Application.GetOpenFilename()
    Set SourceWb1 = Workbooks.Open(SourceFilePath)
    
  Dim SourceWs As Worksheet
  
  Set SourceWs = SourceWb1.Sheets("Sheet1")
  
  Debug.Print (numRows)
   Debug.Print (Ws.Name)
   Debug.Print (SourceWs.Name)
  
   For i = 2 To numRows
'    Debug.Print ("U" & i)

    On Error Resume Next
    Ws.Range("U:U").NumberFormat = "mm/dd/yyyy"
    Ws.Range("V:V").NumberFormat = "mm/dd/yyyy"
    Ws.Range("K:K").NumberFormat = "mm/dd/yyyy"
    SourceWs.Range("H:H").NumberFormat = "0"
    With Application.WorksheetFunction
    
               Ws.Range("U" & i) = .Index(SourceWs.Range("B:B"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0)) ''inception
               Ws.Range("V" & i) = .Index(SourceWs.Range("C:C"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0)) '' expiry
               Ws.Range("K" & i) = .Index(SourceWs.Range("J:J"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0)) '' for requesetd date
               
               '''''''''''''''''''for transaction type
              
                Ws.Range("I" & i).Value = .Index(SourceWs.Range("K:K"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0))
               
               
               '''''''''''''''''''''''''''''for PMO Limits
               Limit = .Index(SourceWs.Range("H:H"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0))
               PMOCurrency = Left(.Index(SourceWs.Range("I:I"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0)), 3)
                
                If Limit <> 0 Then
                    LimitInUSD = Limit * AIGCurrencyConv(PMOCurrency, "USD")
                    Ws.Range("I" & i).Value = Ws.Range("I" & i).Value & "|" & "PMO limit in USD " & LimitInUSD
                End If
                
               PMOCurrency = ""
               Limit = 0
               ''''''''''''''''''''''''''''''EER and CP updating'''''''''''''''''''''''
               
               CP_EER = .Index(SourceWs.Range("D:D"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0))
                    If CP_EER Like ("*Commercial Property*") Then
                        Ws.Range("X" & i) = "CP"
                    ElseIf CP_EER Like ("*Energy & Engineered Risks*") Then
                        Ws.Range("X" & i) = "EER"
                    End If
                CP_EER = ""
                '''''''''''''''''''''terrorism and CAR accounts'''''''''''''''''''
               TR_CAR = .Index(SourceWs.Range("E:E"), .Match(Ws.Range("M" & i), SourceWs.Range("F:F"), 0))
                    If (TR_CAR Like ("*Stand Alone Terrorism*")) Or (TR_CAR Like ("*Political Violence*")) Then
                        Ws.Range("H" & i) = "Y"
                        Ws.Range("B" & i) = Ws.Range("B" & i).Value & " (Terrorism)"
                        Ws.Range("AD" & i) = "N"
                        ''''insert for Triage No for terrorism accounts
                    ElseIf TR_CAR Like ("*Construction*") Or TR_CAR Like ("CAR*2536*") Then
                        Ws.Range("W" & i) = "High"
                        Ws.Range("X" & i) = "EER: CAR/EAR"
                    End If
                TR_CAR = ""

    End With

    Next i

Ws.Activate

End Sub




Private Function AIGCurrencyConv(ByVal strFromCurrency, ByVal strToCurrency, Optional ByVal strResultType = "Value")
On Error GoTo errorHandler
 
'Init
Dim strURL As String
Dim objXMLHttp As Object
Dim strRes As String, dblRes As Double
 
Set objXMLHttp = CreateObject("MSXML2.ServerXMLHTTP")
strURL = "http://finance.yhoo.com/d/quotes.csv?e=.csv&f=c4l1&s=" & strFromCurrency & strToCurrency & "=X"
 
'Send XML request
With objXMLHttp
    .Open "GET", strURL, False
    .setRequestHeader "Content-Type", "application/x-www-form-URLEncoded"
    .Send
    strRes = .responseText
End With

'Parse response
dblRes = Val(Split(strRes, ",")(1))

Select Case strResultType
    Case "Value": AIGCurrencyConv = dblRes
    Case Else: AIGCurrencyConv = "1 " & strFromCurrency & " = " & dblRes & " " & strToCurrency
End Select
 
CleanExit:
    Set objXMLHttp = Nothing
 
Exit Function
 
errorHandler:
    AIGCurrencyConv = 0
    GoTo CleanExit
End Function

Sub simple()

End Sub

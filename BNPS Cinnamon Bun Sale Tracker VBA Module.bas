Attribute VB_Name = "Module1"
' ===================================================================
' BENTLEY NPS CINNAMON BUN SALE - PAYMENT TRACKER VBA
' Replicates Google Apps Script functionality in Excel
' ===================================================================

Option Explicit





' ========= MAIN CONTROL FUNCTIONS =========
Sub RunAllPaymentChecks()
    Application.ScreenUpdating = False
    
    Call ImportPayPalEmails
    Call ImportVenmoEmails
    Call ImportZelleEmails
    
    Application.ScreenUpdating = True
    MsgBox "All payment checks completed!", vbInformation, "NPS Tools"
End Sub

Sub CheckVenmoPayments()
    Call ImportVenmoEmails
End Sub

Sub CheckZellePayments()
    Call ImportZelleEmails
End Sub

Sub CheckPayPalPayments()
    Call ImportPayPalEmails
End Sub

' ========= CLEAR ALL SHEETS UTILITY =========
Sub ClearAllPaymentSheets()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim response As VbMsgBoxResult
    
    response = MsgBox("This will clear all data from Venmo, Zelle, and PayPal sheets. Continue?", _
                      vbYesNo + vbExclamation, "Confirm Clear")
    
    If response = vbNo Then Exit Sub
    
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "Venmo" Or ws.Name = "Zelle" Or ws.Name = "PayPal" Then
            lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            
            If lastRow > 1 Then
                ws.Range("A2:G" & lastRow).ClearContents
            End If
        End If
    Next ws
    
    MsgBox "All payment sheets cleared!", vbInformation, "NPS Tools"
End Sub

' ========= VENMO IMPORT FUNCTION =========
Sub ImportVenmoEmails()
    Dim ws As Worksheet
    Dim outlookApp As Object
    Dim outlookNS As Object
    Dim inbox As Object
    Dim items As Object
    Dim mail As Object
    
    Dim newRows As Collection
    Dim existingKeys As Object
    Dim lastRow As Long
    Dim i As Long
    Dim skippedCount As Long
    
    ' Get or create Venmo sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Venmo")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Venmo"
        Call SetupSheetHeaders(ws)
    End If
    
    ' Load existing data for deduplication
    Set existingKeys = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        For i = 2 To lastRow
            Dim key As String
            key = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 4).Value & "|" & ws.Cells(i, 3).Value
            existingKeys(key) = True
        Next i
    End If
    
    ' Connect to Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    Set outlookNS = outlookApp.GetNamespace("MAPI")
    Set inbox = outlookNS.GetDefaultFolder(6) ' olFolderInbox
    Set items = inbox.items
    
    ' Filter for Venmo emails
    items.Sort "[ReceivedTime]", True
    
    Set newRows = New Collection
    skippedCount = 0
    
    ' Process emails
    For Each mail In items
        If InStr(1, mail.SenderEmailAddress, "venmo@venmo.com", vbTextCompare) > 0 Or _
           InStr(1, mail.SenderName, "Venmo", vbTextCompare) > 0 Then
            
            ' Check subject contains "paid you"
            If InStr(1, mail.Subject, "paid you", vbTextCompare) > 0 Or _
               InStr(1, mail.Subject, "paid to your Venmo account", vbTextCompare) > 0 Or _
               InStr(1, mail.Subject, "paid your", vbTextCompare) > 0 Then
                
                ' Exclude outgoing payments
                If InStr(1, mail.Subject, "You paid", vbTextCompare) = 0 Then
                    Call ProcessVenmoEmail(mail, ws, newRows, existingKeys, skippedCount)
                End If
            End If
        End If
        
        ' Limit search to recent emails (last 100 days)
        If DateDiff("d", mail.ReceivedTime, Date) > 100 Then Exit For
    Next mail
    
    ' Write new rows to sheet
    If newRows.Count > 0 Then
        Call WriteRowsToSheet(ws, newRows)
        Call UpdateTotalBuns(ws, "H1", "H2")
        
        MsgBox "? VENMO: Added " & newRows.Count & " new transactions." & vbCrLf & _
               "Skipped " & skippedCount & " duplicates.", vbInformation, "Venmo Import"
    Else
        MsgBox "No new Venmo payments found." & vbCrLf & _
               "Skipped " & skippedCount & " existing transactions.", vbInformation, "Venmo Import"
    End If
    
    ' Cleanup
    Set outlookApp = Nothing
    Set outlookNS = Nothing
    Set inbox = Nothing
    Set items = Nothing
End Sub

' ========= PAYPAL IMPORT FUNCTION =========
Sub ImportPayPalEmails()
    Dim ws As Worksheet
    Dim outlookApp As Object
    Dim outlookNS As Object
    Dim inbox As Object
    Dim items As Object
    Dim mail As Object
    
    Dim newRows As Collection
    Dim existingKeys As Object
    Dim lastRow As Long
    Dim i As Long
    Dim skippedCount As Long
    
    ' Get or create PayPal sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("PayPal")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "PayPal"
        Call SetupSheetHeaders(ws)
    End If
    
    ' Load existing data for deduplication
    Set existingKeys = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        For i = 2 To lastRow
            Dim key As String
            key = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 4).Value & "|" & ws.Cells(i, 3).Value
            existingKeys(key) = True
        Next i
    End If
    
    ' Connect to Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    Set outlookNS = outlookApp.GetNamespace("MAPI")
    Set inbox = outlookNS.GetDefaultFolder(6)
    Set items = inbox.items
    items.Sort "[ReceivedTime]", True
    
    Set newRows = New Collection
    skippedCount = 0
    
    ' Process emails
    For Each mail In items
        If InStr(1, mail.SenderEmailAddress, "@paypal.com", vbTextCompare) > 0 Or _
           InStr(1, mail.SenderName, "PayPal", vbTextCompare) > 0 Then
            
            If (InStr(1, mail.Subject, "You received", vbTextCompare) > 0 Or _
                InStr(1, mail.Subject, "sent you", vbTextCompare) > 0 Or _
                InStr(1, mail.Subject, "Payment received", vbTextCompare) > 0) And _
               InStr(1, mail.Subject, "You sent", vbTextCompare) = 0 And _
               InStr(1, mail.Subject, "You paid", vbTextCompare) = 0 Then
                
                Call ProcessPayPalEmail(mail, ws, newRows, existingKeys, skippedCount)
            End If
        End If
        
        If DateDiff("d", mail.ReceivedTime, Date) > 100 Then Exit For
    Next mail
    
    ' Write results
    If newRows.Count > 0 Then
        Call WriteRowsToSheet(ws, newRows)
        Call UpdateTotalBuns(ws, "H1", "H2")
        
        MsgBox "? PAYPAL: Added " & newRows.Count & " new transactions." & vbCrLf & _
               "Skipped " & skippedCount & " duplicates.", vbInformation, "PayPal Import"
    Else
        MsgBox "No new PayPal payments found." & vbCrLf & _
               "Skipped " & skippedCount & " existing transactions.", vbInformation, "PayPal Import"
    End If
    
    Set outlookApp = Nothing
End Sub

' ========= ZELLE IMPORT FUNCTION =========
Sub ImportZelleEmails()
    Dim ws As Worksheet
    Dim outlookApp As Object
    Dim outlookNS As Object
    Dim inbox As Object
    Dim items As Object
    Dim mail As Object
    
    Dim newRows As Collection
    Dim existingKeys As Object
    Dim lastRow As Long
    Dim i As Long
    Dim skippedCount As Long
    
    ' Get or create Zelle sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Zelle")
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Zelle"
        Call SetupSheetHeaders(ws)
    End If
    
    ' Load existing data
    Set existingKeys = CreateObject("Scripting.Dictionary")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If lastRow > 1 Then
        For i = 2 To lastRow
            Dim key As String
            key = ws.Cells(i, 1).Value & "|" & ws.Cells(i, 4).Value & "|" & ws.Cells(i, 3).Value
            existingKeys(key) = True
        Next i
    End If
    
    ' Connect to Outlook
    On Error Resume Next
    Set outlookApp = GetObject(, "Outlook.Application")
    If outlookApp Is Nothing Then
        Set outlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
    
    Set outlookNS = outlookApp.GetNamespace("MAPI")
    Set inbox = outlookNS.GetDefaultFolder(6)
    Set items = inbox.items
    items.Sort "[ReceivedTime]", True
    
    Set newRows = New Collection
    skippedCount = 0
    
    ' Process emails
    For Each mail In items
        If InStr(1, mail.Subject, "Zelle", vbTextCompare) > 0 Or _
           InStr(1, mail.Body, "Zelle", vbTextCompare) > 0 Then
            
            If (InStr(1, mail.Subject, "sent you", vbTextCompare) > 0 Or _
                InStr(1, mail.Subject, "you received", vbTextCompare) > 0 Or _
                InStr(1, mail.Subject, "received $", vbTextCompare) > 0 Or _
                InStr(1, mail.Subject, "is now available", vbTextCompare) > 0) And _
               InStr(1, mail.Subject, "You sent", vbTextCompare) = 0 And _
               InStr(1, mail.Subject, "You paid", vbTextCompare) = 0 Then
                
                Call ProcessZelleEmail(mail, ws, newRows, existingKeys, skippedCount)
            End If
        End If
        
        If DateDiff("d", mail.ReceivedTime, Date) > 100 Then Exit For
    Next mail
    
    ' Write results
    If newRows.Count > 0 Then
        Call WriteRowsToSheet(ws, newRows)
        Call UpdateTotalBuns(ws, "H1", "H2")
        
        MsgBox "? ZELLE: Added " & newRows.Count & " new transactions." & vbCrLf & _
               "Skipped " & skippedCount & " duplicates.", vbInformation, "Zelle Import"
    Else
        MsgBox "No new Zelle payments found." & vbCrLf & _
               "Skipped " & skippedCount & " existing transactions.", vbInformation, "Zelle Import"
    End If
    
    Set outlookApp = Nothing
End Sub

' ========= EMAIL PROCESSING FUNCTIONS =========
Sub ProcessVenmoEmail(mail As Object, ws As Worksheet, newRows As Collection, _
                      existingKeys As Object, skippedCount As Long)
    Dim dateStr As String
    Dim amount As Double
    Dim person As String
    Dim note As String
    Dim bunCount As Long
    Dim uniqueKey As String
    Dim rowData As Variant
    
    ' Extract data
    dateStr = Format(mail.ReceivedTime, "m/d/yyyy")
    amount = ExtractAmount(mail.Body, mail.Subject)
    person = ExtractVenmoPerson(mail.Subject)
    note = ExtractVenmoNote(mail.Body)
    
    ' Create unique key
    uniqueKey = dateStr & "|" & person & "|" & amount
    
    ' Check for duplicate
    If existingKeys.Exists(uniqueKey) Then
        skippedCount = skippedCount + 1
        Exit Sub
    End If
    
    ' Calculate buns
    bunCount = CalculateBuns(amount)
    
    ' Add to collection
    rowData = Array(dateStr, "Payment Received", amount, person, note, "Completed", bunCount)
    newRows.Add rowData
    
    ' Mark as seen
    existingKeys(uniqueKey) = True
End Sub

Sub ProcessPayPalEmail(mail As Object, ws As Worksheet, newRows As Collection, _
                       existingKeys As Object, skippedCount As Long)
    Dim dateStr As String
    Dim amount As Double
    Dim person As String
    Dim note As String
    Dim bunCount As Long
    Dim uniqueKey As String
    Dim rowData As Variant
    
    dateStr = Format(mail.ReceivedTime, "m/d/yyyy")
    amount = ExtractAmount(mail.Body, mail.Subject)
    person = ExtractPayPalPerson(mail.Subject, mail.Body)
    note = ExtractPayPalNote(mail.Body)
    
    uniqueKey = dateStr & "|" & person & "|" & amount
    
    If existingKeys.Exists(uniqueKey) Then
        skippedCount = skippedCount + 1
        Exit Sub
    End If
    
    bunCount = CalculateBuns(amount)
    rowData = Array(dateStr, "Payment Received", amount, person, note, "Completed", bunCount)
    newRows.Add rowData
    existingKeys(uniqueKey) = True
End Sub

Sub ProcessZelleEmail(mail As Object, ws As Worksheet, newRows As Collection, _
                      existingKeys As Object, skippedCount As Long)
    Dim dateStr As String
    Dim amount As Double
    Dim person As String
    Dim note As String
    Dim bunCount As Long
    Dim uniqueKey As String
    Dim rowData As Variant
    
    dateStr = Format(mail.ReceivedTime, "m/d/yyyy")
    amount = ExtractAmount(mail.Body, mail.Subject)
    person = ExtractZellePerson(mail.Subject, mail.Body)
    note = ExtractZelleNote(mail.Body, mail.Subject)
    
    uniqueKey = dateStr & "|" & person & "|" & amount
    
    If existingKeys.Exists(uniqueKey) Then
        skippedCount = skippedCount + 1
        Exit Sub
    End If
    
    bunCount = CalculateBuns(amount)
    rowData = Array(dateStr, "Zelle Received", amount, person, note, "Completed", bunCount)
    newRows.Add rowData
    existingKeys(uniqueKey) = True
End Sub

' ========= EXTRACTION HELPER FUNCTIONS =========
Function ExtractAmount(bodyText As String, subjectText As String) As Double
    Dim regEx As Object
    Dim matches As Object
    Dim amountStr As String
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = "\$(\d{1,3}(?:,\d{3})*(?:\.\d{2})?)"
    regEx.Global = False
    
    ' Try subject first
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        amountStr = matches(0).SubMatches(0)
        ExtractAmount = CDbl(Replace(amountStr, ",", ""))
        Exit Function
    End If
    
    ' Try body
    If regEx.Test(bodyText) Then
        Set matches = regEx.Execute(bodyText)
        amountStr = matches(0).SubMatches(0)
        ExtractAmount = CDbl(Replace(amountStr, ",", ""))
        Exit Function
    End If
    
    ExtractAmount = 0
End Function

Function ExtractVenmoPerson(subjectText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    
    ' Pattern: "NAME paid you"
    regEx.Pattern = "^(.+?)\s+paid you"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractVenmoPerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ' Pattern: "NAME paid $ to your Venmo account"
    regEx.Pattern = "^(.+?)\s+paid\s+\$[\d,]+(?:\.\d{2})?\s+to your Venmo account"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractVenmoPerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractVenmoPerson = "Unknown"
End Function

Function ExtractVenmoNote(bodyText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = "class=""transaction-note[^>]*>([^<]+)<"
    
    If regEx.Test(bodyText) Then
        Set matches = regEx.Execute(bodyText)
        ExtractVenmoNote = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractVenmoNote = ""
End Function

Function ExtractPayPalPerson(subjectText As String, bodyText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    
    ' Try subject patterns
    regEx.Pattern = "(.+?)\s+sent you"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractPayPalPerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    regEx.Pattern = "from\s+(.+?)$"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractPayPalPerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractPayPalPerson = "Unknown"
End Function

Function ExtractPayPalNote(bodyText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    regEx.Pattern = "(?:Note|Message):\s*(?:<[^>]*>)?([^<\r\n]+)"
    
    If regEx.Test(bodyText) Then
        Set matches = regEx.Execute(bodyText)
        ExtractPayPalNote = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractPayPalNote = ""
End Function

Function ExtractZellePerson(subjectText As String, bodyText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    
    ' Pattern: "NAME sent you $X with Zelle"
    regEx.Pattern = "^(.+?)\s+sent you\s+\$[\d,]+(?:\.\d{2})?\s+with\s+Zelle"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractZellePerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ' Pattern: "received $ from NAME"
    regEx.Pattern = "received\s+\$[\d,]+(?:\.\d{2})?\s+from\s+(.+)"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractZellePerson = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractZellePerson = "Unknown"
End Function

Function ExtractZelleNote(bodyText As String, subjectText As String) As String
    Dim regEx As Object
    Dim matches As Object
    
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.IgnoreCase = True
    
    ' Look for Memo/Message/Note in body
    regEx.Pattern = "(?:Memo|Message|Note)\s*:\s*([^\r\n]+)"
    If regEx.Test(bodyText) Then
        Set matches = regEx.Execute(bodyText)
        ExtractZelleNote = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ' Look for "for X" in subject
    regEx.Pattern = "\s+for\s+(.+?)$"
    If regEx.Test(subjectText) Then
        Set matches = regEx.Execute(subjectText)
        ExtractZelleNote = Trim(matches(0).SubMatches(0))
        Exit Function
    End If
    
    ExtractZelleNote = ""
End Function

' ========= BUN CALCULATION (MATCHES YOUR ORIGINAL) =========
Function CalculateBuns(amount As Double) As Long
    Dim bunCount As Long
    Dim remaining As Double
    
    bunCount = 0
    
    ' Exact matches
    Select Case amount
        Case 2: bunCount = 1
        Case 4: bunCount = 2
        Case 5: bunCount = 3
        Case 7: bunCount = 4
        Case 10: bunCount = 5
        Case 12: bunCount = 6
        Case 15: bunCount = 8
        Case 20: bunCount = 10
        Case Else
            ' Greedy algorithm
            remaining = amount
            
            Do While remaining >= 10
                bunCount = bunCount + 5
                remaining = remaining - 10
            Loop
            
            Do While remaining >= 5
                bunCount = bunCount + 3
                remaining = remaining - 5
            Loop
            
            Do While remaining >= 2
                bunCount = bunCount + 1
                remaining = remaining - 2
            Loop
            
            If remaining >= 1 Then
                bunCount = bunCount + Int(remaining / 2)
            End If
    End Select
    
    CalculateBuns = bunCount
End Function

' ========= SHEET HELPER FUNCTIONS =========
Sub SetupSheetHeaders(ws As Worksheet)
    With ws
        .Range("A1:G1").Value = Array("Date", "Type (payment or charge)", "Amount", _
                                      "Person", "Note", "Status", "Cinnamon Buns to Make")
        .Range("A1:G1").Font.Bold = True
        .Range("A1:G1").Interior.ColorIndex = 15
        .Columns("A:G").AutoFit
    End With
End Sub

Sub WriteRowsToSheet(ws As Worksheet, newRows As Collection)
    Dim lastRow As Long
    Dim i As Long
    Dim rowData As Variant
    
    ' Sort by date (newest first) - simplified insert at top
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = newRows.Count To 1 Step -1
        ws.Rows(2).Insert Shift:=xlDown
        rowData = newRows(i)
        ws.Range("A2:G2").Value = rowData
    Next i
End Sub

Sub UpdateTotalBuns(ws As Worksheet, headerCell As String, totalCell As String)
    Dim lastRow As Long
    Dim totalBuns As Long
    Dim i As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    totalBuns = 0
    
    For i = 2 To lastRow
        If IsNumeric(ws.Range("G" & i).Value) Then
            totalBuns = totalBuns + ws.Range("G" & i).Value
        End If
    Next i
    
    ws.Range(headerCell).Value = "?? Total Buns"
    ws.Range(totalCell).Value = totalBuns
    ws.Range(headerCell & ":" & totalCell).Font.Bold = True
End Sub


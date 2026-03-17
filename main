Sub Final_MultiField_Merge()
    Dim masterDoc As Document
    Dim recordCount As Long
    Dim i As Integer
    Dim fileName As String
    Dim folderPath As String
    Dim emailAddr As String
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim strBody As String
    
    ' ==========================================================
    ' 1. CONFIGURATION - ADD YOUR EXTRA FIELDS HERE
    ' ==========================================================
    Dim fNameField As String: fNameField = "Full_Name" 
    Dim emailField As String: emailField = "Email"
    Dim deptField As String: deptField = "Department"  ' NEW FIELD EXAMPLE  
    Dim monthField As String: monthField = "Month"     ' NEW FIELD EXAMPLE
    
    Dim sharedMailbox As String: sharedMailbox = "hr-department@company.com" 'EXAMPLE MAIL
    Dim isTestMode As Boolean: isTestMode = False ' Set to True to see emails before they send
    ' ==========================================================

    Set masterDoc = ActiveDocument
    
    ' 2. Check if the fields exist (Prevent Error 5941)
    On Error Resume Next
    Dim test1 As String: test1 = masterDoc.MailMerge.DataSource.DataFields(fNameField).Value
    Dim test2 As String: test2 = masterDoc.MailMerge.DataSource.DataFields(emailField).Value
    Dim test3 As String: test3 = masterDoc.MailMerge.DataSource.DataFields(deptField).Value
    Dim test4 As String: test4 = masterDoc.MailMerge.DataSource.DataFields(monthField).Value
    If Err.Number <> 0 Then
        MsgBox "Error: One of your column headers is wrong. Check Excel for: " & fNameField & ", " & emailField & ", or " & deptField, vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 3. Setup Folders
    folderPath = Options.DefaultFilePath(wdDocumentsPath) & "\Merge_Temp\"
    If Dir(folderPath, vbDirectory) & "" = "" Then MkDir folderPath
    
    Set outlookApp = CreateObject("Outlook.Application")
    masterDoc.MailMerge.DataSource.ActiveRecord = wdLastRecord
    recordCount = masterDoc.MailMerge.DataSource.ActiveRecord
    masterDoc.MailMerge.DataSource.ActiveRecord = wdFirstRecord
    
    ' 4. The Loop
    For i = 1 To recordCount
        masterDoc.MailMerge.DataSource.ActiveRecord = i
        masterDoc.MailMerge.ViewMailMergeFieldCodes = False
        masterDoc.Fields.Update
        
        ' Pull data for the current employee
        fileName = folderPath & masterDoc.MailMerge.DataSource.DataFields(fNameField).Value & ".pdf"
        emailAddr = masterDoc.MailMerge.DataSource.DataFields(emailField).Value
        
        ' --- CUSTOM PERSONALIZE USING THE EXTRA FIELDS ---
        strBody = "Hi " & masterDoc.MailMerge.DataSource.DataFields(fNameField).Value & "," & vbCrLf & vbCrLf & _
                  "Please find the " & masterDoc.MailMerge.DataSource.DataFields(monthField).Value & _
                  " report for the " & masterDoc.MailMerge.DataSource.DataFields(deptField).Value & " department attached." & vbCrLf & vbCrLf & _
                  "Best regards," & vbCrLf & _
                  "HR Management"
        
        ' Export to PDF
        masterDoc.ExportAsFixedFormat OutputFileName:=fileName, _
            ExportFormat:=wdExportFormatPDF, Range:=wdExportOnedocument
            
        ' Create and Send/Display Email
        Set outlookMail = outlookApp.CreateItem(0)
        With outlookMail
            .SentOnBehalfOfName = sharedMailbox
            .To = emailAddr
            .Subject = masterDoc.MailMerge.DataSource.DataFields(monthField).Value & " Document - " & _
                       masterDoc.MailMerge.DataSource.DataFields(fNameField).Value
            .Body = strBody
            .Attachments.Add fileName
            
            If isTestMode Then .Display Else .Send
        End With
        

    Next i
    
    MsgBox "All done! " & recordCount & " personalized emails processed.", vbInformation
End Sub

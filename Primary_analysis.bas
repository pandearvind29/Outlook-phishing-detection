Sub AnalyzeDownloadedEmail(filePath As String)
    Dim olMail As Outlook.MailItem
    Dim headers As String
    Dim replyToAddress As String
    Dim senderAddress As String
    Dim isPhishingLinkFound As Boolean
    
    ' Open the email file (.msg)
    Set olMail = Application.CreateItemFromTemplate(filePath)
    
    If Not olMail Is Nothing Then
        ' Display subject for reference
        Debug.Print "Subject: " & olMail.Subject
        
        ' Get email headers
        headers = GetEmailHeaders(olMail)
        
        ' Check Reply-To address and Sender address
        replyToAddress = GetHeaderField(headers, "Reply-To")
        senderAddress = GetHeaderField(headers, "Sender")
        
        Debug.Print "Reply-To Address: " & replyToAddress
        Debug.Print "Sender Address: " & senderAddress
        
        ' Check SPF and DKIM headers
        Dim spfResult As String
        Dim dkimResult As String
        
        spfResult = GetHeaderField(headers, "Received-SPF")
        dkimResult = GetHeaderField(headers, "DKIM-Signature")
        
        Debug.Print "SPF Result: " & spfResult
        Debug.Print "DKIM Signature: " & dkimResult
        
        ' Check for suspicious keywords and URLs in email body
        isPhishingLinkFound = CheckEmailBodyForPhishingLinks(olMail.Body)
        
        If isPhishingLinkFound Then
            MsgBox "Potential phishing link found in the email body.", vbExclamation
        Else
            MsgBox "No phishing links detected in the email body.", vbInformation
        End If
        
        ' Check for suspicious attachment types
        CheckForSuspiciousAttachments olMail
    Else
        MsgBox "Error: Could not open the email file.", vbCritical
    End If
End Sub

' Function to retrieve the headers of an email
Function GetEmailHeaders(olMail As Outlook.MailItem) As String
    Dim headers As String
    On Error Resume Next
    headers = olMail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    On Error GoTo 0
    GetEmailHeaders = headers
End Function

' Function to get a specific header field value
Function GetHeaderField(headers As String, fieldName As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim fieldValue As String
    
    startPos = InStr(1, headers, fieldName & ":", vbTextCompare)
    If startPos > 0 Then
        startPos = startPos + Len(fieldName) + 1
        endPos = InStr(startPos, headers, vbCrLf)
        fieldValue = Trim(Mid(headers, startPos, endPos - startPos))
    End If
    
    GetHeaderField = fieldValue
End Function

' Function to check URLs in email body
Function CheckEmailBodyForPhishingLinks(bodyText As String) As Boolean
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = "(https?://[^\s""]+)"
    
    Set matches = regex.Execute(bodyText)
    
    If matches.Count > 0 Then
        Debug.Print "URLs found in the email body:"
        For Each match In matches
            Debug.Print match.Value
            ' Example: Check for common phishing keywords
            If InStr(match.Value, "phish") > 0 Or InStr(match.Value, "login") > 0 Then
                CheckEmailBodyForPhishingLinks = True
                Exit Function
            End If
        Next
    End If
    CheckEmailBodyForPhishingLinks = False
End Function

' Function to check for suspicious attachment types
Sub CheckForSuspiciousAttachments(olMail As Outlook.MailItem)
    Dim attachment As Outlook.Attachment
    Dim dangerousExtensions As Variant
    Dim ext As String
    Dim i As Integer
    Dim isSuspicious As Boolean
    
    dangerousExtensions = Array("exe", "js", "vbs", "scr", "bat", "cmd", "com")
    
    For Each attachment In olMail.Attachments
        ext = Right(attachment.FileName, Len(attachment.FileName) - InStrRev(attachment.FileName, "."))
        For i = LBound(dangerousExtensions) To UBound(dangerousExtensions)
            If LCase(ext) = dangerousExtensions(i) Then
                isSuspicious = True
                Debug.Print "Suspicious attachment found: " & attachment.FileName
            End If
        Next i
    Next attachment
    
    If isSuspicious Then
        MsgBox "Warning: Suspicious attachment(s) found!", vbExclamation
    Else
        MsgBox "No suspicious attachments detected.", vbInformation
    End If
End Sub

Attribute VB_Name = "Main"
Option Explicit
Option Compare Text

Public Function Get_Phishing_Emails()

    Application.ScreenUpdating = False
    
    Dim lastRow As Integer
    Dim rowsAboveHeaders As Integer
    
    Dim OutlookApp As Outlook.Application
    Dim OutlookNamespace As Namespace
    Dim Folder As MAPIFolder
    Dim OutlookMail As Object
    Dim i As Integer
    
    Set OutlookApp = New Outlook.Application
    Set OutlookNamespace = OutlookApp.GetNamespace("MAPI")
    Set Folder = OutlookNamespace.GetDefaultFolder(olFolderInbox)
    
    ' From date
    Dim fromDate As Date: fromDate = Sheet1.Range("eMail_FromDate").Value
    Sheet1.Range("eMail_FromDate").Value = Now

    ' Calculate how much to offset from headers
    lastRow = Application.Max(GetLastRow("A", Sheet1), GetLastRow("B", Sheet1), GetLastRow("C", Sheet1))
    rowsAboveHeaders = Sheet1.Range("eMail_Subject").Row - 1
    i = lastRow - rowsAboveHeaders
    
    For Each OutlookMail In Folder.Items
    
        ' Check if this is a new email
        If CDate(OutlookMail.ReceivedTime) > fromDate Then
        
            ' Check Header Data
            Dim headers As String
            headers = GetHeaders(OutlookMail)
            If InStr(headers, "X-PHISH-CRID") > 0 Then
                
                ' Check that this isn't a "Scam of the Week" email
                Dim emailSubject As String
                emailSubject = OutlookMail.Subject
                If Not (InStr(emailSubject, "Scam of the Week") > 0) Then
                
                    With Sheet1
                        
                        Dim headerLineBreaksRemoved As String
                        headerLineBreaksRemoved = Replace(headers, vbCrLf, "")
                        
                        ' Get Sender Name
                        Range("eMail_SenderName").Offset(i, 0).Value = OutlookMail.SenderName
                        
                        ' Get Email Subject
                        Range("eMail_Subject").Offset(i, 0).Value = OutlookMail.Subject
                        
                        ' Get Received Time
                        Range("eMail_ReceivedTime").Offset(i, 0).Value = OutlookMail.ReceivedTime
                        
                        ' Get Header Phishing ID
                        Range("eMail_PhishingID").Offset(i, 0).Value = GetMidText(headerLineBreaksRemoved, "X-PHISH-CRID:", "X-KNOWBE4:")
                        
                        i = i + 1
                        
                    End With
                
                End If
            
            End If
            
        End If
        
    Next OutlookMail

    Set Folder = Nothing
    Set OutlookNamespace = Nothing
    Set OutlookApp = Nothing

    Application.ScreenUpdating = True

End Function

' RETURNS the email's headers
Private Function GetHeaders(email As Object) As String

    Dim headers As String
    Dim propertyAccessor As Object
    Set propertyAccessor = email.propertyAccessor
    
    On Error Resume Next
    
    headers = propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E")
    If Err.Number <> 0 Then
        headers = "Error retrieving headers"
    End If
    
    On Error GoTo 0
    
    GetHeaders = headers
    
End Function

' RETURNS last row in a given column
Private Function GetLastRow(lookupCol As String, lookupSheet As Worksheet) As Long
    
    GetLastRow = lookupSheet.Range(lookupCol & Rows.count).End(xlUp).Row

End Function

' RETURNS text between the first occurence startStr and the first occurence endStr
Private Function GetMidText(sourceStr, startStr, endStr As String) As String

    Dim startPosition As Integer:
    Dim endPosition As Integer
    
    'Get positions of start/end strings
    'Returns zero if not found
    startPosition = InStr(1, sourceStr, startStr, 0)
    endPosition = InStr(1, sourceStr, endStr, 0)
    
    'Check if start/end strings are found in source string
    If startPosition = 0 Or endPosition = 0 Then
        GetMidText = ""
    Else
        GetMidText = Trim(Mid(sourceStr, startPosition + Len(startStr), endPosition - startPosition - Len(startStr)))
    End If

End Function


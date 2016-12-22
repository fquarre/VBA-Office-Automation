Sub removeAttachments()
    'Declaration
    Dim myItems, myItem, myAttachments, myAttachment As Object
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim attchmentList As String

    'work on selected items
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    'for all items do...
    For Each myItem In myOlSel
        Set myAttachments = myItem.Attachments
        
        'if there are some...
        If myAttachments.Count > 0 Then
            attchmentList = "----------------------------------------------" & vbCrLf & "Removed Attachments:" & vbCrLf 'add remark to message text
            While myAttachments.Count > 0
				attchmentList = attchmentList & "File: " & myAttachments(1).DisplayName & vbCrLf  'add name and destination to message text
                myAttachments(1).Delete
            Wend
            attchmentList = attchmentList & "----------------------------------------------" & vbCrLf

            myItem.Body = attchmentList & myItem.Body
            myItem.Save 'save item without attachments
        End If
    Next

    'free variables
    Set myItems = Nothing
    Set myItem = Nothing
    Set myAttachments = Nothing
    Set myAttachment = Nothing
    Set myOlApp = Nothing
    Set myOlExp = Nothing
    Set myOlSel = Nothing
End Sub

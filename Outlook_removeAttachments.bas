Sub listAttachments()
    'Declaration
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim myItem, myAttachments As Object
    Dim attchmentList As String


    'work on selected items
    'outlook elements
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    'for all items do...
    For Each myItem In myOlSel
        Set myAttachments = myItem.Attachments
        
        'if there are some...
        
        If myAttachments.Count > 0 Then
            currentCount = 1
            attchmentList = ""
            While currentCount <= myAttachments.Count
                attchmentList = attchmentList & "<b>" & currentCount & ". " & myAttachments(currentCount).FileName & ":</b>  <br>"
                currentCount = currentCount + 1
            Wend

            myItem.HTMLBody = attchmentList & myItem.HTMLBody
        
        End If
    Next

    'free variables
    Set myItem = Nothing
    Set myAttachments = Nothing
    Set myOlApp = Nothing
    Set myOlExp = Nothing
    Set myOlSel = Nothing
End Sub
                                
Sub saveAttachments()
    'Declaration
    Dim myItems, myItem, myAttachments, myAttachment As Object
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim ShellApp As Object
    Dim attchmentList As String

    'work on selected items
    '1. folder picker
    Set oShell = CreateObject("WScript.Shell")
    openAt = oShell.ExpandEnvironmentStrings("%HOMEDRIVE%") & oShell.ExpandEnvironmentStrings("%HOMEPATH%") & "\Documents\Work" 'where to start the folderPicker
    Set ShellApp = CreateObject("Shell.Application").BrowseForFolder(0, "Please choose a folder", 0, openAt)
    If ShellApp Is Nothing Then
        GoTo ExitSub
    End If
    BrowseForFolder = ShellApp.self.Path
    Set ShellApp = Nothing

    '2. outlook elements
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    'for all items do...
    For Each myItem In myOlSel
        Set myAttachments = myItem.Attachments
        
        'if there are some...
        If myAttachments.Count > 0 Then
            attchmentList = "<p>----------------------------------------------<br>Saved Attachments:<br>" 'add remark to message text
            While myAttachments.Count > 0
                strFile = myAttachments(1).FileName
                strFile = BrowseForFolder & "\" & strFile
                myAttachments(1).SaveAsFile strFile
                attchmentList = attchmentList & "<a href='file://" & strFile & "'>" & strFile & "</a><br>" 'add name and destination to message text // if not HTML strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
                myAttachments(1).Delete
            Wend
            attchmentList = attchmentList & "----------------------------------------------</p>"

            myItem.HTMLBody = attchmentList & myItem.HTMLBody

            myItem.Save 'save item without attachments
        End If
    Next

'branch out in case user cancelled the window / folder select
ExitSub:
    'free variables
    Set myItems = Nothing
    Set myItem = Nothing
    Set myAttachments = Nothing
    Set myAttachment = Nothing
    Set myOlApp = Nothing
    Set myOlExp = Nothing
    Set myOlSel = Nothing
End Sub

Sub removeAttachments()
    'Declaration
    Dim myItems, myItem, myAttachments, myAttachment As Object
    Dim myOlApp As New Outlook.Application
    Dim myOlExp As Outlook.Explorer
    Dim myOlSel As Outlook.Selection
    Dim attchmentList As String

    'work on selected items
    '1. folder picker
    'NA
    
    '2. outlook elements
    Set myOlExp = myOlApp.ActiveExplorer
    Set myOlSel = myOlExp.Selection

    'for all items do...
    For Each myItem In myOlSel
        Set myAttachments = myItem.Attachments
        
        'if there are some...
        If myAttachments.Count > 0 Then
            attchmentList = "<p>----------------------------------------------<br>Removed Attachments:<br>" 'add remark to message text
            While myAttachments.Count > 0
                strFile = myAttachments(1).FileName
                attchmentList = attchmentList & "<a href='file://" & strFile & "'>" & strFile & "</a><br>"
                'add name and destination to message text // if not HTML strDeletedFiles = strDeletedFiles & vbCrLf & "<file://" & strFile & ">"
                myAttachments(1).Delete
            Wend
            attchmentList = attchmentList & "----------------------------------------------</p>"

            myItem.HTMLBody = attchmentList & myItem.HTMLBody

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

Sub autoForward(Item As Outlook.MailItem)
    Set myForward = Item.Forward
    myForward.Recipients.Add "[EMAIL ADDRESS]"
    myForward.Send
End Sub

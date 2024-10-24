Private WithEvents Items As Outlook.Items

Private Sub Application_Startup()
    Dim olApp As Outlook.Application
    Set olApp = Application
    Set Items = olApp.ActiveExplorer.CurrentFolder.Items
End Sub

Private Sub Items_ItemChange(ByVal Item As Object)
    If TypeName(Item) = "MailItem" Then
        Call DisableHyperlinks(Item)
    End If
End Sub

Private Sub DisableHyperlinks(Item As MailItem)
    Dim RegEx As Object
    Set RegEx = CreateObject("VBScript.RegExp")
    Dim Matches As Object
    Dim Match As Object
    Dim HTMLBody As String
    Dim newBody As String

    With RegEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "<a[^>]+href=""([^""]+)""[^>]*>(.*?)<\/a>"
    End With

    HTMLBody = Item.HTMLBody
    Set Matches = RegEx.Execute(HTMLBody)

    newBody = HTMLBody
    For Each Match In Matches
        Dim url As String, text As String
        url = Match.SubMatches(0)
        text = Match.SubMatches(1)
        newBody = Replace(newBody, Match.Value, text)
    Next

    If newBody <> HTMLBody Then
        Item.HTMLBody = newBody
        Item.Save
    End If

    Set RegEx = Nothing
    Set Matches = Nothing
End Sub
Imports Microsoft.Office.Interop.Outlook

Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup

        ASL_Tools.app = Me.Application

        'Check_OfflineCategory()
    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Application_ItemSend(Item As Object, ByRef Cancel As Boolean) Handles Application.ItemSend
        Dim olmi As Outlook.MailItem = Item

        'get the senders domain name.
        'if the domain name is anthony-seaman then enable the store functionality.
        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(olmi.SenderEmailAddress)
        If senderDomain = "asltd.com" Then
            Dim frm As form_EmailSend = New form_EmailSend
            frm.button_record.Enabled = False
            frm.Item = Item
            frm.ShowDialog()

            frm.Close()
        End If
    End Sub

    Public Sub Application_ItemRecord()
        If Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        'get the first selected item.
        'if the store the item resides in is in the domain name asltd.com continue

        Dim emList As List(Of Outlook.MailItem) = New List(Of Outlook.MailItem)
        For Each it As Outlook.MailItem In Application.ActiveExplorer.Selection
            emList.Add(it)
        Next

        Dim lastFld As String = emList.Item(0).Parent.fullfolderpath
        lastFld = lastFld.Substring(lastFld.Length - 5, 5)

        If Not (lastFld.ToUpper = "INBOX") Then Exit Sub


        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(emList.Item(0).Parent.store.displayname.ToString)

        If senderDomain = "asltd.com" Then
            Dim frm As form_EmailRecord = New form_EmailRecord
            frm.button_record.Enabled = False
            frm.emList = emList
            frm.ShowDialog()

            frm.Close()
        End If

    End Sub

    Private Sub Explorer_SelectFolder()

    End Sub

    Private Sub Application_Startup() Handles Application.Startup

    End Sub
End Class

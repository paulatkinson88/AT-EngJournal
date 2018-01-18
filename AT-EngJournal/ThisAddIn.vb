Public Class ThisAddIn

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        ASL_Tools.app = Me.Application
        Check_OfflineCategory()
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
End Class

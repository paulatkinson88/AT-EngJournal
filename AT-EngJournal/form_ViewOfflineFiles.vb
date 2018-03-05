Imports System.Windows.Forms

Public Class form_ViewOfflineFiles
    Private Sub form_ViewOfflineFiles_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        label_status.Text = "Getting Files"
        ASL_Tools.GetOfflineMessages()

        label_status.Text = "Display Files"

        For Each itm As Outlook.MailItem In ASL_Tools.msList
            Dim st(2) As String
            st(0) = itm.EntryID.ToString
            st(1) = itm.Subject
            st(2) = itm.Parent.folderpath
            Dim lstItm As ListViewItem = New ListViewItem(st)
            ListView1.Items.Add(lstItm)
        Next

        label_status.Text = "Idle"
        label_count.Text = ASL_Tools.msList.Count.ToString
    End Sub

    Private Sub ListView1_DoubleClick(sender As Object, e As EventArgs) Handles ListView1.DoubleClick
        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Dim lItm As ListViewItem = ListView1.SelectedItems(0)
        Dim EntryIdItem = lItm.SubItems(0).Text
        Dim EntryIDStore = ASL_Tools.aslStore.StoreID

        Dim itm = Globals.ThisAddIn.Application.Session.GetItemFromID(EntryIdItem, EntryIDStore)
        itm.display
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If ListView1.SelectedItems.Count = 0 Then Exit Sub

        Dim lItm As ListViewItem = ListView1.SelectedItems(0)
        Dim EntryIdItem = lItm.SubItems(0).Text
        Dim EntryIDStore = ASL_Tools.aslStore.StoreID

        Dim itm = Globals.ThisAddIn.Application.Session.GetItemFromID(EntryIdItem, EntryIDStore)
        itm.display
    End Sub
End Class
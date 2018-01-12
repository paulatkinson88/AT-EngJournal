Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon

Public Class ASL_Ribbon

    Private Sub ASL_Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        checkForNetwork()
    End Sub

    Private Sub button_checkForNetwork_Click(sender As Object, e As RibbonControlEventArgs)
        checkForNetwork()
    End Sub

    Public Sub checkForNetwork()
        If ASL_Tools.Check_For_Network Then
            label_connection.Label = "Ready: True"
        Else
            label_connection.Label = "Ready: False"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        checkForNetwork()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        'MsgBox(asl.app.Inspectors.Count.ToString)

        Dim mIns As Outlook.Inspector = ASL_Tools.app.ActiveInspector
        If Not (IsNothing(mIns)) Then
            Debug.Print(mIns.GetType.ToString)
        End If

        Dim st As String = InputBox("Find Project:", "Find")

        Dim fld As Outlook.Folder = ASL_Tools.Get_ProjectFolder_From_ASL_Store_Inbox(st)
        If IsNothing(fld) Then
            MsgBox("No Project Found", vbCritical, "Error")
        Else
            MsgBox("Folder NameOf: " & fld.Name & "  Message Count: " & fld.Items.Count.ToString)
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As RibbonControlEventArgs) Handles Button3.Click
        ASL_Tools.Get_ASL_Store()

    End Sub

    Public Sub test()
        Dim app As Outlook.Application = Globals.ThisAddIn.Application
        Dim accs As Outlook.Accounts = app.Session.Accounts
        Dim stor As Outlook.Stores = app.Session.Stores


        'Dim inBox As Outlook.MAPIFolder = ASL_Tools.app.ActiveExplorer.Session.Folder
        Dim aex As Outlook.Explorer = ASL_Tools.app.ActiveExplorer

        Dim aexs As Outlook.Explorers = ASL_Tools.app.Explorers

        For Each aexd In aexs
            Dim obj = aexs.Parent
            MsgBox(obj.GetType.ToString)
            Debug.Print("Caption: " & aexd.Caption)

        Next

        For Each ac As Outlook.Account In accs
            Debug.Print("Account Display Name: " & ac.DisplayName)

        Next
        For Each st As Outlook.Store In stor
            Debug.Print("Store Display Name: " & st.DisplayName & "-" & st.FilePath)
            For Each fl As Outlook.Category In st.Categories
                Debug.Print("    Cat: " & fl.Name)
            Next
        Next
        'Debug.Print("Active Explorer caption:" & aex.Caption)
        'Debug.Print(aex.CurrentFolder.FolderPath)


        'Dim subFolder As Outlook.MAPIFolder



        'If StrComp(folderNames(0), "olFolderInbox", vbTextCompare) = 0 Then
        'Dim Get_Outlook_Folder = outNs.GetDefaultFolder(olFolderInbox)

        'Else
        'On Error Resume Next   'trap error if first folder name doesn't exist
        'Dim Get_Outlook_Folder = outNs.Folders(folderNames(0))
        'On Error GoTo 0
        'End If

        'i = 0
        'While Not Get_Outlook_Folder Is Nothing And i < UBound(folderNames)
        'i = i + 1
        'If folderNames(i) = ".." Then
        '   Set Get_Outlook_Folder = Get_Outlook_Folder.Parent
        'Else
        '    Set subFolder = Nothing
        ' On Error Resume Next   'trap error if subfolder doesn't exist
        '    Set subFolder = Get_Outlook_Folder.Folders(folderNames(i))
        'On Error GoTo 0
        '   Set Get_Outlook_Folder = subFolder
        'End If
        'Wend
    End Sub
End Class

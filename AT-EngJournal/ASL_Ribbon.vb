Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon
Imports System.Reflection

Public Class ASL_Ribbon

    Private Sub ASL_Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        checkForNetwork()

        Dim bob = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version

        label_version.Label = bob.ToString
        'ApplicationDeployment.CurrentDeployment.CurrentVersion
        Fill_Label_OffLineFileCount()
    End Sub

    Private Sub Fill_Label_OffLineFileCount()
        ASL_Tools.Get_OffLineFileCount()
        label_offlinefilecount.Label = ASL_Tools.offlineFileCount
    End Sub

    Private Sub button_checkForNetwork_Click(sender As Object, e As RibbonControlEventArgs)
        checkForNetwork()
    End Sub

    Public Sub checkForNetwork()
        If ASL_Tools.Check_For_Network Then
            label_connection.Label = "True"
        Else
            label_connection.Label = "False"
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click
        checkForNetwork()
    End Sub

    Private Sub Button2_Click(sender As Object, e As RibbonControlEventArgs) Handles Button2.Click
        Dim mIns As Outlook.Inspector = ASL_Tools.app.ActiveInspector
        If Not (IsNothing(mIns)) Then
            Debug.Print(mIns.GetType.ToString)
        End If

        Dim st As String = InputBox("Find Project:", "Find")


        'begin the structure for the send function

        'copy the message to the folder
        '   find the projectfolder in the aslstore

        Dim fld As Outlook.Folder = ASL_Tools.Get_ProjectFolder_In_ASLStoreInbox(st)

        'if the project is not found then create it.
        If IsNothing(fld) Then
            MsgBox("No Project Found", vbCritical, "Error")
            'create the project folder
            fld = ASL_Tools.Create_ProjectFolder_In_ASLStoreInbox(st)
            If IsNothing(fld) Then
                MsgBox("Error getting or creating project", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'with the project folder
        'get the sent folder
        Dim fldSent As Outlook.Folder = ASL_Tools.Get_ProjectFolderSent(fld)

        If IsNothing(fldSent) Then
            fldSent = ASL_Tools.Create_ProjectFolderSent(fld)
            If IsNothing(fldSent) Then
                MsgBox("Error creating sent folder", vbCritical, "Error")
                Exit Sub
            End If
        End If

        'we now have the project folder and the sent folder to save the sent message to.


        MsgBox("Folder NameOf: " & fld.Name & "  Message Count: " & fld.Items.Count.ToString)

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

    Private Sub button_OfflineFilesCount_Click(sender As Object, e As RibbonControlEventArgs) Handles button_OfflineFilesCount.Click
        Fill_Label_OffLineFileCount()

    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles Button4.Click
        'for each file that is offline. push them to the server.
        For Each mo As Outlook.MailItem In ASL_Tools.msList
            MsgBox(mo.Parent.ToString)
        Next

    End Sub
End Class

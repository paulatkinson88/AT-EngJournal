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

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles button_pushOfflineFilestoServer.Click
        'for each file that is offline. push them to the server.
        For Each mo As Outlook.MailItem In ASL_Tools.msList
            Dim msgProperties As ASLmessageProperties = ASL_Tools.Get_StampProperty(mo)

            If msgProperties.messagetype = "re" Or msgProperties.messagetype = "se" Then
                If ASL_Tools.networkReady = True Then
                    'copy to network
                    Dim di As System.IO.DirectoryInfo = ASL_Tools.Check_For_ProjectDirectoryEngJournal(msgProperties.proj)
                    If Not (IsNothing(di)) Then
                        'use the messageKeyValue as the message name when saving to the 
                        'network
                        mo.Categories = ""
                        mo.SaveAs(di.FullName & "\(" & msgProperties.proj & ")(" & msgProperties.timestamp & ")(" & msgProperties.messagetype & ")" & ".msg")
                    End If
                End If
            Else
                MsgBox("Message is damaged and cannot be stored on Network." & vbLf & mo.Subject, vbCritical, "Error")
            End If
        Next

        Fill_Label_OffLineFileCount()
    End Sub

    Private Sub button_recordEmail_Click(sender As Object, e As RibbonControlEventArgs) Handles button_recordEmail.Click
        Globals.ThisAddIn.Application_ItemRecord()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As RibbonControlEventArgs) Handles button_viewOfflineFiles.Click
        Try
            ASL_Tools.offLineFileForm.Show()
        Catch ex As Exception
            ASL_Tools.offLineFileForm = New form_ViewOfflineFiles
            ASL_Tools.offLineFileForm.Show()
        End Try

        ASL_Tools.offLineFileForm.Focus()

    End Sub

    Private Sub button_MoveEmail_Click(sender As Object, e As RibbonControlEventArgs) Handles button_MoveEmail.Click
        'allow the user to move emails from one project to another.
        'check to see if they are in a project directory or in a sent folder under the project directory
        'you cannot move an email if you are not connected to the network
        If Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        If Not (ASL_Tools.networkReady) Then
            MsgBox("You need to be connected to the network to move emails to another project.", vbCritical, "Error")
            Exit Sub
        End If

        Dim frm As form_emailMove = New form_emailMove
        Dim emList As List(Of Outlook.MailItem) = New List(Of Outlook.MailItem)
        For Each obj As Outlook.MailItem In Globals.ThisAddIn.Application.ActiveExplorer.Selection
            emList.Add(obj)

        Next
        Globals.ThisAddIn.Application.ActiveExplorer.ClearSelection()

        For Each em As Outlook.MailItem In emList
            Dim fld As String = em.Parent.FullFolderPath
            Dim fldObj() As String = fld.Split("\")

            Dim msgProp As ASLmessageProperties = ASL_Tools.Get_StampProperty(em)
            If msgProp.messagetype = "se" Then
                'sent item will be in the sent directory under the project
                'check to see the message inbox word is in the third point of the back of the folder array
                Dim fldInbox As String = fldObj(fldObj.Length - 3)
                If fldInbox = "Inbox" Then
                    'the message is in the correct spot.
                    'get the users input for the project to be in.
                    frm.em2 = em
                    frm.msgProp = msgProp
                    frm.ShowDialog()
                    frm.Close()
                End If
            ElseIf msgProp.messagetype = "re" Then
                'received item will be in the project directory
                'check to see the message inbox word is in the second point of the back of the folder array
                Dim fldinbox As String = fldObj(fldObj.Length - 2)
                If fldinbox = "Inbox" Then
                    'the message is in the correct spot.
                    'get the users input for the project to be in.
                    frm.em2 = em
                    frm.msgProp = msgProp
                    frm.ShowDialog()
                    frm.Close()
                End If
            Else
                MsgBox("Message is damaged and cannot be moved." & vbLf & em.Subject, vbCritical, "Error")
            End If

        Next

    End Sub

    Private Sub button_getUserProperties_Click(sender As Object, e As RibbonControlEventArgs) Handles button_getUserProperties.Click
        If Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        Debug.Print(Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count.ToString)
        'get the first selected item.
        'if the store the item resides in is in the domain name asltd.com continue

        Dim em As Outlook.MailItem = Globals.ThisAddIn.Application.ActiveExplorer.Selection.Item(1)

        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(em.Parent.store.displayname.ToString)

        If senderDomain = "asltd.com" Then
            Dim resp As ASLmessageProperties = ASL_Tools.Get_StampProperty(em)

            If resp.proj = "" Then
                MsgBox("No message key set")
            Else
                MsgBox("Project: " & resp.proj & vbLf & "TimeStamp: " & resp.timestamp & vbLf & "Type: " & resp.messagetype)
            End If
        End If

    End Sub
End Class

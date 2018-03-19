Imports System.Diagnostics
Imports Microsoft.Office.Tools.Ribbon
Imports System.Reflection

Public Class ASL_Ribbon

    Private Sub ASL_Ribbon_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        checkForNetwork()

        Dim bob = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version

        label_version.Label = bob.ToString

        ASL_Tools.GetOfflineMessages()
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
        'test
        ASL_Tools.ProcessSentMessages()

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

        Dim resp = MsgBox("This can be a long process. Do you wish to continue?", vbYesNo, "Get Offline Files")
        If resp = vbYes Then
            ASL_Tools.GetOfflineMessages()
        End If

        label_offlinefilecount.Label = ASL_Tools.offlineFileCount.ToString
    End Sub

    Private Sub Button4_Click(sender As Object, e As RibbonControlEventArgs) Handles button_pushOfflineFilestoServer.Click

        'push offline files to server

        If ASL_Tools.aslDiscipline = "" Then
            MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline.", vbCritical, "Error")
            Exit Sub
        End If

        ASL_Tools.ProcessOfflineMessages()

        ASL_Tools.GetOfflineMessages()
        label_offlinefilecount.Label = ASL_Tools.offlineFileCount


    End Sub

    Private Sub button_recordEmail_Click(sender As Object, e As RibbonControlEventArgs) Handles button_recordEmail.Click
        If ASL_Tools.aslDiscipline = "" Then
            MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline.", vbCritical, "Error")
            Exit Sub
        End If

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

        Dim projNumber As String = ""

        If ASL_Tools.aslDiscipline = "" Then
            MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline.", vbCritical, "Error")
            Exit Sub
        End If

        'allow the user to move emails from one project to another.
        'check to see if they are in a project directory or in a sent folder under the project directory
        'you cannot move an email if you are not connected to the network
        If Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        If Not (ASL_Tools.networkReady) Then
            MsgBox("You need to be connected to the network to move emails to another project.", vbCritical, "Error")
            Exit Sub
        End If

        'get the first message in the selection and get its project number.
        Dim firstMessage As Outlook.MailItem = Globals.ThisAddIn.Application.ActiveExplorer.Selection(1)
        Dim firstMailItem As class_MailItemTools = New class_MailItemTools
        firstMailItem.maItem = firstMessage
        firstMailItem.Get_PropertyAccessorObj()
        Dim oldProj As String = firstMailItem.proj

        'get the new project number
        Dim frm As form_emailMove = New form_emailMove
        frm.ShowDialog()

        projNumber = frm.proj

        If projNumber = "" Then
            'no project number selected
            Exit Sub
        ElseIf frm.proj = oldProj Then
            'the user selected the same project number
            Exit Sub
        End If


        'get all the selected messages and store them in a colleciton
        Dim emList As List(Of Outlook.MailItem) = New List(Of Outlook.MailItem)
        For Each obj As Outlook.MailItem In Globals.ThisAddIn.Application.ActiveExplorer.Selection
            emList.Add(obj)
        Next
        Globals.ThisAddIn.Application.ActiveExplorer.ClearSelection()

        'process the all the messages with the new project number
        For Each em As Outlook.MailItem In emList
            Dim fld As String = em.Parent.FullFolderPath
            Dim fldObj() As String = fld.Split("\")

            Dim ma As class_MailItemTools = New class_MailItemTools
            ma.maItem = em
            ma.Get_PropertyAccessorObj()

            If ma.messagetype = "se" Then
                Dim fldInbox As String = fldObj(fldObj.Length - 3)
                If fldInbox = "Inbox" Then
                    ma.Move_MailItem_OnStore(frm.proj)
                End If
            ElseIf ma.messagetype = "re" Then
                Dim fldinbox As String = fldObj(fldObj.Length - 2)
                If fldinbox = "Inbox" Then
                    ma.Move_MailItem_OnStore(frm.proj)
                End If
            Else
                MsgBox("Message is damaged and cannot be moved." & vbLf & em.Subject, vbCritical, "Error")
            End If
        Next
        frm.Close()

    End Sub

    Private Sub button_getUserProperties_Click(sender As Object, e As RibbonControlEventArgs) Handles button_getUserProperties.Click
        If Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        'Debug.Print(Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count.ToString)
        'get the first selected item.
        'if the store the item resides in is in the domain name asltd.com continue

        Dim em As Outlook.MailItem = Globals.ThisAddIn.Application.ActiveExplorer.Selection.Item(1)

        Dim senderDomain As String = ASL_Tools.Get_Domain_From_Address(em.Parent.store.displayname.ToString)

        If senderDomain = "asltd.com" Then
            Dim resp As class_MailItemTools = New class_MailItemTools

            resp.maItem = em
            resp.Get_PropertyAccessorObj()

            If resp.proj = "" Then
                MsgBox("No message key set")
            Else
                MsgBox("Project: " & resp.proj & vbLf & "TimeStamp: " & resp.timestamp & vbLf & "Type: " & resp.messagetype & vbLf & "Processed: " & resp.processed & vbLf & "Stored: " & resp.stored)
            End If
        End If

    End Sub

    Private Sub button_discipline_Click(sender As Object, e As RibbonControlEventArgs) Handles button_discipline.Click
        'open the dialogue box to show the discipline
        Dim frm As form_setDiscipline = New form_setDiscipline

        frm.ShowDialog()

        frm.Close()
    End Sub

    Private Sub Button5_Click(sender As Object, e As RibbonControlEventArgs) Handles Button5.Click
        If Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

        'Debug.Print(Globals.ThisAddIn.Application.ActiveExplorer.Selection.Count.ToString)
        'get the first selected item.
        'if the store the item resides in is in the domain name asltd.com continue

        Dim em As Outlook.MailItem = Globals.ThisAddIn.Application.ActiveExplorer.Selection.Item(1)

        Dim resp As class_MailItemTools = New class_MailItemTools

        resp.maItem = em
        resp.Reset_PropertyAccessorObj()

    End Sub
End Class

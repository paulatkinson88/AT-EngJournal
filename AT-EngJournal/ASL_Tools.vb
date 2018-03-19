Imports System.Diagnostics
Imports System.Runtime.InteropServices

Module ASL_Tools
    Public app As Outlook.Application

    Public dirObj As String = "J:\"
    Public networkReady As Boolean = False
    Public aslStore As Outlook.Store

    Public offlineFileCount As Integer = 0

    Public aslDiscipline As String = ""
    Public aslSentMailFolder As Outlook.Folder

    'global working project folder reference
    Public wProjFld As Outlook.Folder

    'global list of files that are not put to network yet.
    Public msList As List(Of Outlook.MailItem)

    'global form variable used for ViewOffLineFiles
    Public offLineFileForm As form_ViewOfflineFiles = New form_ViewOfflineFiles

    Public Function Check_For_Network() As Boolean
        Dim retVal As Boolean = False

        'global variable in asl.vb
        ASL_Tools.networkReady = False
        Dim fnd1 As Boolean = False
        Dim fnd2 As Boolean = False


        For Each dr In System.IO.DriveInfo.GetDrives()
            'If dr.Name = "J:\" And dr.DriveType.ToString = "Network" And dr.IsReady = True Then
            If dr.Name = ASL_Tools.dirObj And dr.DriveType.ToString = "Network" And dr.IsReady = True Then
                'there is a j drive that is networked.
                'check to see that the information in the drive is consistent with ASL J drive info.
                For Each dio In System.IO.Directory.GetDirectories(dr.Name)
                    'Debug.Print(dirObj)
                    Select Case dio.ToUpper
                        Case "J:\18XX"
                            fnd1 = True
                        Case "J:\19XX"
                            fnd2 = True
                        Case "K:\18XX"
                            fnd1 = True
                        Case "K:\19XX"
                            fnd2 = True
                    End Select
                Next

                'if the j drive has the project directories in it then we can assume
                'the j drive is the correct drive.
                If fnd1 = True And fnd2 = True Then
                    ASL_Tools.networkReady = True
                    retVal = True
                End If
            End If
        Next

        Return retVal
    End Function

    Public Function Get_Domain_From_Address(email As String) As String
        Dim retVal As String = ""
        Dim pnt As Integer = email.IndexOf("@")
        If Not (pnt = -1) Then
            Dim test As String = email.Substring(pnt + 1, email.Length - pnt - 1)
            Debug.Print(test)
            retVal = test
        End If
        Return retVal
    End Function

    '=====================
    'Finding the ASL Store
    '=====================
    ''' <summary>
    ''' Finds the asl store in all the stores on the outlook application session
    ''' </summary>
    Public Sub Get_ASL_Store()
        Dim app As Outlook.Application = Globals.ThisAddIn.Application

        Dim stores As Outlook.Stores = app.Session.Stores

        For Each st As Outlook.Store In stores
            Debug.Print(st.DisplayName & " - " & st.FilePath)

            If st.DisplayName.Contains("@asltd.com") Then
                ASL_Tools.aslStore = st
                Exit For
            End If
        Next

    End Sub

    Public Function Get_ASL_Store_SentItemsFolder() As Outlook.Folder
        Dim retVal As Outlook.Folder = Nothing


        Enumerate_SentItemsFolder(aslStore.GetRootFolder)

        If Not (IsNothing(aslSentMailFolder)) Then
            retVal = aslSentMailFolder
        End If

        Return retVal
    End Function

    Private Sub Enumerate_SentItemsFolder(ByVal oFolder As Outlook.Folder)
        Dim folders As Outlook.Folders = Nothing
        Dim Folder As Outlook.Folder = Nothing
        Dim foldercount As Integer = Nothing

        On Error Resume Next
        folders = oFolder.Folders
        foldercount = folders.Count
        'Check if there are any folders below oFolder 
        If foldercount Then
            For Each Folder In folders
                'Debug.Print(Folder.FolderPath)
                If Folder.Name = "Sent Mail" Then
                    aslSentMailFolder = Folder
                    Exit Sub
                End If

                Enumerate_SentItemsFolder(Folder)
            Next
        End If
    End Sub

    ''' <summary>
    ''' gets the discipline from the registry
    ''' </summary>
    ''' <returns></returns>
    Public Function get_discipline() As String
        Dim disc As String = ""

        Dim tempdisc As String = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\AT_EngJournal", "discipline", "")
        If Not (tempdisc = "") Then
            disc = tempdisc
        End If
        Return disc
    End Function

    ''' <summary>
    ''' sets the discipline to the registry
    ''' </summary>
    ''' <param name="disc"></param>
    Public Sub set_discipline(disc)
        If Not (disc = "") Then
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\AT_EngJournal", "discipline", disc)
        End If
    End Sub

    ''' <summary>
    ''' enables everything to do with the discipline and sets the registry
    ''' </summary>
    ''' <param name="disc"></param>
    Public Sub enable_discipline(disc As String)
        'store the discipline in the registry
        ASL_Tools.set_discipline(disc)

        'set teh discipline in the global variable
        ASL_Tools.aslDiscipline = disc

        'update the ribbon bar to show the discipline
        Globals.Ribbons.ASL_Ribbon.label_discipline.Label = disc

        'enable the buttons on the ribbon bar.
        Globals.Ribbons.ASL_Ribbon.button_MoveEmail.Enabled = True
        Globals.Ribbons.ASL_Ribbon.button_pushOfflineFilestoServer.Enabled = True
        Globals.Ribbons.ASL_Ribbon.button_recordEmail.Enabled = True
    End Sub

    ''' <summary>
    ''' disables everyting to do with the registry
    ''' </summary>
    Public Sub disable_discipline()
        MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline.", vbCritical, "Error")

        'set teh discipline in the global variable
        ASL_Tools.aslDiscipline = ""

        'update the ribbon bar to show the discipline
        Globals.Ribbons.ASL_Ribbon.label_discipline.Label = "-"

        'enable the buttons on the ribbon bar.
        Globals.Ribbons.ASL_Ribbon.button_MoveEmail.Enabled = False
        Globals.Ribbons.ASL_Ribbon.button_pushOfflineFilestoServer.Enabled = False
        Globals.Ribbons.ASL_Ribbon.button_recordEmail.Enabled = False

    End Sub

    Public Sub GetOfflineMessages()
        Dim frm As form_process = New form_process

        frm.Show()

        ASL_Tools.offlineFileCount = 0

        ASL_Tools.msList = New List(Of Outlook.MailItem)

        If IsNothing(ASL_Tools.aslStore) Then
            ASL_Tools.Get_ASL_Store()
        End If

        If IsNothing(ASL_Tools.aslStore) Then
            MsgBox("Unable to get ASL Store.", vbCritical, "Error")
            Exit Sub
        End If

        Dim fldRoot As Outlook.Folder = ASL_Tools.aslStore.GetRootFolder
        Dim fldIn As Outlook.Folder = Nothing

        frm.Refresh()

        For Each fld As Outlook.Folder In fldRoot.Folders
            If fld.Name = "Inbox" Then
                fldIn = fld
                Exit For
            End If
        Next

        ASL_Tools.Enumerate_InboxFolder(fldIn)

        frm.Close()
    End Sub

    Private Sub Enumerate_InboxFolder(ByVal oFolder As Outlook.Folder)
        Dim folders As Outlook.Folders = Nothing
        Dim Folder As Outlook.Folder = Nothing
        Dim foldercount As Integer = Nothing

        On Error Resume Next
        folders = oFolder.Folders
        foldercount = folders.Count
        'Check if there are any folders below oFolder 
        If foldercount Then
            For Each Folder In folders
                'Debug.Print(Folder.FolderPath)
                For Each msIt As Outlook.MailItem In Folder.Items

                    If (TypeOf msIt Is Outlook.MailItem) Then
                        'put the mail item into the custom class mail item
                        Dim ma As class_MailItemTools = New class_MailItemTools
                        ma.maItem = msIt
                        ma.Get_PropertyAccessorObj()
                        Dim sto As String = ma.stored
                        If sto = "False" Then
                            ASL_Tools.offlineFileCount = ASL_Tools.offlineFileCount + 1
                            ASL_Tools.msList.Add(msIt)
                        End If
                    End If
                Next
                Enumerate_InboxFolder(Folder)
            Next
        End If
    End Sub

    Public Sub ProcessOfflineMessages()
        'for each file that is offline. push them to the server.
        If ASL_Tools.networkReady = True Then
            For Each obj As Outlook.MailItem In ASL_Tools.msList
                If (TypeOf obj Is Outlook.MailItem) Then
                    'put the mail item into the custom class mail item
                    Dim ma As class_MailItemTools = New class_MailItemTools
                    ma.maItem = obj
                    ma.Get_PropertyAccessorObj()

                    Try
                        'get the other custom properties of the message.
                        If ma.stored = False Then
                            Try
                                ma.Put_MailItem_OnNetwork()
                            Catch ex As Exception
                                MsgBox("ProcessOffline.NetworkStore:" & ex.Message)
                            End Try
                        End If
                    Catch ex As Exception
                        MsgBox("ProcessOffline.properties:" & ex.Message)
                    End Try

                End If
            Next
        End If

        ASL_Tools.GetOfflineMessages()

    End Sub

    Public Sub SaveMessage(ma As class_MailItemTools, proj As String)
        Dim ret As Outlook.MailItem = Nothing

        Try
            ma.Get_PropertyAccessorObj()
            'get the other custom properties of the message.
            'ma.msgProp.Get_AllProperties(ma.maItem)

            'ma.msgProp.Set_ProjectProperty(proj, ma.maItem)
            Dim st As String = ma.Format_DateTimeStamp()
            ma.timestamp = st
            ma.messagetype = "re"
            ma.processed = "False"
            ma.stored = "False"

            ma.Set_PropertyAccessorObj()

            'ma.msgProp.Set_TimeStampProperty(st, ma.maItem)
            'ma.msgProp.Set_MessageTypeProperty("re", ma.maItem)
            'ma.msgProp.Set_ProcessedProperty("False", ma.maItem)
            'ma.msgProp.Set_StoredProperty("False", ma.maItem)

            Try
                ret = ma.Store_MailItem_OnStore()
                If Not (IsNothing(ret)) Then
                    ma.Put_MailItem_OnNetwork()
                End If
            Catch ex As Exception
                MsgBox("Record.NetworkStore:" & ex.Message)
            End Try
        Catch ex As Exception
            MsgBox("Record.properties:" & ex.Message)
        End Try

    End Sub

    Public Sub ProcessSentMessages()
        'get the sent message folder
        'for each message that was sent in the last 2 weeks look at the properties.
        'look for messages with a category "Project=0000" these have need to be recorded.
        'if a message needs to be recorded then get all the message properties
        'save it to the message store under the job number.
        'save it to the network under the job number.

        Dim sDate As Date = Now.AddDays(-14)
        Dim searchCriteria As String = "[ReceivedTime]>'" & Format(sDate, "M/d/yyyy H:mm") & "'"
        Dim cat As String = ""              'category
        Dim resultItem As Object = Nothing

        'Apply filter. use the receivedtime of the email
        resultItem = ASL_Tools.aslSentMailFolder.Items.Restrict(searchCriteria)

        For Each obj As Object In resultItem
            If (TypeOf obj Is Outlook.MailItem) Then
                'put the mail item into the custom class mail item
                Dim ma As class_MailItemTools = New class_MailItemTools
                ma.maItem = obj

                'try and get the custom category of the mail item that was stored when sent.
                Try
                    ma.Get_PropertyAccessorObj()

                    If ma.proj.Length > 1 Then
                        Dim st As String = ma.Format_DateTimeStamp()
                        ma.timestamp = st
                        ma.messagetype = "se"
                        ma.processed = "False"
                        ma.stored = "False"

                        ma.Set_PropertyAccessorObj()

                        'once all the mail properties have been set
                        'try to process the items.
                        'if it fails then reset the processed property
                        Try
                            ma.Store_MailItem_OnStore()
                            ma.Put_MailItem_OnNetwork()
                        Catch ex As Exception
                            MsgBox("MailProcess.MailStore:" & ex.Message)
                        End Try

                    End If
                Catch ex As Exception
                    MsgBox("MailProcess.Category:" & ex.Message)
                End Try

            End If
        Next

    End Sub
End Module

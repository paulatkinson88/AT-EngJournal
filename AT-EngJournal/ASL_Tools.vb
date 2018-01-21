Imports System.Diagnostics

Module ASL_Tools
    Public app As Outlook.Application

    Public dirObj As String = "K:\"
    Public networkReady As Boolean = False
    Public aslStore As Outlook.Store

    Public offlineFileCount As Integer = 0

    'global working project folder reference
    Public wProjFld As Outlook.Folder

    Public msList As List(Of Outlook.MailItem)

    Public Sub Get_OffLineFileCount()
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

        For Each fld As Outlook.Folder In fldRoot.Folders
            If fld.Name = "Inbox" Then
                fldIn = fld
                Exit For
            End If
        Next

        ASL_Tools.EnumerateFolders(fldIn)


    End Sub

    Private Sub EnumerateFolders(ByVal oFolder As Outlook.Folder)
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
                    If Not (IsNothing(msIt.Categories)) Then
                        'Debug.Print(msIt.Categories)

                        Dim cat As String = msIt.Categories.ToUpper
                        If Not (cat.IndexOf("OFFLINE") = -1) Then

                            ASL_Tools.offlineFileCount = ASL_Tools.offlineFileCount + 1
                            ASL_Tools.msList.Add(msIt)
                        End If
                    End If
                Next
                EnumerateFolders(Folder)
            Next
        End If
    End Sub

    Public Sub Check_OfflineCategory()
        Dim fndOC As Outlook.Category = Nothing

        For Each oc As Outlook.Category In ASL_Tools.app.Session.Categories
            If oc.Name = "OffLine" Then
                fndOC = oc
                Exit For
            End If
        Next

        If IsNothing(fndOC) Then
            ASL_Tools.app.Session.Categories.Add("OffLine", Outlook.OlCategoryColor.olCategoryColorDarkRed)
        End If
    End Sub

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

    Public Function Check_For_ProjectDirectoryEngJournal(pro As String) As System.IO.DirectoryInfo
        Dim retVal As System.IO.DirectoryInfo = Nothing
        Dim di As System.IO.DirectoryInfo = Nothing

        Dim projDirectory As String = dirObj & pro.Substring(0, 2) & "XX\" & pro
        'if the project directory exists then move to next check
        If System.IO.Directory.Exists(projDirectory) Then
            If Not (System.IO.Directory.Exists(projDirectory & "\EngJournal")) Then
                'check to see if the engineering journal exists.
                'if it does not exist then create it.
                di = New System.IO.DirectoryInfo(projDirectory)
                di = di.CreateSubdirectory("EngJournal")
            Else
                di = New System.IO.DirectoryInfo(projDirectory & "\EngJournal")
            End If
            'add the message file.

            retVal = di

        Else
            'project directory missing
            MsgBox("Project Directory does not exist", vbCritical, "Error")
        End If

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

    ''' <summary>
    ''' Test function not used.
    ''' </summary>
    Public Sub Get_ASL_Store_Folders()
        If IsNothing(ASL_Tools.aslStore) Then
            ASL_Tools.Get_ASL_Store()
        End If

        If IsNothing(ASL_Tools.aslStore) Then
            MsgBox("Unable to get ASL Store.", vbCritical, "Error")
            Exit Sub
        End If

        Dim rtFld As Outlook.Folder = ASL_Tools.aslStore.GetRootFolder

        For Each fld As Outlook.Folder In rtFld.Folders
            If fld.Name = "Inbox" Then
                For Each subFld As Outlook.Folder In fld.Folders
                    Debug.Print("     Sub Name: " & subFld.Name)
                Next
            End If
            Debug.Print("   Name: " & fld.Name)
        Next
    End Sub

    '====================
    'ASL Store Functions
    '====================
    ''' <summary>
    ''' Finds the project folder in the ASL store and returns a reference to it
    ''' </summary>
    ''' <param name="proj"></param>
    ''' <returns></returns>
    Public Function Get_ProjectFolder_In_ASLStoreInbox(ByVal proj As String) As Outlook.Folder
        Dim retVal As Outlook.Folder = Nothing

        If IsNothing(ASL_Tools.aslStore) Then
            ASL_Tools.Get_ASL_Store()
        End If

        If Not (IsNothing(ASL_Tools.aslStore)) Then
            'if the store is in memory then look at the root of the folder
            'then select the inbox foulder and get the sub folders in it.

            Dim rtFld As Outlook.Folder = ASL_Tools.aslStore.GetRootFolder

            For Each fld As Outlook.Folder In rtFld.Folders
                If fld.Name = "Inbox" Then
                    For Each subFld As Outlook.Folder In fld.Folders
                        If subFld.Name.Substring(0, 4) = proj Then
                            'if the first four characters of the folder name
                            'match the project number then we have found our folder

                            retVal = subFld
                            Exit For

                        End If
                    Next
                End If
            Next

        Else
            MsgBox("Unable to get ASL Store.", vbCritical, "Error")
        End If

        Return retVal
    End Function

    ''' <summary>
    ''' creates a new project folder in the inbox and returns the reference.
    ''' </summary>
    ''' <param name="proj"></param>
    ''' <returns></returns>
    Public Function Create_ProjectFolder_In_ASLStoreInbox(ByVal proj As String) As Outlook.Folder
        Dim retVal As Outlook.Folder = Nothing

        If IsNothing(ASL_Tools.aslStore) Then
            ASL_Tools.Get_ASL_Store()
        End If

        If Not (IsNothing(ASL_Tools.aslStore)) Then
            'if the store is in memory then look at the root of the folder
            'then select the inbox folder and get the sub folders in it.

            Dim rtFld As Outlook.Folder = ASL_Tools.aslStore.GetRootFolder

            For Each fld As Outlook.Folder In rtFld.Folders
                If fld.Name = "Inbox" Then
                    retVal = fld.Folders.Add(proj)
                    Exit For
                End If
            Next

        Else
            MsgBox("Unable to get ASL Store.", vbCritical, "Error")
        End If

        Return retVal
    End Function

    ''' <summary>
    ''' Checks to see if the ProjectFolderSend exists a passed outlook folder reference
    ''' </summary>
    ''' <param name="pf"></param>
    ''' <returns></returns>
    Public Function Get_ProjectFolderSent(ByVal pf As Outlook.Folder) As Outlook.Folder
        Dim retVal As Outlook.Folder = Nothing

        If Not (IsNothing(pf)) Then
            For Each fld As Outlook.Folder In pf.Folders
                If fld.Name = "SENT" Then
                    retVal = fld
                    Exit For
                End If
            Next
        End If

        Return retVal
    End Function

    ''' <summary>
    ''' create sent folder to a ProjectFolder passed by the user
    ''' check to see if one exists first.
    ''' </summary>
    ''' <param name="pf"></param>
    ''' <returns></returns>
    Public Function Create_ProjectFolderSent(ByVal pf As Outlook.Folder) As Outlook.Folder
        Dim fld As Outlook.Folder = Nothing

        fld = pf.Folders.Add("SENT")

        Return fld
    End Function

    Public Function get_ProjectNumber_From_MsgFolder(ms As Outlook.MailItem) As String
        Dim retVal As String = ""

        'ms.

        Return retVal
    End Function
End Module

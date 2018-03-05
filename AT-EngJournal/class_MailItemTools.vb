Imports System.Diagnostics

Public Class class_MailItemTools
    Public maItem As Outlook.MailItem
    Public msgProp As class_MailProperties

    Public proj As String
    Public timestamp As String
    Public messagetype As String
    Public processed As String
    Public stored As String

    Public Sub New()
        msgProp = New class_MailProperties

        proj = ""
        timestamp = ""
        messagetype = ""
        processed = ""
        stored = ""
    End Sub

    Public Sub show_properties()
        MsgBox("id: " & maItem.EntryID & vbLf & "Project: " & proj & vbLf & "TimeStamp: " & timestamp & vbLf & "Type: " & messagetype & vbLf & "Processed: " & processed & vbLf & "Stored: " & stored)
    End Sub

    Public Sub Set_PropertyAccessorObj()
        Dim prop1 = "http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/78965"
        'Dim propSt = "111-222-Kudo"

        'Dim prop As String = "http://schemas.microsoft.com/mapi/string/{00020386-0000-0000-C000-000000000046}/78965"
        Dim propSt As String = "(" & proj & ")(" & timestamp & ")(" & messagetype & ")(" & processed & ")(" & stored & ")()()"
        maItem.PropertyAccessor.SetProperty(prop1, propSt)
        maItem.Save()
    End Sub

    Public Sub Get_PropertyAccessorObj()
        'PR_TRANSPORT_MESSAGE_HEADERS
        Dim PropName = "http://schemas.microsoft.com/mapi/string/{FFF40745-D92F-4C11-9E14-92701F001EB3}/78965"
        Try
            Dim retVal As String = maItem.PropertyAccessor.GetProperty(PropName)
            Debug.Print(retVal)
            Dim arObj = retVal.Split("(")
            If Not (arObj.Count = 0) Then
                proj = arObj(0)
                If arObj(1).Length > 0 Then
                    timestamp = arObj(1).Substring(0, arObj(1).Length - 1)
                End If
                If arObj(2).Length > 0 Then
                    messagetype = arObj(2).Substring(0, arObj(2).Length - 1)
                End If
                If arObj(3).Length > 0 Then
                    processed = arObj(3).Substring(0, arObj(3).Length - 1)
                End If
                If arObj(4).Length > 0 Then
                    stored = arObj(4).Substring(0, arObj(4).Length - 1)
                End If
            End If
        Catch ex As Exception
            proj = ""
            timestamp = ""
            messagetype = ""
            processed = ""
            stored = ""
        End Try

    End Sub

    '###################################################################
    '####   outlook store section
    '###################################################################
    Public Function Store_MailItem_OnStore()
        Dim retVal As Outlook.MailItem = Nothing
        Debug.Print("Store_MailItem_OnStore")
        'copy the message to the folder
        '   find the projectfolder in the aslstore
        Debug.Print("Store_MailItem_OnStore - proj:" & msgProp.proj)
        Dim fld As Outlook.Folder = Get_ProjectFolder_OnStoreInbox(msgProp.proj)
        Debug.Print("Store_MailItem_OnStore - folder:" & fld.FolderPath)
        If IsNothing(fld) Then
            'create the project folder
            fld = Create_ProjectFolder_OnStoreInbox(msgProp.proj)

            If IsNothing(fld) Then
                MsgBox("Error creating project in Inbox", vbCritical, "Error")
                Return retVal
            End If
        End If

        Try
            'if the message is a sent item then change the store folder to the sent item.
            If msgProp.messagetype = "se" Then fld = Get_ProjectFolderSent_OnStoreInbox(fld)

            'if the folder is there then move the item
            If Not (IsNothing(fld)) Then
                Debug.Print("Store_MailItem_OnStore - move:" & maItem.EntryID)
                maItem = maItem.Move(fld)
                Debug.Print("Store_MailItem_OnStore - to:" & maItem.EntryID)
                retVal = maItem
                msgProp.Set_ProcessedProperty(True, maItem)
                Debug.Print("Store_MailItem_OnStore - processed:" & msgProp.processed)
                Debug.Print("Store_MailItem_OnStore - save")
                maItem.Save()
            End If

        Catch ex As Exception
            MsgBox("MailItem.Store_MailItem_OnStore:" & ex.Message)
        End Try

        Return retVal
    End Function

    Public Function Move_MailItem_OnStore(newProj As String)
        Dim retVal As Outlook.MailItem = Nothing
        Debug.Print("move_mailitem_onstore")
        'get the project folders for the move.
        Dim newFld As Outlook.Folder = Get_ProjectFolder_OnStoreInbox(newProj)
        Debug.Print("move_mailitem_onstore - New Proj:" & newProj)

        If IsNothing(newFld) Then
            newFld = Create_ProjectFolder_OnStoreInbox(newProj)

            If IsNothing(newFld) Then
                MsgBox("Error creating project in Inbox", vbCritical, "Error")
                Return retVal
            End If
        End If


        Try
            'remove anything on the network
            Debug.Print("move_mailitem_onstore - remove from network")
            Remove_MailItem_OnNetwork()

            'if the message is a sent item then change the store folder to the sent item.
            If msgProp.messagetype = "se" Then
                newFld = Get_ProjectFolderSent_OnStoreInbox(newFld)
            End If

            'if the folder is there then move the item
            If Not (IsNothing(newFld)) Then
                Debug.Print("move_mailitem_onstore - move " & maItem.EntryID)
                retVal = maItem.Move(newFld)
                Debug.Print("move_mailitem_onstore - " & retVal.EntryID)
                maItem = retVal

                msgProp.Set_ProjectProperty(newProj, maItem)
                msgProp.Set_ProcessedProperty(True, maItem)
                maItem.Save()
                Debug.Print("move_mailitem_onstore - save")
                'once moved, store it on the network
                Put_MailItem_OnNetwork()
            End If

        Catch ex As Exception
            MsgBox("MailItem.Store_MailItem_OnStore:" & ex.Message)
            msgProp.Set_ProcessedProperty(False, maItem)
        End Try

        Return retVal
    End Function

    Private Function Create_ProjectFolder_OnStoreInbox(newProj As String) As Outlook.Folder
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
                    retVal = fld.Folders.Add(newProj)
                    retVal.Folders.Add("SENT")
                    Exit For
                End If
            Next

        Else
            MsgBox("Unable to get ASL Store.", vbCritical, "Error")
        End If

        Return retVal
    End Function

    Private Function Get_ProjectFolder_OnStoreInbox(newProj As String) As Outlook.Folder
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
                        If subFld.Name.Substring(0, 4) = newProj Then
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

    Private Function Get_ProjectFolderSent_OnStoreInbox(fld As Outlook.Folder) As Outlook.Folder
        Dim retVal As Outlook.Folder = Nothing

        If Not (IsNothing(fld)) Then
            For Each subFld As Outlook.Folder In fld.Folders
                If subFld.Name = "SENT" Then
                    retVal = subFld
                End If
            Next
        End If

        Return retVal
    End Function
    '###################################################################
    '####   network section
    '###################################################################
    Public Function Put_MailItem_OnNetwork() As Boolean
        'return true if successful or false if not.
        Dim retVal As Boolean = False

        'automatically get the project correspondence directory and the users name.
        Dim diPath As System.IO.DirectoryInfo = Get_ProjectDirectory_OnNetwork()

        If IsNothing(diPath) Then
            Return retVal
        End If

        Dim st As String = "(" & msgProp.proj & ")(" & msgProp.timestamp & ")(" & msgProp.messagetype & ")"

        'check to see if the user is in the office.
        'if they are then save a copy of the email to the project folder
        'if not then flag the message with the category offline
        'offline messages can get copied to network at a later date.


        If ASL_Tools.networkReady = True Then
            Try
                stored = "True"
                Set_PropertyAccessorObj()

                'msgProp.Set_StoredProperty(True, maItem)
                maItem.Save()
                maItem.SaveAs(diPath.FullName & "\" & st & ".msg")

                retVal = True
            Catch ex As Exception
                MsgBox("MailItem.Put_MailItem_OnNetwork:" & ex.Message)
                msgProp.Set_StoredProperty(False, maItem)
            End Try
        Else
            msgProp.Set_StoredProperty(False, maItem)
        End If

        Return retVal
    End Function

    Public Sub Remove_MailItem_OnNetwork()
        Dim fld As System.IO.DirectoryInfo = Get_ProjectDirectory_OnNetwork()

        Dim st As String = "(" & msgProp.proj & ")(" & msgProp.timestamp & ")(" & msgProp.messagetype & ")"
        Dim fl = (fld.FullName & "\" & st & ".msg")

        If System.IO.File.Exists(fl) Then
            System.IO.File.Delete(fl)
        End If

    End Sub

    Public Function Get_ProjectDirectory_OnNetwork() As System.IO.DirectoryInfo
        Dim retVal As System.IO.DirectoryInfo = Nothing
        Dim di As System.IO.DirectoryInfo = Nothing

        Dim disc As String = ""
        Dim discPath As String = ""

        Select Case ASL_Tools.aslDiscipline
            Case = "Electrical"
                disc = "E"
                discPath = "\E\CORRESPONDENCE"
            Case = "Mechanical"
                disc = "M"
                discPath = "\M\CORRESPONDENCE"
            Case = "Structural"
                disc = "S"
                discPath = "\S\CORRESPONDENCE"
            Case Else
                MsgBox("No Discipline set." & vbLf & "Click the Change Discipline button on the ASL Ribbon bar and set the Discipline.", vbCritical, "Error")
                Return di
                Exit Function
        End Select

        Dim projDirectory As String = dirObj & msgProp.proj.Substring(0, 2) & "XX\" & msgProp.proj
        Dim username As String = ASL_Tools.aslStore.DisplayName

        'if the project directory exists then move to next check
        If System.IO.Directory.Exists(projDirectory) Then
            If Not (System.IO.Directory.Exists(projDirectory & discPath & "\" & username)) Then
                'check to see if the CORRESPONDENCE DIRECTORY exists.
                'if it does not exist then create it.
                di = New System.IO.DirectoryInfo(projDirectory)
                di = di.CreateSubdirectory(disc)
                di = di.CreateSubdirectory("CORRESPONDENCE")
                di = di.CreateSubdirectory(username)
            Else
                di = New System.IO.DirectoryInfo(projDirectory & discPath & "\" & username)
            End If
            'add the message file.

            retVal = di

        Else
            'project directory missing
            MsgBox("Project Directory on the network does not exist", vbCritical, "Error")
        End If

        Return retVal
    End Function
End Class

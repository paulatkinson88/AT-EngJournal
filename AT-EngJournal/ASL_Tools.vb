Imports System.Diagnostics

Module ASL_Tools
    Public app As Outlook.Application

    Public networkReady As Boolean = False
    Public aslStore As Outlook.Store


    Public Function Check_For_Network() As Boolean
        Dim retVal As Boolean = False

        'global variable in asl.vb
        ASL_Tools.networkReady = False
        Dim fnd1 As Boolean = False
        Dim fnd2 As Boolean = False


        For Each dr In System.IO.DriveInfo.GetDrives()
            If dr.Name = "J:\" And dr.DriveType.ToString = "Network" And dr.IsReady = True Then
                'there is a j drive that is networked.
                'check to see that the information in the drive is consistent with ASL J drive info.
                For Each dirObj In System.IO.Directory.GetDirectories(dr.Name)
                    'Debug.Print(dirObj)
                    Select Case dirObj.ToUpper
                        Case "J:\18XX"
                            fnd1 = True
                        Case "J:\19XX"
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

    Public Function Check_For_Project_Directory(pro As String) As Boolean
        Dim retVal As Boolean = False

        Dim projDirectory As String = "J:\" & pro.Substring(0, 2) & "XX\" & pro
        'if the project directory exists then move to next check
        If System.IO.Directory.Exists(projDirectory) Then
            If Not (System.IO.Directory.Exists(projDirectory & "\EngJournal")) Then
                'check to see if the engineering journal exists.
                'if it does not exist then create it.

            End If
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

    Public Function Get_ProjectFolder_From_ASL_Store_Inbox(ByVal proj As String) As Outlook.Folder
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
End Module

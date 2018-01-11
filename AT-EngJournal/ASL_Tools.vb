Imports System.Diagnostics

Module ASL_Tools
    Public app As Outlook.Application

    Public networkReady As Boolean = False


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
End Module

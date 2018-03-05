Public Class class_StoreTools

    Public projFolder As Outlook.Folder





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

    Public Function Create_ProjectFolderSent(ByVal pf As Outlook.Folder) As Outlook.Folder
        Dim fld As Outlook.Folder = Nothing

        fld = pf.Folders.Add("SENT")

        Return fld
    End Function
End Class

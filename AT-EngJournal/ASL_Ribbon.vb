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
    End Sub
End Class

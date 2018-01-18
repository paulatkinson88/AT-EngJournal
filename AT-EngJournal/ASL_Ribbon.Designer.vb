Partial Class ASL_Ribbon
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Tab1 = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.label_connection = Me.Factory.CreateRibbonLabel
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.label_offlineFileCount = Me.Factory.CreateRibbonLabel
        Me.ButtonGroup1 = Me.Factory.CreateRibbonButtonGroup
        Me.button_OfflineFilesCount = Me.Factory.CreateRibbonButton
        Me.Button4 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.ButtonGroup1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Label = "ASL"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.label_connection)
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Label = "Network"
        Me.Group1.Name = "Group1"
        '
        'label_connection
        '
        Me.label_connection.Label = "Connection:"
        Me.label_connection.Name = "label_connection"
        '
        'Button1
        '
        Me.Button1.Image = Global.AT_EngJournal.My.Resources.Resources.network_connect_3
        Me.Button1.Label = "Check"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Group2"
        Me.Group2.Name = "Group2"
        '
        'Button2
        '
        Me.Button2.Label = "Button2"
        Me.Button2.Name = "Button2"
        '
        'Button3
        '
        Me.Button3.Label = "Button3"
        Me.Button3.Name = "Button3"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.label_offlineFileCount)
        Me.Group3.Items.Add(Me.ButtonGroup1)
        Me.Group3.Items.Add(Me.Button4)
        Me.Group3.Label = "Offline Files"
        Me.Group3.Name = "Group3"
        '
        'label_offlineFileCount
        '
        Me.label_offlineFileCount.Label = "Offline File Count: -"
        Me.label_offlineFileCount.Name = "label_offlineFileCount"
        '
        'ButtonGroup1
        '
        Me.ButtonGroup1.Items.Add(Me.button_OfflineFilesCount)
        Me.ButtonGroup1.Name = "ButtonGroup1"
        '
        'button_OfflineFilesCount
        '
        Me.button_OfflineFilesCount.Image = Global.AT_EngJournal.My.Resources.Resources.face_wink_4
        Me.button_OfflineFilesCount.Label = "Refresh"
        Me.button_OfflineFilesCount.Name = "button_OfflineFilesCount"
        Me.button_OfflineFilesCount.ShowImage = True
        '
        'Button4
        '
        Me.Button4.Image = Global.AT_EngJournal.My.Resources.Resources.face_glasses
        Me.Button4.Label = "Push Offline Files to Server"
        Me.Button4.Name = "Button4"
        Me.Button4.ShowImage = True
        '
        'ASL_Ribbon
        '
        Me.Name = "ASL_Ribbon"
        Me.RibbonType = "Microsoft.Outlook.Explorer, Microsoft.Outlook.Mail.Compose, Microsoft.Outlook.Mai" &
    "l.Read, Microsoft.Outlook.Post.Compose, Microsoft.Outlook.Post.Read, Microsoft.O" &
    "utlook.Resend"
        Me.Tabs.Add(Me.Tab1)
        Me.Tab1.ResumeLayout(False)
        Me.Tab1.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.ButtonGroup1.ResumeLayout(False)
        Me.ButtonGroup1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents label_connection As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents label_offlineFileCount As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents ButtonGroup1 As Microsoft.Office.Tools.Ribbon.RibbonButtonGroup
    Friend WithEvents Button4 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents button_OfflineFilesCount As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ASL_Ribbon() As ASL_Ribbon
        Get
            Return Me.GetRibbon(Of ASL_Ribbon)()
        End Get
    End Property
End Class

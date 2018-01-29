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
        Me.Box1 = Me.Factory.CreateRibbonBox
        Me.label_con = Me.Factory.CreateRibbonLabel
        Me.label_connection = Me.Factory.CreateRibbonLabel
        Me.Group3 = Me.Factory.CreateRibbonGroup
        Me.Box2 = Me.Factory.CreateRibbonBox
        Me.label = Me.Factory.CreateRibbonLabel
        Me.Label2 = Me.Factory.CreateRibbonLabel
        Me.label_offlinefilecount = Me.Factory.CreateRibbonLabel
        Me.Group5 = Me.Factory.CreateRibbonGroup
        Me.Group2 = Me.Factory.CreateRibbonGroup
        Me.Group4 = Me.Factory.CreateRibbonGroup
        Me.Box3 = Me.Factory.CreateRibbonBox
        Me.Label1 = Me.Factory.CreateRibbonLabel
        Me.label_version = Me.Factory.CreateRibbonLabel
        Me.Button1 = Me.Factory.CreateRibbonButton
        Me.button_OfflineFilesCount = Me.Factory.CreateRibbonButton
        Me.button_pushOfflineFilestoServer = Me.Factory.CreateRibbonButton
        Me.button_recordEmail = Me.Factory.CreateRibbonButton
        Me.Button2 = Me.Factory.CreateRibbonButton
        Me.Button3 = Me.Factory.CreateRibbonButton
        Me.Tab1.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Box1.SuspendLayout()
        Me.Group3.SuspendLayout()
        Me.Box2.SuspendLayout()
        Me.Group5.SuspendLayout()
        Me.Group2.SuspendLayout()
        Me.Group4.SuspendLayout()
        Me.Box3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Tab1
        '
        Me.Tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.Tab1.Groups.Add(Me.Group1)
        Me.Tab1.Groups.Add(Me.Group3)
        Me.Tab1.Groups.Add(Me.Group5)
        Me.Tab1.Groups.Add(Me.Group2)
        Me.Tab1.Groups.Add(Me.Group4)
        Me.Tab1.Label = "ASL"
        Me.Tab1.Name = "Tab1"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.Button1)
        Me.Group1.Items.Add(Me.Box1)
        Me.Group1.Label = "Network"
        Me.Group1.Name = "Group1"
        '
        'Box1
        '
        Me.Box1.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box1.Items.Add(Me.label_con)
        Me.Box1.Items.Add(Me.label_connection)
        Me.Box1.Name = "Box1"
        '
        'label_con
        '
        Me.label_con.Label = "Connection:"
        Me.label_con.Name = "label_con"
        '
        'label_connection
        '
        Me.label_connection.Label = "-"
        Me.label_connection.Name = "label_connection"
        '
        'Group3
        '
        Me.Group3.Items.Add(Me.button_OfflineFilesCount)
        Me.Group3.Items.Add(Me.button_pushOfflineFilestoServer)
        Me.Group3.Items.Add(Me.Box2)
        Me.Group3.Label = "Offline Files"
        Me.Group3.Name = "Group3"
        '
        'Box2
        '
        Me.Box2.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box2.Items.Add(Me.label)
        Me.Box2.Items.Add(Me.Label2)
        Me.Box2.Items.Add(Me.label_offlinefilecount)
        Me.Box2.Name = "Box2"
        '
        'label
        '
        Me.label.Label = "Offline File"
        Me.label.Name = "label"
        '
        'Label2
        '
        Me.Label2.Label = "Count:"
        Me.Label2.Name = "Label2"
        '
        'label_offlinefilecount
        '
        Me.label_offlinefilecount.Label = "0"
        Me.label_offlinefilecount.Name = "label_offlinefilecount"
        '
        'Group5
        '
        Me.Group5.Items.Add(Me.button_recordEmail)
        Me.Group5.Label = "Record"
        Me.Group5.Name = "Group5"
        '
        'Group2
        '
        Me.Group2.Items.Add(Me.Button2)
        Me.Group2.Items.Add(Me.Button3)
        Me.Group2.Label = "Test"
        Me.Group2.Name = "Group2"
        Me.Group2.Visible = False
        '
        'Group4
        '
        Me.Group4.Items.Add(Me.Box3)
        Me.Group4.Label = "About"
        Me.Group4.Name = "Group4"
        '
        'Box3
        '
        Me.Box3.BoxStyle = Microsoft.Office.Tools.Ribbon.RibbonBoxStyle.Vertical
        Me.Box3.Items.Add(Me.Label1)
        Me.Box3.Items.Add(Me.label_version)
        Me.Box3.Name = "Box3"
        '
        'Label1
        '
        Me.Label1.Label = "Version:"
        Me.Label1.Name = "Label1"
        '
        'label_version
        '
        Me.label_version.Label = "-"
        Me.label_version.Name = "label_version"
        '
        'Button1
        '
        Me.Button1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.Button1.Image = Global.AT_EngJournal.My.Resources.Resources.network_connect_3
        Me.Button1.Label = "Check"
        Me.Button1.Name = "Button1"
        Me.Button1.ShowImage = True
        '
        'button_OfflineFilesCount
        '
        Me.button_OfflineFilesCount.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.button_OfflineFilesCount.Image = Global.AT_EngJournal.My.Resources.Resources.view_refresh_4
        Me.button_OfflineFilesCount.Label = "ReScan"
        Me.button_OfflineFilesCount.Name = "button_OfflineFilesCount"
        Me.button_OfflineFilesCount.ShowImage = True
        '
        'button_pushOfflineFilestoServer
        '
        Me.button_pushOfflineFilestoServer.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.button_pushOfflineFilestoServer.Image = Global.AT_EngJournal.My.Resources.Resources.document_save_3
        Me.button_pushOfflineFilestoServer.Label = "Push Offline Files to Server"
        Me.button_pushOfflineFilestoServer.Name = "button_pushOfflineFilestoServer"
        Me.button_pushOfflineFilestoServer.ShowImage = True
        '
        'button_recordEmail
        '
        Me.button_recordEmail.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge
        Me.button_recordEmail.Image = Global.AT_EngJournal.My.Resources.Resources.mail_receive
        Me.button_recordEmail.Label = "Record Email"
        Me.button_recordEmail.Name = "button_recordEmail"
        Me.button_recordEmail.ShowImage = True
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
        Me.Box1.ResumeLayout(False)
        Me.Box1.PerformLayout()
        Me.Group3.ResumeLayout(False)
        Me.Group3.PerformLayout()
        Me.Box2.ResumeLayout(False)
        Me.Box2.PerformLayout()
        Me.Group5.ResumeLayout(False)
        Me.Group5.PerformLayout()
        Me.Group2.ResumeLayout(False)
        Me.Group2.PerformLayout()
        Me.Group4.ResumeLayout(False)
        Me.Group4.PerformLayout()
        Me.Box3.ResumeLayout(False)
        Me.Box3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Tab1 As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button1 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group2 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Button2 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Button3 As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group3 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents button_pushOfflineFilestoServer As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents button_OfflineFilesCount As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents Group4 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Label1 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Box1 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents label_con As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents label_connection As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Box2 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents label As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Label2 As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents label_offlinefilecount As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Box3 As Microsoft.Office.Tools.Ribbon.RibbonBox
    Friend WithEvents label_version As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents Group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents button_recordEmail As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property ASL_Ribbon() As ASL_Ribbon
        Get
            Return Me.GetRibbon(Of ASL_Ribbon)()
        End Get
    End Property
End Class

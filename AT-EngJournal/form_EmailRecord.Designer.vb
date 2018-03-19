<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_EmailRecord
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.button_record = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.button_skip = New System.Windows.Forms.Button()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.button_record)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Location = New System.Drawing.Point(8, 10)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(506, 89)
        Me.Panel1.TabIndex = 8
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = Global.AT_EngJournal.My.Resources.Resources.face_laugh_2
        Me.PictureBox1.Location = New System.Drawing.Point(56, 46)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(39, 30)
        Me.PictureBox1.TabIndex = 4
        Me.PictureBox1.TabStop = False
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(14, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(440, 21)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "If you wish to record this email, please enter a job number in the text box and s" &
    "elect Record."
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(190, 52)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(62, 20)
        Me.TextBox1.TabIndex = 1
        '
        'button_record
        '
        Me.button_record.Image = Global.AT_EngJournal.My.Resources.Resources.mail_receive
        Me.button_record.Location = New System.Drawing.Point(272, 40)
        Me.button_record.Name = "button_record"
        Me.button_record.Size = New System.Drawing.Size(104, 42)
        Me.button_record.TabIndex = 3
        Me.button_record.Text = "Record"
        Me.button_record.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.button_record.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(101, 55)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Project Number:"
        '
        'button_skip
        '
        Me.button_skip.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.button_skip.Image = Global.AT_EngJournal.My.Resources.Resources.dialog_close_2
        Me.button_skip.Location = New System.Drawing.Point(410, 105)
        Me.button_skip.Name = "button_skip"
        Me.button_skip.Size = New System.Drawing.Size(104, 42)
        Me.button_skip.TabIndex = 9
        Me.button_skip.Text = "Cancel"
        Me.button_skip.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.button_skip.UseVisualStyleBackColor = True
        '
        'form_EmailRecord
        '
        Me.AcceptButton = Me.button_record
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.button_skip
        Me.ClientSize = New System.Drawing.Size(522, 147)
        Me.ControlBox = False
        Me.Controls.Add(Me.button_skip)
        Me.Controls.Add(Me.Panel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "form_EmailRecord"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Record Email"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents button_skip As Windows.Forms.Button
    Friend WithEvents Panel1 As Windows.Forms.Panel
    Friend WithEvents PictureBox1 As Windows.Forms.PictureBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents button_record As Windows.Forms.Button
    Friend WithEvents Label2 As Windows.Forms.Label
End Class

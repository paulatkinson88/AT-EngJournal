<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class form_setDiscipline
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
        Me.combobox_discipline = New System.Windows.Forms.ComboBox()
        Me.button_save = New System.Windows.Forms.Button()
        Me.button_cancel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'combobox_discipline
        '
        Me.combobox_discipline.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.combobox_discipline.FormattingEnabled = True
        Me.combobox_discipline.Items.AddRange(New Object() {"Electrical", "Mechanical", "Structural"})
        Me.combobox_discipline.Location = New System.Drawing.Point(25, 23)
        Me.combobox_discipline.Name = "combobox_discipline"
        Me.combobox_discipline.Size = New System.Drawing.Size(214, 21)
        Me.combobox_discipline.TabIndex = 0
        '
        'button_save
        '
        Me.button_save.Location = New System.Drawing.Point(25, 50)
        Me.button_save.Name = "button_save"
        Me.button_save.Size = New System.Drawing.Size(104, 42)
        Me.button_save.TabIndex = 1
        Me.button_save.Text = "Save"
        Me.button_save.UseVisualStyleBackColor = True
        '
        'button_cancel
        '
        Me.button_cancel.Image = Global.AT_EngJournal.My.Resources.Resources.dialog_close_2
        Me.button_cancel.Location = New System.Drawing.Point(135, 50)
        Me.button_cancel.Name = "button_cancel"
        Me.button_cancel.Size = New System.Drawing.Size(104, 42)
        Me.button_cancel.TabIndex = 2
        Me.button_cancel.Text = "Cancel"
        Me.button_cancel.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageBeforeText
        Me.button_cancel.UseVisualStyleBackColor = True
        '
        'form_setDiscipline
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(264, 117)
        Me.ControlBox = False
        Me.Controls.Add(Me.button_cancel)
        Me.Controls.Add(Me.button_save)
        Me.Controls.Add(Me.combobox_discipline)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "form_setDiscipline"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Set Discipline"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents combobox_discipline As Windows.Forms.ComboBox
    Friend WithEvents button_save As Windows.Forms.Button
    Friend WithEvents button_cancel As Windows.Forms.Button
End Class

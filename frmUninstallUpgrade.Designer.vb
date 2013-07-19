<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmUninstallUpgrade
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
        Me.cmdUninstall = New System.Windows.Forms.Button()
        Me.lblCurrent = New System.Windows.Forms.Label()
        Me.lblCurrentVal = New System.Windows.Forms.Label()
        Me.lblNew = New System.Windows.Forms.Label()
        Me.lblNewVal = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdUninstall
        '
        Me.cmdUninstall.Location = New System.Drawing.Point(579, 262)
        Me.cmdUninstall.Name = "cmdUninstall"
        Me.cmdUninstall.Size = New System.Drawing.Size(101, 23)
        Me.cmdUninstall.TabIndex = 3
        Me.cmdUninstall.Text = "Uninstall"
        Me.cmdUninstall.UseVisualStyleBackColor = True
        '
        'lblCurrent
        '
        Me.lblCurrent.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrent.Location = New System.Drawing.Point(12, 75)
        Me.lblCurrent.Name = "lblCurrent"
        Me.lblCurrent.Size = New System.Drawing.Size(668, 46)
        Me.lblCurrent.TabIndex = 6
        '
        'lblCurrentVal
        '
        Me.lblCurrentVal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblCurrentVal.Location = New System.Drawing.Point(12, 20)
        Me.lblCurrentVal.Name = "lblCurrentVal"
        Me.lblCurrentVal.Size = New System.Drawing.Size(668, 42)
        Me.lblCurrentVal.TabIndex = 7
        '
        'lblNew
        '
        Me.lblNew.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNew.Location = New System.Drawing.Point(12, 136)
        Me.lblNew.Name = "lblNew"
        Me.lblNew.Size = New System.Drawing.Size(668, 46)
        Me.lblNew.TabIndex = 8
        '
        'lblNewVal
        '
        Me.lblNewVal.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblNewVal.Location = New System.Drawing.Point(12, 200)
        Me.lblNewVal.Name = "lblNewVal"
        Me.lblNewVal.Size = New System.Drawing.Size(668, 42)
        Me.lblNewVal.TabIndex = 9
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lblNew)
        Me.GroupBox1.Controls.Add(Me.lblNewVal)
        Me.GroupBox1.Location = New System.Drawing.Point(0, 0)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(698, 256)
        Me.GroupBox1.TabIndex = 10
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Upgrade Information"
        '
        'frmUninstallUpgrade
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(710, 292)
        Me.Controls.Add(Me.lblCurrentVal)
        Me.Controls.Add(Me.lblCurrent)
        Me.Controls.Add(Me.cmdUninstall)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmUninstallUpgrade"
        Me.Text = "Uninstall and Upgrade"
        Me.GroupBox1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdUninstall As System.Windows.Forms.Button
    Friend WithEvents lblCurrent As System.Windows.Forms.Label
    Friend WithEvents lblCurrentVal As System.Windows.Forms.Label
    Friend WithEvents lblNew As System.Windows.Forms.Label
    Friend WithEvents lblNewVal As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
End Class

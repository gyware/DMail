<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DMailForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DMailForm))
        Me.ExitBtn = New System.Windows.Forms.Button()
        Me.RunBtn = New System.Windows.Forms.Button()
        Me.GitHubBtn = New System.Windows.Forms.Button()
        Me.DMailTitle = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'ExitBtn
        '
        Me.ExitBtn.Location = New System.Drawing.Point(207, 81)
        Me.ExitBtn.Name = "ExitBtn"
        Me.ExitBtn.Size = New System.Drawing.Size(75, 23)
        Me.ExitBtn.TabIndex = 0
        Me.ExitBtn.Text = "Exit"
        Me.ExitBtn.UseVisualStyleBackColor = True
        '
        'RunBtn
        '
        Me.RunBtn.Location = New System.Drawing.Point(45, 81)
        Me.RunBtn.Name = "RunBtn"
        Me.RunBtn.Size = New System.Drawing.Size(75, 23)
        Me.RunBtn.TabIndex = 1
        Me.RunBtn.Text = "Run"
        Me.RunBtn.UseVisualStyleBackColor = True
        '
        'GitHubBtn
        '
        Me.GitHubBtn.Location = New System.Drawing.Point(126, 81)
        Me.GitHubBtn.Name = "GitHubBtn"
        Me.GitHubBtn.Size = New System.Drawing.Size(75, 23)
        Me.GitHubBtn.TabIndex = 2
        Me.GitHubBtn.Text = "GitHub"
        Me.GitHubBtn.UseVisualStyleBackColor = True
        '
        'DMailTitle
        '
        Me.DMailTitle.AutoSize = True
        Me.DMailTitle.Font = New System.Drawing.Font("Segoe UI", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.DMailTitle.Location = New System.Drawing.Point(126, 9)
        Me.DMailTitle.Name = "DMailTitle"
        Me.DMailTitle.Size = New System.Drawing.Size(88, 37)
        Me.DMailTitle.TabIndex = 3
        Me.DMailTitle.Text = "$safeprojectname$"
        '
        'DMailForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(334, 116)
        Me.Controls.Add(Me.DMailTitle)
        Me.Controls.Add(Me.GitHubBtn)
        Me.Controls.Add(Me.RunBtn)
        Me.Controls.Add(Me.ExitBtn)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "DMailForm"
        Me.Text = "$safeprojectname$ (VB.NET)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents ExitBtn As Button
    Friend WithEvents RunBtn As Button
    Friend WithEvents GitHubBtn As Button
    Friend WithEvents DMailTitle As Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TestForm
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(TestForm))
        Me.lbFiles = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.btnReg = New System.Windows.Forms.Button()
        Me.btnUnReg = New System.Windows.Forms.Button()
        Me.btnExReg = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'lbFiles
        '
        Me.lbFiles.FormattingEnabled = True
        Me.lbFiles.Location = New System.Drawing.Point(12, 49)
        Me.lbFiles.Name = "lbFiles"
        Me.lbFiles.ScrollAlwaysVisible = True
        Me.lbFiles.Size = New System.Drawing.Size(372, 147)
        Me.lbFiles.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(12, 19)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(100, 23)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Files recieved: "
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(308, 202)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "End"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'btnReg
        '
        Me.btnReg.Location = New System.Drawing.Point(227, 237)
        Me.btnReg.Name = "btnReg"
        Me.btnReg.Size = New System.Drawing.Size(75, 23)
        Me.btnReg.TabIndex = 3
        Me.btnReg.Text = "Register"
        Me.btnReg.UseVisualStyleBackColor = True
        '
        'btnUnReg
        '
        Me.btnUnReg.Location = New System.Drawing.Point(308, 237)
        Me.btnUnReg.Name = "btnUnReg"
        Me.btnUnReg.Size = New System.Drawing.Size(75, 23)
        Me.btnUnReg.TabIndex = 4
        Me.btnUnReg.Text = "UnRegister"
        Me.btnUnReg.UseVisualStyleBackColor = True
        '
        'btnExReg
        '
        Me.btnExReg.Location = New System.Drawing.Point(12, 237)
        Me.btnExReg.Name = "btnExReg"
        Me.btnExReg.Size = New System.Drawing.Size(123, 23)
        Me.btnExReg.TabIndex = 5
        Me.btnExReg.Text = "Remove bar Ext"
        Me.btnExReg.UseVisualStyleBackColor = True
        '
        'TestForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(407, 278)
        Me.Controls.Add(Me.btnExReg)
        Me.Controls.Add(Me.btnUnReg)
        Me.Controls.Add(Me.btnReg)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.lbFiles)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "TestForm"
        Me.Text = "Shell Extension Test App"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lbFiles As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents btnReg As System.Windows.Forms.Button
    Friend WithEvents btnUnReg As System.Windows.Forms.Button
    Friend WithEvents btnExReg As System.Windows.Forms.Button

End Class

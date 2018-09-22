<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MatchDB
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MatchDB))
        Me.Label1 = New System.Windows.Forms.Label
        Me.DBPathOld = New System.Windows.Forms.TextBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.DBPathNew = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.Button3 = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.DBPassOld = New System.Windows.Forms.TextBox
        Me.DbPassNew = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.HasPassDbNew = New System.Windows.Forms.CheckBox
        Me.HasPassDbOld = New System.Windows.Forms.CheckBox
        Me.CnRO = New System.Windows.Forms.CheckBox
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(168, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Old Database (having less tables):"
        '
        'DBPathOld
        '
        Me.DBPathOld.Location = New System.Drawing.Point(15, 25)
        Me.DBPathOld.Name = "DBPathOld"
        Me.DBPathOld.Size = New System.Drawing.Size(260, 20)
        Me.DBPathOld.TabIndex = 1
        Me.DBPathOld.Text = "C:\Users\Pak Thunders\Desktop\DB1.accdb"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(364, 23)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Browse"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(364, 67)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 23)
        Me.Button2.TabIndex = 5
        Me.Button2.Text = "Browse"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DBPathNew
        '
        Me.DBPathNew.Location = New System.Drawing.Point(15, 69)
        Me.DBPathNew.Name = "DBPathNew"
        Me.DBPathNew.Size = New System.Drawing.Size(260, 20)
        Me.DBPathNew.TabIndex = 4
        Me.DBPathNew.Text = "C:\Users\Pak Thunders\Desktop\DB2.accdb"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(12, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(179, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "New Database (having more tables):"
        '
        'Button3
        '
        Me.Button3.Location = New System.Drawing.Point(338, 111)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(99, 23)
        Me.Button3.TabIndex = 6
        Me.Button3.Text = "Start Upgrade"
        Me.Button3.UseVisualStyleBackColor = True
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'DBPassOld
        '
        Me.DBPassOld.Location = New System.Drawing.Point(25, 17)
        Me.DBPassOld.Name = "DBPassOld"
        Me.DBPassOld.Size = New System.Drawing.Size(48, 20)
        Me.DBPassOld.TabIndex = 7
        Me.DBPassOld.UseSystemPasswordChar = True
        '
        'DbPassNew
        '
        Me.DbPassNew.Location = New System.Drawing.Point(25, 58)
        Me.DbPassNew.Name = "DbPassNew"
        Me.DbPassNew.Size = New System.Drawing.Size(48, 20)
        Me.DbPassNew.TabIndex = 8
        Me.DbPassNew.UseSystemPasswordChar = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(22, 1)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(53, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Password"
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.HasPassDbNew)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.HasPassDbOld)
        Me.Panel1.Controls.Add(Me.DbPassNew)
        Me.Panel1.Controls.Add(Me.DBPassOld)
        Me.Panel1.Location = New System.Drawing.Point(278, 9)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(81, 86)
        Me.Panel1.TabIndex = 10
        '
        'HasPassDbNew
        '
        Me.HasPassDbNew.AutoSize = True
        Me.HasPassDbNew.Location = New System.Drawing.Point(8, 61)
        Me.HasPassDbNew.Name = "HasPassDbNew"
        Me.HasPassDbNew.Size = New System.Drawing.Size(15, 14)
        Me.HasPassDbNew.TabIndex = 15
        Me.HasPassDbNew.UseVisualStyleBackColor = True
        '
        'HasPassDbOld
        '
        Me.HasPassDbOld.AutoSize = True
        Me.HasPassDbOld.Location = New System.Drawing.Point(7, 20)
        Me.HasPassDbOld.Name = "HasPassDbOld"
        Me.HasPassDbOld.Size = New System.Drawing.Size(15, 14)
        Me.HasPassDbOld.TabIndex = 14
        Me.HasPassDbOld.UseVisualStyleBackColor = True
        '
        'CnRO
        '
        Me.CnRO.AutoSize = True
        Me.CnRO.Location = New System.Drawing.Point(15, 102)
        Me.CnRO.Name = "CnRO"
        Me.CnRO.Size = New System.Drawing.Size(147, 17)
        Me.CnRO.TabIndex = 13
        Me.CnRO.Text = "Compact and Repair Only"
        Me.CnRO.UseVisualStyleBackColor = True
        '
        'MatchDB
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(447, 144)
        Me.Controls.Add(Me.CnRO)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.DBPathNew)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.DBPathOld)
        Me.Controls.Add(Me.Label1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MatchDB"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "NCS Match DB"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents DBPathOld As System.Windows.Forms.TextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents DBPathNew As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents DBPassOld As System.Windows.Forms.TextBox
    Friend WithEvents DbPassNew As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents CnRO As System.Windows.Forms.CheckBox
    Friend WithEvents HasPassDbOld As System.Windows.Forms.CheckBox
    Friend WithEvents HasPassDbNew As System.Windows.Forms.CheckBox
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog

End Class

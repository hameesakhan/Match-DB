<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Log
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
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Log))
        Me.LogBox = New System.Windows.Forms.RichTextBox
        Me.SQLBox = New System.Windows.Forms.ListBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'LogBox
        '
        Me.LogBox.Location = New System.Drawing.Point(12, 12)
        Me.LogBox.Name = "LogBox"
        Me.LogBox.Size = New System.Drawing.Size(258, 250)
        Me.LogBox.TabIndex = 0
        Me.LogBox.Text = ""
        '
        'SQLBox
        '
        Me.SQLBox.FormattingEnabled = True
        Me.SQLBox.Location = New System.Drawing.Point(277, 11)
        Me.SQLBox.Name = "SQLBox"
        Me.SQLBox.Size = New System.Drawing.Size(870, 251)
        Me.SQLBox.TabIndex = 1
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 30000
        '
        'Log
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1159, 274)
        Me.Controls.Add(Me.SQLBox)
        Me.Controls.Add(Me.LogBox)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Log"
        Me.Text = "Log"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LogBox As System.Windows.Forms.RichTextBox
    Friend WithEvents SQLBox As System.Windows.Forms.ListBox
    Private WithEvents Timer1 As System.Windows.Forms.Timer
End Class

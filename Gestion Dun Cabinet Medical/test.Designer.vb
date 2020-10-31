<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class test
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
        Me.components = New System.ComponentModel.Container()
        Me.Timer5 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer6 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer7 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'Timer5
        '
        Me.Timer5.Enabled = True
        Me.Timer5.Interval = 1000
        '
        'Timer6
        '
        Me.Timer6.Interval = 60000
        '
        'Timer7
        '
        Me.Timer7.Interval = 60000
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 1
        '
        'Timer2
        '
        Me.Timer2.Interval = 1000
        '
        'Timer3
        '
        Me.Timer3.Interval = 60000
        '
        'test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(10, 10)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Name = "test"
        Me.Opacity = 0.0R
        Me.ShowInTaskbar = False
        Me.Text = "test"
        Me.WindowState = System.Windows.Forms.FormWindowState.Minimized
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Timer5 As System.Windows.Forms.Timer
    Friend WithEvents Timer6 As System.Windows.Forms.Timer
    Friend WithEvents Timer7 As System.Windows.Forms.Timer
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents Timer3 As System.Windows.Forms.Timer
End Class

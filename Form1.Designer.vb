<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.txtTicker = New System.Windows.Forms.TextBox()
        Me.txtNumPoints = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.butUpdate = New System.Windows.Forms.Button()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.clbChart = New System.Windows.Forms.CheckedListBox()
        Me.SuspendLayout()
        '
        'txtTicker
        '
        Me.txtTicker.Location = New System.Drawing.Point(12, 30)
        Me.txtTicker.Name = "txtTicker"
        Me.txtTicker.Size = New System.Drawing.Size(100, 20)
        Me.txtTicker.TabIndex = 0
        '
        'txtNumPoints
        '
        Me.txtNumPoints.Location = New System.Drawing.Point(138, 31)
        Me.txtNumPoints.Name = "txtNumPoints"
        Me.txtNumPoints.Size = New System.Drawing.Size(100, 20)
        Me.txtNumPoints.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(43, 10)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Ticker"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(155, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(83, 13)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Number of Days"
        '
        'butUpdate
        '
        Me.butUpdate.Location = New System.Drawing.Point(890, 27)
        Me.butUpdate.Name = "butUpdate"
        Me.butUpdate.Size = New System.Drawing.Size(75, 23)
        Me.butUpdate.TabIndex = 4
        Me.butUpdate.Text = "Update"
        Me.butUpdate.UseVisualStyleBackColor = True
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(292, 2)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(70, 13)
        Me.Label3.TabIndex = 7
        Me.Label3.Text = "Select Charts"
        '
        'Panel1
        '
        Me.Panel1.AutoScroll = True
        Me.Panel1.BackColor = System.Drawing.SystemColors.Highlight
        Me.Panel1.Location = New System.Drawing.Point(-5, 66)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(904, 426)
        Me.Panel1.TabIndex = 13
        '
        'clbChart
        '
        Me.clbChart.CheckOnClick = True
        Me.clbChart.FormattingEnabled = True
        Me.clbChart.Items.AddRange(New Object() {"Candlestick with keltner(20,2,10)", "Heikin-Ashi with keltner(20,2,10)", "Bollinger bands with SMA50,SMA100,SMA200", "20 day Donchian bands with SMA50,SMA100,SMA200", "Volume", "RSI(14)", "MACD(12,26,9)", "Weekly MACD(60,130,45)", "OBV On Balance Volume", "CMF(20) Chaikin Money Flow", "MFI(14) Money Flow Index", "Stochastic RSI(14)"})
        Me.clbChart.Location = New System.Drawing.Point(295, 17)
        Me.clbChart.Name = "clbChart"
        Me.clbChart.Size = New System.Drawing.Size(393, 49)
        Me.clbChart.TabIndex = 16
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSize = True
        Me.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.ClientSize = New System.Drawing.Size(996, 504)
        Me.Controls.Add(Me.clbChart)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.butUpdate)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.txtNumPoints)
        Me.Controls.Add(Me.txtTicker)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtTicker As TextBox
    Friend WithEvents txtNumPoints As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents butUpdate As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents clbChart As CheckedListBox
End Class

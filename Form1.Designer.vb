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
    Me.lblNumPoints = New System.Windows.Forms.Label()
    Me.butUpdate = New System.Windows.Forms.Button()
    Me.Label3 = New System.Windows.Forms.Label()
    Me.Panel1 = New System.Windows.Forms.Panel()
    Me.clbChart = New System.Windows.Forms.CheckedListBox()
    Me.gbSelect = New System.Windows.Forms.GroupBox()
    Me.rbDateRange = New System.Windows.Forms.RadioButton()
    Me.rbNumOfDays = New System.Windows.Forms.RadioButton()
    Me.lblStartDate = New System.Windows.Forms.Label()
    Me.txtStartDate = New System.Windows.Forms.TextBox()
    Me.lblEndDate = New System.Windows.Forms.Label()
    Me.txtEndDate = New System.Windows.Forms.TextBox()
    Me.gbSelect.SuspendLayout()
    Me.SuspendLayout()
    '
    'txtTicker
    '
    Me.txtTicker.Location = New System.Drawing.Point(12, 30)
    Me.txtTicker.Name = "txtTicker"
    Me.txtTicker.Size = New System.Drawing.Size(68, 20)
    Me.txtTicker.TabIndex = 0
    '
    'txtNumPoints
    '
    Me.txtNumPoints.Location = New System.Drawing.Point(259, 30)
    Me.txtNumPoints.Name = "txtNumPoints"
    Me.txtNumPoints.Size = New System.Drawing.Size(62, 20)
    Me.txtNumPoints.TabIndex = 1
    '
    'Label1
    '
    Me.Label1.AutoSize = True
    Me.Label1.Location = New System.Drawing.Point(12, 9)
    Me.Label1.Name = "Label1"
    Me.Label1.Size = New System.Drawing.Size(37, 13)
    Me.Label1.TabIndex = 2
    Me.Label1.Text = "Ticker"
    '
    'lblNumPoints
    '
    Me.lblNumPoints.AutoSize = True
    Me.lblNumPoints.Location = New System.Drawing.Point(256, 9)
    Me.lblNumPoints.Name = "lblNumPoints"
    Me.lblNumPoints.Size = New System.Drawing.Size(83, 13)
    Me.lblNumPoints.TabIndex = 3
    Me.lblNumPoints.Text = "Number of Days"
    '
    'butUpdate
    '
    Me.butUpdate.Location = New System.Drawing.Point(909, 27)
    Me.butUpdate.Name = "butUpdate"
    Me.butUpdate.Size = New System.Drawing.Size(75, 23)
    Me.butUpdate.TabIndex = 4
    Me.butUpdate.Text = "Update"
    Me.butUpdate.UseVisualStyleBackColor = True
    '
    'Label3
    '
    Me.Label3.AutoSize = True
    Me.Label3.Location = New System.Drawing.Point(488, 3)
    Me.Label3.Name = "Label3"
    Me.Label3.Size = New System.Drawing.Size(70, 13)
    Me.Label3.TabIndex = 7
    Me.Label3.Text = "Select Charts"
    '
    'Panel1
    '
    Me.Panel1.AutoScroll = True
    Me.Panel1.BackColor = System.Drawing.SystemColors.Highlight
    Me.Panel1.Location = New System.Drawing.Point(0, 70)
    Me.Panel1.Name = "Panel1"
    Me.Panel1.Size = New System.Drawing.Size(905, 440)
    Me.Panel1.TabIndex = 13
    '
    'clbChart
    '
    Me.clbChart.CheckOnClick = True
    Me.clbChart.FormattingEnabled = True
    Me.clbChart.Items.AddRange(New Object() {"Candlestick with keltner(20,2,10)", "Heikin-Ashi with keltner(20,2,10)", "Bollinger bands with SMA50,SMA100,SMA200", "20 day Donchian bands with SMA50,SMA100,SMA200", "Volume", "RSI(14)", "MACD(12,26,9)", "Weekly MACD(60,130,45)", "OBV On Balance Volume", "CMF(20) Chaikin Money Flow", "MFI(14) Money Flow Index", "Stochastic RSI(14)", "RMI(20,5) Relative Momentum Index"})
    Me.clbChart.Location = New System.Drawing.Point(491, 18)
    Me.clbChart.Name = "clbChart"
    Me.clbChart.Size = New System.Drawing.Size(393, 49)
    Me.clbChart.TabIndex = 16
    '
    'gbSelect
    '
    Me.gbSelect.Controls.Add(Me.rbDateRange)
    Me.gbSelect.Controls.Add(Me.rbNumOfDays)
    Me.gbSelect.Location = New System.Drawing.Point(96, 3)
    Me.gbSelect.Name = "gbSelect"
    Me.gbSelect.Size = New System.Drawing.Size(131, 61)
    Me.gbSelect.TabIndex = 17
    Me.gbSelect.TabStop = False
    Me.gbSelect.Text = "Select Dates Using"
    '
    'rbDateRange
    '
    Me.rbDateRange.AutoSize = True
    Me.rbDateRange.Location = New System.Drawing.Point(6, 34)
    Me.rbDateRange.Name = "rbDateRange"
    Me.rbDateRange.Size = New System.Drawing.Size(121, 17)
    Me.rbDateRange.TabIndex = 1
    Me.rbDateRange.TabStop = True
    Me.rbDateRange.Text = "Start and End Dates"
    Me.rbDateRange.UseVisualStyleBackColor = True
    '
    'rbNumOfDays
    '
    Me.rbNumOfDays.AutoSize = True
    Me.rbNumOfDays.Location = New System.Drawing.Point(6, 15)
    Me.rbNumOfDays.Name = "rbNumOfDays"
    Me.rbNumOfDays.Size = New System.Drawing.Size(101, 17)
    Me.rbNumOfDays.TabIndex = 0
    Me.rbNumOfDays.TabStop = True
    Me.rbNumOfDays.Text = "Number of Days"
    Me.rbNumOfDays.UseVisualStyleBackColor = True
    '
    'lblStartDate
    '
    Me.lblStartDate.AutoSize = True
    Me.lblStartDate.Location = New System.Drawing.Point(257, 9)
    Me.lblStartDate.Name = "lblStartDate"
    Me.lblStartDate.Size = New System.Drawing.Size(92, 13)
    Me.lblStartDate.TabIndex = 19
    Me.lblStartDate.Text = "Start Date M/D/Y"
    '
    'txtStartDate
    '
    Me.txtStartDate.Location = New System.Drawing.Point(260, 30)
    Me.txtStartDate.Name = "txtStartDate"
    Me.txtStartDate.Size = New System.Drawing.Size(79, 20)
    Me.txtStartDate.TabIndex = 18
    '
    'lblEndDate
    '
    Me.lblEndDate.AutoSize = True
    Me.lblEndDate.Location = New System.Drawing.Point(368, 9)
    Me.lblEndDate.Name = "lblEndDate"
    Me.lblEndDate.Size = New System.Drawing.Size(89, 13)
    Me.lblEndDate.TabIndex = 21
    Me.lblEndDate.Text = "End Date M/D/Y"
    '
    'txtEndDate
    '
    Me.txtEndDate.Location = New System.Drawing.Point(371, 30)
    Me.txtEndDate.Name = "txtEndDate"
    Me.txtEndDate.Size = New System.Drawing.Size(74, 20)
    Me.txtEndDate.TabIndex = 20
    '
    'Form1
    '
    Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
    Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
    Me.AutoSize = True
    Me.BackColor = System.Drawing.SystemColors.AppWorkspace
    Me.ClientSize = New System.Drawing.Size(996, 521)
    Me.Controls.Add(Me.lblEndDate)
    Me.Controls.Add(Me.txtEndDate)
    Me.Controls.Add(Me.lblStartDate)
    Me.Controls.Add(Me.txtStartDate)
    Me.Controls.Add(Me.gbSelect)
    Me.Controls.Add(Me.clbChart)
    Me.Controls.Add(Me.Panel1)
    Me.Controls.Add(Me.Label3)
    Me.Controls.Add(Me.butUpdate)
    Me.Controls.Add(Me.lblNumPoints)
    Me.Controls.Add(Me.Label1)
    Me.Controls.Add(Me.txtNumPoints)
    Me.Controls.Add(Me.txtTicker)
    Me.ForeColor = System.Drawing.SystemColors.ControlText
    Me.Name = "Form1"
    Me.Text = "Form1"
    Me.gbSelect.ResumeLayout(False)
    Me.gbSelect.PerformLayout()
    Me.ResumeLayout(False)
    Me.PerformLayout()

  End Sub

  Friend WithEvents txtTicker As TextBox
    Friend WithEvents txtNumPoints As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents lblNumPoints As Label
    Friend WithEvents butUpdate As Button
    Friend WithEvents Label3 As Label
    Friend WithEvents Panel1 As Panel
    Friend WithEvents clbChart As CheckedListBox
  Friend WithEvents gbSelect As GroupBox
  Friend WithEvents rbDateRange As RadioButton
  Friend WithEvents rbNumOfDays As RadioButton
  Friend WithEvents lblStartDate As Label
  Friend WithEvents txtStartDate As TextBox
  Friend WithEvents lblEndDate As Label
  Friend WithEvents txtEndDate As TextBox
End Class

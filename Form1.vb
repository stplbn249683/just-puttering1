﻿Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Imports Skender.Stock.Indicators
Imports System.Net.Http.Headers

Public Class Form1
  Public num_charts%, chart_sizes$()
  Public charts() As Chart, updated_successfully As Boolean
  Dim tt As ToolTip = Nothing
  Dim tl As Point = Point.Empty
  Dim clbChartHt%

  Private Sub butUpdate_Click(sender As Object, e As EventArgs) Handles butUpdate.Click
    Dim error1%
    updated_successfully = False
    Panel1.Controls.Clear()
    Panel1.Hide()
    num_charts = 0
    Me.Cursor = Cursors.WaitCursor
    Dim ticker$, ticker2$, num_of_days%, dates_selected_using%, num_checked%
    Dim start_date, end_date As Date
    Dim count1% = 0, count2% = 0, count1_t2% = 0, count2_t2% = 0

    ticker = txtTicker.Text.Trim
    ticker2 = txtTicker2.Text.Trim
    If rbNumOfDays.Checked Then
      dates_selected_using = 0
      num_of_days = CInt(txtNumPoints.Text)
      If num_of_days < 2 Then
        Me.Cursor = Cursors.Default
        Exit Sub
      End If
    Else
      dates_selected_using = 1
      If IsDate(txtStartDate.Text) = False Then
        MessageBox.Show("Invalid Start Date")
        Me.Cursor = Cursors.Default
        Exit Sub
      End If
      If IsDate(txtEndDate.Text) = False Then
        MessageBox.Show("Invalid End Date")
        Me.Cursor = Cursors.Default
        Exit Sub
      End If

      start_date = CDate(txtStartDate.Text)
      end_date = CDate(txtEndDate.Text)
      If end_date <= start_date Then
        MessageBox.Show("Invalid Date Range")
        Me.Cursor = Cursors.Default
        Exit Sub
      End If

      error1 = GetCounts(ticker, UserInput.data_source, start_date, end_date, count1, count2)
      If error1 < 0 Or count2 - count1 < 2 Then
        MessageBox.Show("Error getting data for ticker symbol " & ticker)
        Me.Cursor = Cursors.Default
        Exit Sub
      End If


      If ticker2.Length > 0 Then
        error1 = GetCounts(ticker2, UserInput.data_source, start_date, end_date, count1_t2, count2_t2)
        If error1 < 0 Or count2_t2 - count1_t2 < 2 Then
          MessageBox.Show("Error getting data for ticker symbol " & ticker)
          Me.Cursor = Cursors.Default
          Exit Sub
        End If
      End If
    End If

    num_checked = clbChart.CheckedItems.Count
    num_charts = num_checked
    If ticker2.Length > 0 Then num_charts += 1
    If num_charts <= 0 Then
      Me.Cursor = Cursors.Default
      Exit Sub
    End If

    ReDim charts(0 To num_charts - 1)
    For i = 0 To num_charts - 1
      charts(i) = New Chart
    Next
    Dim chart_desc$()
    ReDim chart_desc(0 To num_charts - 1)
    For i = 0 To num_charts - 1
      If ticker2.Length > 0 And i = num_charts - 1 Then
        chart_desc(num_charts - 1) = "ratio " & ticker & "/" & ticker2
      Else
        chart_desc(i) = clbChart.CheckedItems.Item(i).ToString
      End If
    Next

    error1 = UpdateChart(ticker, ticker2, dates_selected_using, num_of_days, count1, count2, count1_t2, count2_t2,
      UserInput.data_source, num_charts, chart_desc)
    Me.Cursor = Cursors.Default
    If error1 < 0 Then Exit Sub
    updated_successfully = True ' Do this last so that the Chart_MouseMove event handler is not using invalid data
    Panel1.Show()

    UserInput.ticker = ticker
    UserInput.ticker2 = ticker2
    UserInput.num_check_box_indices = num_checked
    ReDim UserInput.check_box_indices(0 To num_checked - 1)
    For i = 0 To num_checked - 1
      UserInput.check_box_indices(i) = clbChart.CheckedIndices.Item(i)
    Next
    UserInput.dates_selected_using = dates_selected_using
    If dates_selected_using = 0 Then
      UserInput.num_of_days = num_of_days
    Else
      UserInput.start_date = start_date.ToShortDateString
      UserInput.end_date = end_date.ToShortDateString
    End If

    Dim AppPath$, sFileName$
    AppPath$ = Application.StartupPath
    sFileName = AppPath$ & "\StockChart.ini"
    error1 = SaveDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error saving file " & sFileName)
  End Sub

  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    Dim AppPath$, error1%, sFileName$
    clbChartHt = clbChart.Height
    updated_successfully = False
    num_charts = 0
    InitializeDefaults()
    AppPath$ = Application.StartupPath
    sFileName = AppPath$ & "\DataSource.ini"
    error1 = ReadDataSource(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
    sFileName = AppPath$ & "\StockChart.ini"
    error1 = ReadDefaults(sFileName)
    If error1 < 0 Then MessageBox.Show("Error reading file " & sFileName)
    With UserInput
      txtTicker.Text = .ticker
      txtTicker2.Text = .ticker2
      txtNumPoints.Text = .num_of_days
      If .num_check_box_indices > 0 Then
        For i = 0 To .num_check_box_indices - 1
          If .check_box_indices(i) <= clbChart.Items.Count - 1 Then
            clbChart.SetItemChecked(.check_box_indices(i), True)
          End If
        Next
      End If
      If .dates_selected_using <= 0 Then
        rbNumOfDays.Checked = True
      Else
        rbDateRange.Checked = True
      End If
      txtStartDate.Text = .start_date
      txtEndDate.Text = .end_date
    End With

    lblNumPoints.Visible = False
    txtNumPoints.Visible = False
    lblStartDate.Visible = False
    txtStartDate.Visible = False
    lblEndDate.Visible = False
    txtEndDate.Visible = False
    If rbNumOfDays.Checked = True Then
      lblNumPoints.Visible = True
      txtNumPoints.Visible = True
    Else
      lblStartDate.Visible = True
      txtStartDate.Visible = True
      lblEndDate.Visible = True
      txtEndDate.Visible = True
    End If

    Dim screenHeight As Integer = My.Computer.Screen.WorkingArea.Height
    Dim screenWidth As Integer = My.Computer.Screen.WorkingArea.Width
    Me.Location = New Point(screenWidth / 10, screenHeight / 10)
    Me.Height = screenHeight * 80 / 100
    Me.Width = screenWidth * 80 / 100
    Panel1.Hide()
  End Sub

  Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
    If updated_successfully = False Then Exit Sub
    SetControlSizes()
  End Sub
  Sub SetControlSizes()
    Dim vertical_location%, panel_chart_ht%, panel_chart_wdth%
    Panel1.AutoScrollPosition = New Point(0, 0)
    Panel1.Location = New Point(Me.ClientSize.Width / 100, txtTicker2.Location.Y + txtTicker2.Size.Height + 10)
    Panel1.Height = 98 * (Me.ClientSize.Height - Panel1.Location.Y) / 100
    'Panel1.Location = New Point(Me.Width / 50, Me.Height / 12)
    'Panel1.Height = Me.Height * 85 / 100
    Panel1.Width = Me.ClientSize.Width * 98 / 100

    Dim i%
    panel_chart_ht = 100 * Panel1.Height / 100
    panel_chart_wdth = 98 * Panel1.Width / 100
    If num_charts <= 0 Then Exit Sub
    If num_charts = 1 Then
      If chart_sizes(0) = "large" Then
        charts(0).Height = panel_chart_ht
        charts(0).Width = panel_chart_wdth
        charts(0).Location = New Point(0, 0)
      Else
        charts(0).Height = panel_chart_ht * 25 / 100
        charts(0).Width = panel_chart_wdth
        charts(0).Location = New Point(0, 0)
      End If
    ElseIf num_charts = 2 Then
      vertical_location = 0

      If chart_sizes(0) = "large" Then
        charts(0).Height = panel_chart_ht * 75 / 100
        charts(0).Width = panel_chart_wdth
        charts(0).Location = New Point(0, vertical_location)
        If chart_sizes(1) = "large" Then charts(0).Height = panel_chart_ht * 50 / 100
      Else
        charts(0).Height = panel_chart_ht * 25 / 100
        charts(0).Width = panel_chart_wdth
        charts(0).Location = New Point(0, vertical_location)
      End If
      vertical_location = vertical_location + charts(0).Height

      If chart_sizes(1) = "large" Then
        charts(1).Height = panel_chart_ht * 75 / 100
        charts(1).Width = panel_chart_wdth
        charts(1).Location = New Point(0, vertical_location)
        If chart_sizes(0) = "large" Then charts(1).Height = panel_chart_ht * 50 / 100
      Else
        charts(1).Height = panel_chart_ht * 25 / 100
        charts(1).Width = panel_chart_wdth
        charts(1).Location = New Point(0, vertical_location)
      End If
    Else
      vertical_location = 0
      For i = 0 To num_charts - 1
        If chart_sizes(i) = "large" Then
          charts(i).Height = panel_chart_ht * 50 / 100
          charts(i).Width = panel_chart_wdth
          charts(i).Location = New Point(0, vertical_location)
        Else
          charts(i).Height = panel_chart_ht * 25 / 100
          charts(i).Width = panel_chart_wdth
          charts(i).Location = New Point(0, vertical_location)
        End If
        vertical_location = vertical_location + charts(i).Height
      Next
    End If
  End Sub
  Function UpdateChart%(ticker$, ticker2$, dates_selected_using%, num_of_days%, count1%, count2%, count1_t2%, count2_t2%,
    data_source$, num_charts%, chart_desc$())
    UpdateChart = -1
    Dim error1%, i%, row1%, row2%, row1_t2%, row2_t2%
    Dim num_for_chart%, num_for_chart_t2%, max_num_points%, num_from_db%, num_from_db_t2%
    Dim quotes As IEnumerable(Of Quote)
    Dim quotes_t2 As IEnumerable(Of Quote)
    quotes_t2 = Nothing

    num_for_chart_t2 = 0
    If dates_selected_using = 0 Then
      num_for_chart = num_of_days
      max_num_points = num_for_chart + 720 'add some points so that errors have time to die out
      quotes = GetQuotes(max_num_points, ticker, data_source).Validate()
      If ticker2.Length > 0 Then
        num_for_chart_t2 = num_of_days
        max_num_points = num_for_chart_t2 + 720 'add some points so that errors have time to die out
        quotes_t2 = GetQuotes(max_num_points, ticker2, data_source).Validate()
      End If
    Else
      num_for_chart = count2 - count1
      row1 = count1 - 720
      If row1 < 1 Then row1 = 1
      row2 = count2
      quotes = GetQuotesForRange(row1, row2, ticker, data_source).Validate()
      If count2_t2 - count1_t2 > 0 Then
        num_for_chart_t2 = count2_t2 - count1_t2
        row1_t2 = count1_t2 - 720
        If row1_t2 < 1 Then row1_t2 = 1
        row2_t2 = count2_t2
        quotes_t2 = GetQuotesForRange(row1_t2, row2_t2, ticker2, data_source).Validate()
      End If
    End If

    num_from_db = quotes.Count
    If num_from_db <= 0 Then
      MessageBox.Show("ticker symbol " & ticker & " not in database")
      Exit Function
    End If

    If num_from_db <= 25 Then
      MessageBox.Show("Not enough points for calculations for ticker symbol " & ticker)
      Exit Function
    ElseIf num_for_chart > num_from_db Then
      num_for_chart = num_from_db
      MessageBox.Show("Reducing points for ticker symbol " & ticker & " -- may reduce accuracy or date range")
    End If

    If num_for_chart_t2 > 0 Then
      num_from_db_t2 = quotes_t2.Count

      If num_from_db_t2 <= 0 Then
        MessageBox.Show("ticker symbol " & ticker2 & " not in database")
        Exit Function
      End If

      If quotes.Last.Date <> quotes_t2.Last.Date Then
        MessageBox.Show("Dates for ticker symbols " & ticker & " and " & ticker2 & " do not match")
        Exit Function
      End If

      If num_from_db_t2 <= 25 Then
        MessageBox.Show("Not enough points for calculations for ticker symbol " & ticker2)
        Exit Function
      ElseIf num_for_chart_t2 > num_from_db_t2 Then
        num_for_chart_t2 = num_from_db_t2
        MessageBox.Show("Reducing points for ticker symbol " & ticker2 & " -- may reduce accuracy or date range")
      End If

      If num_for_chart < num_for_chart_t2 Then
        num_for_chart_t2 = num_for_chart
      End If
    End If

    Panel1.Controls.Clear()
    ReDim chart_sizes(0 To num_charts - 1)
    For i = 0 To num_charts - 1
      Dim chart_size$, chart_name$
      chart_size = "large"
      chart_name = "Chart" & i.ToString.Trim
      error1 = AddChart(chart_name, chart_desc(i), charts(i), num_for_chart, num_for_chart_t2, quotes, quotes_t2, chart_size)
      If error1 < 0 Then
        Panel1.Controls.Clear()
        MessageBox.Show("Error on chart " & chart_desc(i))
        Exit Function
      End If
      chart_sizes(i) = chart_size
    Next

    SetControlSizes()
    UpdateChart = 0
  End Function

  Function AddChart%(chart_name$, chart_desc$, new_chart As Chart, num_for_chart%, num_for_chart_t2%, quotes As List(Of Skender.Stock.Indicators.Quote),
    quotes_t2 As List(Of Skender.Stock.Indicators.Quote), ByRef chart_size$)
    AddChart = -1
    Dim error1%, num_aligned_points%
    chart_size = "small"
    Panel1.Controls.Add(new_chart)
    AddHandler new_chart.MouseMove, AddressOf Me.Chart_MouseMove
    With new_chart
      .Name = chart_name
      .ChartAreas.Clear()
      .ChartAreas.Add("ChartArea0")
      .Series.Clear()
      .Cursor = Cursors.Cross
      .ChartAreas(0).CursorX.IsUserEnabled = True
      .ChartAreas(0).CursorY.IsUserEnabled = True
      .ChartAreas(0).AxisX.LabelStyle.Format = "M/d/yyyy"
      '.ChartAreas(0).AxisX.Interval = 1
      '.ChartAreas(0).AxisX.IntervalType = DateTimeIntervalType.Months
      '.ChartAreas(0).AxisX.IntervalOffset = 1
      .ChartAreas(0).AxisY.LabelStyle.Format = "0.00"
      .ChartAreas(0).AxisY.IsStartedFromZero = False
      .ChartAreas(0).AxisX.Title = "Date"
      .Visible = True
      .Legends.Clear()
      Dim legend1 As New Legend()
      .Legends.Add(legend1) ' when a chart is created dynamically, the legend needs to be added manually
      .Titles.Clear()
      Dim title1 As New Title()
      .Titles.Add(title1)

      If chart_desc.StartsWith("Heikin-Ashi") Or chart_desc.StartsWith("Candlestick") Then
        chart_size = "large"
        Dim lstDate As New List(Of Date)
        Dim lstHigh, lstLow, lstOpen, lstClose As New List(Of Double)


        If chart_desc.StartsWith("Candlestick") Then
          Call GetQuoteLists(quotes, lstDate, lstHigh, lstLow, lstOpen, lstClose)
        Else
          Dim heikin_ashi_result = quotes.GetHeikinAshi()
          Call GetHeikinAshiLists(heikin_ashi_result, lstDate, lstHigh, lstLow, lstOpen, lstClose)
        End If

        Dim lstCenterLine, lstUpperBand, lstLowerBand As New List(Of Double)
        If chart_desc.StartsWith("Candlestick with keltner") Or chart_desc.StartsWith("Heikin-Ashi with keltner") Then
          Dim keltner_result = quotes.GetKeltner(20, 2, 10)
          lstDate.Clear()
          Call GetKeltnerLists(keltner_result, lstDate, lstCenterLine, lstUpperBand, lstLowerBand)
          error1 = ResizeLists(num_for_chart, lstDate, lstHigh, lstLow, lstOpen, lstClose, lstCenterLine, lstUpperBand, lstLowerBand)
          If error1 < 0 Then Exit Function
        End If

        Dim lstSar As New List(Of Double)
        If chart_desc.StartsWith("Candlestick with parabolic SAR") Then
          Dim parabolic_sar_result = quotes.GetParabolicSar(0.02, 0.2)
          lstDate.Clear()
          GetParabolicSarLists(parabolic_sar_result, lstDate, lstSar)
          error1 = ResizeLists(num_for_chart, lstDate, lstHigh, lstLow, lstOpen, lstClose, lstSar)
          If error1 < 0 Then Exit Function
        End If

        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Candlestick
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True ' So that the chart does not include blank space for weekends and holidays
          .Points.Clear()
          .Points.DataBindXY(lstDate, lstHigh, lstLow, lstOpen, lstClose)
          Dim s1$
          If chart_desc.StartsWith("Candlestick") Then
            s1 = "Candlestick"
            .Color = Color.Green
            .BackSecondaryColor = Color.Red
          Else
            s1 = "Heikin-Ashi"
            .Color = Color.Green
            .BackSecondaryColor = Color.Red
          End If
          .LegendText = s1
        End With

        Dim perc#, keltner_range#
        perc# = 0#
        If chart_desc.StartsWith("Candlestick with keltner") Or chart_desc.StartsWith("Heikin-Ashi with keltner") Then
          keltner_range = lstUpperBand.Last - lstLowerBand.Last
          If keltner_range > 0.001 Then
            perc = 100.0 * (lstClose.Last - lstLowerBand.Last) / keltner_range
          End If
        End If

        Dim days_rising_or_falling%
        If chart_desc.StartsWith("Candlestick") Then
          days_rising_or_falling% = DaysRisingOrFalling(40, lstOpen, lstClose, True)
        Else
          days_rising_or_falling% = DaysRisingOrFalling(40, lstOpen, lstClose, False)
        End If

        Dim s2$, gain#
        s2 = ""
        If chart_desc.StartsWith("Candlestick with keltner") Or chart_desc.StartsWith("Heikin-Ashi with keltner") Then
          s2 = s2 & "% of Keltner range = " & Format(perc, "0.00") & "   "
        End If
        s2 = s2 & "Consecutive days rising/falling = " & Format(days_rising_or_falling, "0")
        If chart_desc.StartsWith("Candlestick") Then
          gain = 0.0
          If lstClose.First > 0.0001 Then
            gain = 100.0 * (lstClose.Last - lstClose.First) / lstClose.First
            s2 = s2 & "   Gain/loss % over this time period = " & Format(gain, "0.00")
          End If
        End If

        title1.Text = s2
        If chart_desc.StartsWith("Candlestick with keltner") Or chart_desc.StartsWith("Heikin-Ashi with keltner") Then
          Dim newSeries1 As New Series()
          .Series.Add(newSeries1)
          With newSeries1
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstCenterLine)
            .LegendText = "Keltner center line"
            .Color = Color.Blue
          End With

          Dim newSeries2 As New Series()
          .Series.Add(newSeries2)
          With newSeries2
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstUpperBand)
            .LegendText = "Keltner upper band"
            .Color = Color.Blue
          End With

          Dim newSeries3 As New Series()
          .Series.Add(newSeries3)
          With newSeries3
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstLowerBand)
            .LegendText = "Keltner lower band"
            .Color = Color.Blue
          End With
        End If

        If chart_desc.StartsWith("Candlestick with parabolic SAR") Then
          Dim newSeries1 As New Series()
          .Series.Add(newSeries1)
          With newSeries1
            .BorderDashStyle = ChartDashStyle.Dot
            .BorderWidth = 2
            '.BorderColor = Color.Purple
            .ChartType = SeriesChartType.Point
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstSar)
            .LegendText = "parabolic SAR       "
            .Color = Color.Purple
          End With
        End If
      ElseIf chart_desc.StartsWith("Bollinger") Or chart_desc.StartsWith("20 day Donchian") Then
        Dim s1$ = ""
        chart_size = "large"
        Dim lstDate As New List(Of Date)
        Dim lstClose As New List(Of Double)
        Call GetQuoteCloseLists(quotes, lstDate, lstClose)

        Dim lstSma, lstCenterLine, lstUpperBand, lstLowerBand As New List(Of Double)
        If chart_desc.StartsWith("Bollinger") Then
          Dim bollinger_result = quotes.GetBollingerBands(20, 2)
          lstDate.Clear()
          Call GetBollingerLists(bollinger_result, lstDate, lstSma, lstUpperBand, lstLowerBand)
          error1 = ResizeLists(num_for_chart, lstDate, lstClose, lstSma, lstUpperBand, lstLowerBand)
          If error1 < 0 Then Exit Function
        Else
          Dim donchian_result = quotes.GetDonchian(20)
          lstDate.Clear()
          Call GetDonchianLists(donchian_result, lstDate, lstCenterLine, lstUpperBand, lstLowerBand)
          error1 = ResizeLists(num_for_chart, lstDate, lstClose, lstCenterLine, lstUpperBand, lstLowerBand)
          If error1 < 0 Then Exit Function
        End If
        num_aligned_points = lstDate.Count 'XValueIndexed = True requires the points for all series to be aligned (have the same range of dates)

        .ChartAreas(0).AxisX.MajorGrid.LineColor = Color.Gainsboro
        .ChartAreas(0).AxisY.MajorGrid.LineColor = Color.Gainsboro
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstClose)
          .LegendText = "price"
          .Color = Color.Blue
        End With

        If chart_desc.StartsWith("Bollinger") Then
          Dim perc#, bollinger_range#
          perc# = 0#
          bollinger_range = lstUpperBand.Last - lstLowerBand.Last
          If bollinger_range > 0.001 Then
            perc = 100.0 * (lstClose.Last - lstLowerBand.Last) / bollinger_range
          End If
          Dim sma20#
          sma20 = lstSma.Last

          Dim newSeries1 As New Series()
          .Series.Add(newSeries1)
          With newSeries1
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstSma)
            .LegendText = "SMA(20)"
            .Color = Color.Gray
          End With

          Dim newSeries2 As New Series()
          .Series.Add(newSeries2)
          With newSeries2
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstUpperBand)
            .LegendText = "Bollinger upper band"
            .Color = Color.Gray
          End With

          Dim newSeries3 As New Series()
          .Series.Add(newSeries3)
          With newSeries3
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstLowerBand)
            .LegendText = "Bollinger lower band"
            .Color = Color.Gray
          End With
          s1 = "% of Bollinger range = " & Format(perc, "0.00") & "  Price = " & Format(lstClose.Last, "0.00") & "  SMA(20) = " & Format(sma20, "0.00")
        Else
          Dim perc#, donchian_range#
          perc# = 0#
          donchian_range = lstUpperBand.Last - lstLowerBand.Last
          If donchian_range > 0.001 Then
            perc = 100.0 * (lstClose.Last - lstLowerBand.Last) / donchian_range
          End If

          Dim newSeries1 As New Series()
          .Series.Add(newSeries1)
          With newSeries1
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstCenterLine)
            .LegendText = "Donchian center line"
            .Color = Color.Gray
          End With

          Dim newSeries2 As New Series()
          .Series.Add(newSeries2)
          With newSeries2
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstUpperBand)
            .LegendText = "Donchian upper band"
            .Color = Color.Gray
          End With

          Dim newSeries3 As New Series()
          .Series.Add(newSeries3)
          With newSeries3
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstLowerBand)
            .LegendText = "Donchian lower band"
            .Color = Color.Gray
          End With
          s1 = "% of 20 day Donchian range = " & Format(perc, "0.00") & "  Price = " & Format(lstClose.Last, "0.00")
        End If

        Dim sma_result As IEnumerable(Of SmaResult)
        sma_result = quotes.GetSma(50)
        lstDate.Clear()
        lstSma.Clear()
        Call GetSmaLists(sma_result, lstDate, lstSma)
        error1 = ResizeLists(num_for_chart, lstDate, lstSma)

        If error1 >= 0 And lstDate.Count = num_aligned_points Then
          Dim sma50#
          sma50 = lstSma.Last

          Dim newSeries4 As New Series()
          .Series.Add(newSeries4)
          With newSeries4
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstSma)
            .LegendText = "SMA(50)"
            .Color = Color.Purple
          End With
          s1 = s1 & "  SMA(50) = " & Format(sma50, "0.00")
        End If

        sma_result = quotes.GetSma(100)
        lstDate.Clear()
        lstSma.Clear()
        Call GetSmaLists(sma_result, lstDate, lstSma)
        error1 = ResizeLists(num_for_chart, lstDate, lstSma)

        If error1 >= 0 And lstDate.Count = num_aligned_points Then
          Dim sma100#
          sma100 = lstSma.Last

          Dim newSeries5 As New Series()
          .Series.Add(newSeries5)
          With newSeries5
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstSma)
            .LegendText = "SMA(100)"
            .Color = Color.Green
          End With
          s1 = s1 & "  SMA(100) = " & Format(sma100, "0.00")
        End If

        sma_result = quotes.GetSma(200)
        lstDate.Clear()
        lstSma.Clear()
        Call GetSmaLists(sma_result, lstDate, lstSma)
        error1 = ResizeLists(num_for_chart, lstDate, lstSma)

        If error1 >= 0 And lstDate.Count = num_aligned_points Then
          Dim sma200#
          sma200 = lstSma.Last

          Dim newSeries6 As New Series()
          .Series.Add(newSeries6)
          With newSeries6
            .ChartType = SeriesChartType.Line
            .XValueType = ChartValueType.DateTime
            .IsXValueIndexed = True
            .Points.DataBindXY(lstDate, lstSma)
            .LegendText = "SMA(200)"
            .Color = Color.Red
          End With
          s1 = s1 & "  SMA(200) = " & Format(sma200, "0.00")
        End If
        title1.Text = s1
      ElseIf chart_desc.StartsWith("Volume") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim lstVolume As New List(Of Double)
        Call GetQuoteVolumeLists(quotes, lstDate, lstVolume)
        error1 = ResizeLists(num_for_chart, lstDate, lstVolume)
        If error1 < 0 Then Exit Function

        title1.Text = "Voume = " & Format(lstVolume.Last, "0.00")
        .ChartAreas(0).AxisY.LabelStyle.Format = "0.00E-00"
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Column
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstVolume)
          .IsVisibleInLegend = False
          '.LegendText = "Voume"
          .Color = Color.Blue
        End With
      ElseIf chart_desc.StartsWith("RSI(14)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim rsi_result = quotes.GetRsi(14)
        Dim lstRsi As New List(Of Double)
        Call GetRsiLists(rsi_result, lstDate, lstRsi)
        error1 = ResizeLists(num_for_chart, lstDate, lstRsi)
        If error1 < 0 Then Exit Function

        .ChartAreas(0).AxisY.Maximum = 100.0
        .ChartAreas(0).AxisY.Minimum = 0.0
        .ChartAreas(0).AxisY.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.LineColor = Color.Black
        .ChartAreas(0).AxisY.MinorGrid.Enabled = True
        .ChartAreas(0).AxisY.MinorGrid.Interval = 10.0
        .ChartAreas(0).AxisY.MinorGrid.LineColor = Color.Gainsboro
        title1.Text = "RSI(14) = " & Format(lstRsi.Last, "0.00")
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstRsi)
          .IsVisibleInLegend = False
          '.LegendText = "RSI(14)"
          .Color = Color.Blue
        End With
      ElseIf chart_desc.StartsWith("MACD(12,26,9)") Or chart_desc.StartsWith("Weekly MACD(60,130,45)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim macd_result As IEnumerable(Of MacdResult)
        If chart_desc.StartsWith("MACD(12,26,9)") Then
          macd_result = quotes.GetMacd(12, 26, 9)
        Else
          macd_result = quotes.GetMacd(60, 130, 45)
        End If
        Dim lstMacd, lstSignal, lstHistogram As New List(Of Double)
        Call GetMacdLists(macd_result, lstDate, lstMacd, lstSignal, lstHistogram)
        error1 = ResizeLists(num_for_chart, lstDate, lstMacd, lstSignal, lstHistogram)
        If error1 < 0 Then Exit Function

        Dim days_rising_or_falling%, days_rising_or_falling1%
        If chart_desc.StartsWith("MACD(12,26,9)") Then
          days_rising_or_falling = DaysRisingOrFalling1(40, lstMacd)
          days_rising_or_falling1 = DaysRisingOrFalling1(40, lstHistogram)
          title1.Text = "MACD(12,26,9) = " & Format(lstMacd.Last, "0.00") & "   MACD - Signal = " & Format((lstMacd.Last - lstSignal.Last), "0.00") &
              "  MACD consecutive days rising/falling = " & Format(days_rising_or_falling, "0") & "  Histogram consecutive days rising/falling = " & Format(days_rising_or_falling1, "0")
        Else
          days_rising_or_falling = DaysRisingOrFalling1(200, lstMacd)
          days_rising_or_falling1 = DaysRisingOrFalling1(200, lstHistogram)
          title1.Text = "Weekly MACD(60,130,45) = " & Format(lstMacd.Last, "0.00") & "   MACD - Signal = " & Format((lstMacd.Last - lstSignal.Last), "0.00") &
              "  MACD consecutive days rising/falling = " & Format(days_rising_or_falling, "0") & "  Histogram consecutive days rising/falling = " & Format(days_rising_or_falling1, "0")
        End If

        ' Dim lstHistogram As List(Of Double) = lstMacd.Zip(lstSignal, Function(x, y) x - y).ToList

        Dim newSeries2 As New Series()
        .Series.Add(newSeries2)
        With newSeries2
          .ChartType = SeriesChartType.Column
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstHistogram)
          .IsVisibleInLegend = False
          .Color = Color.LightGray
        End With

        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstMacd)
          .IsVisibleInLegend = False
          .Color = Color.Green
        End With

        Dim newSeries1 As New Series()
        .Series.Add(newSeries1)
        With newSeries1
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstSignal)
          .IsVisibleInLegend = False
          .Color = Color.Red
        End With
      ElseIf chart_desc.StartsWith("OBV") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim obv_result = quotes.GetObv()
        Dim lstObv As New List(Of Double)
        Call GetObvLists(obv_result, lstDate, lstObv)
        error1 = ResizeLists(num_for_chart, lstDate, lstObv)
        If error1 < 0 Then Exit Function

        title1.Text = "OBV = " & Format(lstObv.Last, "0.00")
        .ChartAreas(0).AxisY.LabelStyle.Format = "0.00E-00"
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstObv)
          .IsVisibleInLegend = False
          '.LegendText = "OBV"
          .Color = Color.Blue
        End With
      ElseIf chart_desc.StartsWith("CMF(20)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim cmf_result = quotes.GetCmf(20)
        Dim lstCmf As New List(Of Double)
        Call GetCmfLists(cmf_result, lstDate, lstCmf)
        error1 = ResizeLists(num_for_chart, lstDate, lstCmf)
        If error1 < 0 Then Exit Function

        title1.Text = "CMF(20) = " & Format(lstCmf.Last, "0.000")
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstCmf)
          .IsVisibleInLegend = False
          '.LegendText = "CMF(20)"
          .Color = Color.Blue
        End With
      ElseIf chart_desc.StartsWith("MFI(14)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim mfi_result = quotes.GetMfi(14)
        Dim lstMfi As New List(Of Double)
        Call GetMfiLists(mfi_result, lstDate, lstMfi)
        error1 = ResizeLists(num_for_chart, lstDate, lstMfi)
        If error1 < 0 Then Exit Function

        .ChartAreas(0).AxisY.Maximum = 100.0
        .ChartAreas(0).AxisY.Minimum = 0.0
        .ChartAreas(0).AxisY.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.LineColor = Color.Black
        .ChartAreas(0).AxisY.MinorGrid.Enabled = True
        .ChartAreas(0).AxisY.MinorGrid.Interval = 10.0
        .ChartAreas(0).AxisY.MinorGrid.LineColor = Color.Gainsboro
        title1.Text = "MFI(14) = " & Format(lstMfi.Last, "0.00")
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstMfi)
          .IsVisibleInLegend = False
          '.LegendText = "MFI(14)"
          .Color = Color.Blue
        End With
      ElseIf chart_desc.StartsWith("Stochastic RSI(14)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim stoch_rsi_result As IEnumerable(Of StochRsiResult)
        stoch_rsi_result = quotes.GetStochRsi(14, 14, 3, 1)
        Dim lstStochRsi, lstSignal As New List(Of Double)
        Call GetStochRsiLists(stoch_rsi_result, lstDate, lstStochRsi, lstSignal)
        error1 = ResizeLists(num_for_chart, lstDate, lstStochRsi, lstSignal)
        If error1 < 0 Then Exit Function

        .ChartAreas(0).AxisY.Maximum = 100.0
        .ChartAreas(0).AxisY.Minimum = 0.0
        .ChartAreas(0).AxisY.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.LineColor = Color.Black
        .ChartAreas(0).AxisY.MinorGrid.Enabled = True
        .ChartAreas(0).AxisY.MinorGrid.Interval = 10.0
        .ChartAreas(0).AxisY.MinorGrid.LineColor = Color.Gainsboro

        title1.Text = "Stochastic RSI(14) = " & Format(lstStochRsi.Last, "0.00") & "   Signal = " & Format(lstSignal.Last, "0.00") &
        "  where Signal is SMA(3) of Stochastic RSI"
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstStochRsi)
          .IsVisibleInLegend = False
          '.LegendText = "Stochastic RSI(14)"
          .Color = Color.Green
        End With

        Dim newSeries1 As New Series()
        .Series.Add(newSeries1)
        With newSeries1
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstSignal)
          .IsVisibleInLegend = False
          '.LegendText = "Signal"
          .Color = Color.Red
        End With
      ElseIf chart_desc.StartsWith("RMI(20,5)") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim lstRmi, lstSignal As New List(Of Double)
        error1 = FindRmi(quotes, 20, 5, 9, lstDate, lstRmi, lstSignal)
        If error1 < 0 Then Exit Function
        error1 = ResizeLists(num_for_chart, lstDate, lstRmi, lstSignal)
        If error1 < 0 Then Exit Function

        .ChartAreas(0).AxisY.Maximum = 100.0
        .ChartAreas(0).AxisY.Minimum = 0.0
        .ChartAreas(0).AxisY.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.Interval = 20.0
        .ChartAreas(0).AxisY.MajorGrid.LineColor = Color.Black
        .ChartAreas(0).AxisY.MinorGrid.Enabled = True
        .ChartAreas(0).AxisY.MinorGrid.Interval = 10.0
        .ChartAreas(0).AxisY.MinorGrid.LineColor = Color.Gainsboro
        title1.Text = "RMI(20,5) = " & Format(lstRmi.Last, "0.00") & "  Signal line is EMA(9) of RMI = " & Format(lstSignal.Last, "0.00")

        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstRmi)
          .IsVisibleInLegend = False
          '.LegendText = "RMI(20,5)"
          .Color = Color.Green
        End With

        Dim newSeries1 As New Series()
        .Series.Add(newSeries1)
        With newSeries1
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstSignal)
          .IsVisibleInLegend = False
          .Color = Color.Red
        End With
      ElseIf chart_desc.StartsWith("ratio") Then
        chart_size = "small"
        Dim lstDate As New List(Of Date)
        Dim lstCloseT1, lstCloseT2 As New List(Of Double)
        Call GetQuoteCloseLists(quotes, lstDate, lstCloseT1)
        lstDate.Clear()
        Call GetQuoteCloseLists(quotes_t2, lstDate, lstCloseT2)
        error1 = ResizeLists(num_for_chart_t2, lstDate, lstCloseT1, lstCloseT2)
        If error1 < 0 Then Exit Function
        Dim bResult = lstCloseT2.Exists(Function(value As Integer)
                                          Return value <= 0.01
                                        End Function)
        If bResult = True Then
          MessageBox.Show("Error in " & chart_desc & " --- denominator <= 0")
          Exit Function
        End If

        Dim scale = lstCloseT2.First / lstCloseT1.First
        Dim lstRatio As List(Of Double) = lstCloseT1.Zip(lstCloseT2, Function(x, y) scale * x / y).ToList

        title1.Text = chart_desc & " = " & Format(lstRatio.Last, "0.000")
        Dim newSeries0 As New Series()
        .Series.Add(newSeries0)
        With newSeries0
          .ChartType = SeriesChartType.Line
          .XValueType = ChartValueType.DateTime
          .IsXValueIndexed = True
          .Points.DataBindXY(lstDate, lstRatio)
          .IsVisibleInLegend = False
          '.LegendText = chan_desc
          .Color = Color.Blue
        End With
      End If
      new_chart.Update()
    End With
    AddChart = 0
  End Function

  Private Sub Chart_MouseMove(sender As System.Object, e As MouseEventArgs)
    If updated_successfully = False Then Exit Sub
    Dim s$, i%, j%, index1%
    s = CType(CType(sender, System.Windows.Forms.DataVisualization.Charting.Chart).Name, String).Trim
    If s.StartsWith("Chart") And sender.visible Then
      j = s.IndexOf("t")
      If j < 0 Then Exit Sub
      If j + 2 > s.Length Then Exit Sub
      i = CInt(s.Substring(j + 1))
      If IsNothing(tt) Then tt = New ToolTip()

      Dim ca As ChartArea = charts(i).ChartAreas(0)
      Dim r = InnerPlotPositionClientRectangle(charts(i), ca)
      If r.Contains(e.Location) Then
        ca.RecalculateAxesScale()
        Dim ax As Axis = ca.AxisX
        Dim ay As Axis = ca.AxisY
        Dim x As Double = ax.PixelPositionToValue(e.X) ' returns the point # when .IsXValueIndexed = True
        index1 = Math.Round(x) - 1
        If index1 >= 0 And index1 <= charts(i).Series(0).Points.Count - 1 Then
          Dim y As Double = ay.PixelPositionToValue(e.Y)
          Dim s1$ = DateTime.FromOADate(charts(i).Series(0).Points(index1).XValue).ToString("M/d/yyyy")
          'Dim s1$ = DateTime.FromOADate(x).ToString() does not work when .IsXValueIndexed = True
          If (e.Location <> tl) Then
            tt.SetToolTip(charts(i), String.Format("X,Y: {0}   {1:0.00}", s1, y))
            tl = e.Location
          End If
        End If
      Else
        tt.Hide(charts(i))
      End If
    End If
  End Sub

  Private Sub rbNumOfDays_CheckedChanged(sender As Object, e As EventArgs) Handles rbNumOfDays.CheckedChanged
    If rbNumOfDays.Checked = True Then
      lblNumPoints.Visible = True
      txtNumPoints.Visible = True
      lblStartDate.Visible = False
      txtStartDate.Visible = False
      lblEndDate.Visible = False
      txtEndDate.Visible = False
    Else
      lblNumPoints.Visible = False
      txtNumPoints.Visible = False
      lblStartDate.Visible = True
      txtStartDate.Visible = True
      lblEndDate.Visible = True
      txtEndDate.Visible = True
    End If
  End Sub

  Private Sub rbDateRange_CheckedChanged(sender As Object, e As EventArgs) Handles rbDateRange.CheckedChanged
    If rbDateRange.Checked = True Then
      lblNumPoints.Visible = False
      txtNumPoints.Visible = False
      lblStartDate.Visible = True
      txtStartDate.Visible = True
      lblEndDate.Visible = True
      txtEndDate.Visible = True
    Else
      lblNumPoints.Visible = True
      txtNumPoints.Visible = True
      lblStartDate.Visible = False
      txtStartDate.Visible = False
      lblEndDate.Visible = False
      txtEndDate.Visible = False
    End If
  End Sub

  Function ChartAreaClientRectangle(chart As Chart, ca As ChartArea) As RectangleF
    Dim ca_rect As RectangleF = ca.Position.ToRectangleF()
    Dim wd! = chart.ClientSize.Width / 100.0F
    Dim ht! = chart.ClientSize.Height / 100.0F
    ChartAreaClientRectangle = New RectangleF(wd * ca_rect.X, ht * ca_rect.Y, wd * ca_rect.Width, ht * ca_rect.Height)
  End Function

  Function InnerPlotPositionClientRectangle(chart As Chart, ca As ChartArea) As RectangleF

    Dim ipp_rect As RectangleF = ca.InnerPlotPosition.ToRectangleF()
    Dim cac_rect As RectangleF = ChartAreaClientRectangle(chart, ca)

    Dim wd! = cac_rect.Width / 100.0F
    Dim ht! = cac_rect.Height / 100.0F

    InnerPlotPositionClientRectangle = New RectangleF(cac_rect.X + wd * ipp_rect.X, cac_rect.Y + ht * ipp_rect.Y,
                            wd * ipp_rect.Width, ht * ipp_rect.Height)
  End Function

  Private Sub clbChart_GotFocus(sender As Object, e As System.EventArgs) _
Handles clbChart.GotFocus, clbChart.MouseEnter
    clbChart.Height = 20 + (16 * clbChart.Items.Count) ' expand the height
  End Sub

  Private Sub clbChart_LostFocus(sender As Object, e As System.EventArgs) _
Handles clbChart.LostFocus, clbChart.MouseLeave
    clbChart.Height = clbChartHt ' set the height back to it's initial size
    clbChart.SelectedIndex = -1
  End Sub
End Class
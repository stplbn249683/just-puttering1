' Modified on 1Jun22 to make a few minor changes that did not affect the operation; delete unneccessary lines, etc.
' Modified on 19Jun22 to check the quotes using quotes.Validate()
' Modifed on 20Jun22 to add "Panel1.AutoScrollPosition = New Point(0, 0)" at the begining of SetControlSizes. This
' keeps the chart controls at the top of the panel if the panel is scrolled downward when the form is resized.
' Modified on 25Jun22 to add a chart of the RMI. I used my own routine to calculate it since I could not find
' it in the Skender indicators.
' Modified on 27Jun22 to remove some unnecessary lines.
' Modified on 3Aug22 to allow the user to select the dates using the start and end date instead of the number of days.
' Last modified on 3Aug22

Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Imports Skender.Stock.Indicators
Structure INPUTTYPE
  Dim data_source$
  Dim ticker$
  Dim num_of_days$
  Dim num_check_box_indices%
  Dim check_box_indices%()
  Dim dates_selected_using%
  Dim start_date$
  Dim end_date$
End Structure

Module Module1
  Public UserInput As INPUTTYPE
  Sub GetQuoteLists(quotes As IEnumerable(Of Quote), ByRef lstDate As List(Of Date), ByRef lstHigh As List(Of Double), ByRef lstLow As List(Of Double),
                         ByRef lstOpen As List(Of Double), ByRef lstClose As List(Of Double))
    lstDate = (From x In quotes
               Select date_value = x.[Date]
               Select CDate(date_value)).ToList

    lstHigh = (From x In quotes
               Select x1 = x.High
               Select CDbl(x1)).ToList

    lstLow = (From x In quotes
              Select x1 = x.Low
              Select CDbl(x1)).ToList

    lstOpen = (From x In quotes
               Select x1 = x.Open
               Select CDbl(x1)).ToList

    lstClose = (From x In quotes
                Select x1 = x.Close
                Select CDbl(x1)).ToList
  End Sub
  Sub GetHeikinAshiLists(result As IEnumerable(Of HeikinAshiResult), ByRef lstDate As List(Of Date), ByRef lstHigh As List(Of Double), ByRef lstLow As List(Of Double),
                         ByRef lstOpen As List(Of Double), ByRef lstClose As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date]
               Select CDate(date_value)).ToList

    lstHigh = (From x In result
               Select x1 = x.High
               Select CDbl(x1)).ToList

    lstLow = (From x In result
              Select x1 = x.Low
              Select CDbl(x1)).ToList

    lstOpen = (From x In result
               Select x1 = x.Open
               Select CDbl(x1)).ToList

    lstClose = (From x In result
                Select x1 = x.Close
                Select CDbl(x1)).ToList
  End Sub

  Sub GetKeltnerLists(result As IEnumerable(Of KeltnerResult), ByRef lstDate As List(Of Date), ByRef lstCenterLine As List(Of Double), ByRef lstUpperBand As List(Of Double),
                         ByRef lstLowerBand As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Centerline
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstCenterLine = (From x In result
                     Select x1 = x.Centerline
                     Where x1 IsNot Nothing
                     Select CDbl(x1)).ToList

    lstUpperBand = (From x In result
                    Select x1 = x.UpperBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList

    lstLowerBand = (From x In result
                    Select x1 = x.LowerBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList
  End Sub


  Sub GetQuoteCloseLists(quotes As IEnumerable(Of Quote), ByRef lstDate As List(Of Date), ByRef lstClose As List(Of Double))
    lstDate = (From x In quotes
               Select date_value = x.[Date]
               Select CDate(date_value)).ToList

    lstClose = (From x In quotes
                Select x1 = x.Close
                Select CDbl(x1)).ToList
  End Sub


  Sub GetQuoteVolumeLists(quotes As IEnumerable(Of Quote), ByRef lstDate As List(Of Date), ByRef lstVolume As List(Of Double))
    lstDate = (From x In quotes
               Select date_value = x.[Date]
               Select CDate(date_value)).ToList

    lstVolume = (From x In quotes
                 Select x1 = x.Volume
                 Select CDbl(x1)).ToList
  End Sub
  Sub GetBollingerLists(result As IEnumerable(Of BollingerBandsResult), ByRef lstDate As List(Of Date), ByRef lstSma As List(Of Double), ByRef lstUpperBand As List(Of Double),
                         ByRef lstLowerBand As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Sma
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstSma = (From x In result
              Select x1 = x.Sma
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList

    lstUpperBand = (From x In result
                    Select x1 = x.UpperBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList

    lstLowerBand = (From x In result
                    Select x1 = x.LowerBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList
  End Sub

  Sub GetDonchianLists(result As IEnumerable(Of DonchianResult), ByRef lstDate As List(Of Date), ByRef lstCenterLine As List(Of Double), ByRef lstUpperBand As List(Of Double),
                         ByRef lstLowerBand As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Centerline
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstCenterLine = (From x In result
                     Select x1 = x.Centerline
                     Where x1 IsNot Nothing
                     Select CDbl(x1)).ToList

    lstUpperBand = (From x In result
                    Select x1 = x.UpperBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList

    lstLowerBand = (From x In result
                    Select x1 = x.LowerBand
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList
  End Sub
  Sub GetEmaLists(result As IEnumerable(Of EmaResult), ByRef lstDate As List(Of Date), ByRef lstEma As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Ema
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstEma = (From x In result
              Select x1 = x.Ema
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub

  Sub GetSmaLists(result As IEnumerable(Of SmaResult), ByRef lstDate As List(Of Date), ByRef lstSma As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Sma
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstSma = (From x In result
              Select x1 = x.Sma
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub

  Sub GetRsiLists(result As IEnumerable(Of RsiResult), ByRef lstDate As List(Of Date), ByRef lstRsi As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Rsi
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstRsi = (From x In result
              Select x1 = x.Rsi
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub

  Sub GetMacdLists(result As IEnumerable(Of MacdResult), ByRef lstDate As List(Of Date), ByRef lstMacd As List(Of Double), ByRef lstSignal As List(Of Double),
    ByRef lstHistogram As List(Of Double))

    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Macd
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstMacd = (From x In result
               Select x1 = x.Macd
               Where x1 IsNot Nothing
               Select CDbl(x1)).ToList

    lstSignal = (From x In result
                 Select x1 = x.Signal
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList

    lstHistogram = (From x In result
                    Select x1 = x.Histogram
                    Where x1 IsNot Nothing
                    Select CDbl(x1)).ToList
  End Sub
  Sub GetObvLists(result As IEnumerable(Of ObvResult), ByRef lstDate As List(Of Date), ByRef lstObv As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date]
               Select CDate(date_value)).ToList

    lstObv = (From x In result
              Select x1 = x.Obv
              Select CDbl(x1)).ToList
  End Sub

  Sub GetCmfLists(result As IEnumerable(Of CmfResult), ByRef lstDate As List(Of Date), ByRef lstCmf As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Cmf
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstCmf = (From x In result
              Select x1 = x.Cmf
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub

  Sub GetMfiLists(result As IEnumerable(Of MfiResult), ByRef lstDate As List(Of Date), ByRef lstMfi As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.Mfi
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstMfi = (From x In result
              Select x1 = x.Mfi
              Where x1 IsNot Nothing
              Select CDbl(x1)).ToList
  End Sub

  Sub GetStochRsiLists(result As IEnumerable(Of StochRsiResult), ByRef lstDate As List(Of Date), ByRef lstStochRsi As List(Of Double), ByRef lstSignal As List(Of Double))
    lstDate = (From x In result
               Select date_value = x.[Date], x1 = x.StochRsi
               Where x1 IsNot Nothing
               Select CDate(date_value)).ToList

    lstStochRsi = (From x In result
                   Select x1 = x.StochRsi
                   Where x1 IsNot Nothing
                   Select CDbl(x1)).ToList

    lstSignal = (From x In result
                 Select x1 = x.Signal
                 Where x1 IsNot Nothing
                 Select CDbl(x1)).ToList
  End Sub
  Function GetCounts%(ticker$, data_source$, start_date As Date, end_date As Date, ByRef count1%, ByRef count2%)
    GetCounts = -1
    count1 = 0
    count2 = 0
    Dim date1$ = start_date.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
    Dim date2$ = end_date.ToString("yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!

    cn.ConnectionString = market_price_db
    cn.Open()
    Dim cmd As New SqlCommand, dr As SqlDataReader
    cmd.Connection = cn

    Try
      cmd.CommandText = "Select COUNT(Date) As c1 From dbo.market_price Where (Ticker = '" & ticker & "') And (Date < " & date1 & ")"
      dr = cmd.ExecuteReader
      If dr.HasRows Then
        dr.Read()
        count1 = CInt(dr("c1"))
        dr.Close()
      End If
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Function
    End Try

    Try
      cmd.CommandText = "Select COUNT(Date) As c2 From dbo.market_price Where (Ticker = '" & ticker & "') And (Date <= " & date2 & ")"
      dr = cmd.ExecuteReader
      If dr.HasRows Then
        dr.Read()
        count2 = CInt(dr("c2"))
        dr.Close()
      End If
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Function
    End Try
    GetCounts = 0
  End Function
  Function GetQuotes(max_num_points%, ticker$, data_source$) As List(Of Skender.Stock.Indicators.Quote)
    Dim query1$
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!

    Dim quotes As New List(Of Skender.Stock.Indicators.Quote)
    quotes.Clear()
    GetQuotes = quotes
    cn.ConnectionString = market_price_db
    ' I want the BOTTOM records of the original table in ascending order
    query1 = "Select * FROM (Select TOP " & Trim$(Str$(max_num_points)) & " * FROM market_price t1 WHERE Ticker='" & ticker & "' ORDER BY t1.Date DESC) t2 ORDER BY t2.Date ASC"
    Try
      Dim sda As New SqlDataAdapter(query1, cn)
      Dim dt As DataTable = New DataTable
      sda.Fill(dt)
      If dt.Rows.Count > 0 Then
        quotes = (From x In dt.AsEnumerable()
                  Select date1 = x.Field(Of Int32)("Date"), high1 = x.Field(Of Decimal)("High"), low1 = x.Field(Of Decimal)("Low"),
                    open1 = x.Field(Of Decimal)("Open"), close1 = x.Field(Of Decimal)("Close"), volume1 = x.Field(Of Long)("Volume")
                  Select New Skender.Stock.Indicators.Quote With {
                  .[Date] = ConvertDate(date1),
                  .High = CDbl(high1),
                  .Low = CDbl(low1),
                  .Open = CDbl(open1),
                  .Close = CDbl(close1),
                  .Volume = CDbl(volume1)}
                 ).ToList
      End If
    Catch e As Exception
      MessageBox.Show(e.Message)
      Exit Function
    End Try
    GetQuotes = quotes
  End Function
  Function GetQuotesForRange(row1%, row2%, ticker$, data_source$) As List(Of Skender.Stock.Indicators.Quote)
    Dim query1$
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!

    Dim quotes As New List(Of Skender.Stock.Indicators.Quote)
    quotes.Clear()
    GetQuotesForRange = quotes
    cn.ConnectionString = market_price_db
    query1 = "Select * from(Select ROW_NUMBER() OVER (order by Date ASC) as 'row_num',* From dbo.market_price Where (Ticker ='" & ticker & "')) as tbl " &
    "Where tbl.row_num >= " & row1.ToString & " and tbl.row_num <= " & row2.ToString
    Try
      Dim sda As New SqlDataAdapter(query1, cn)
      Dim dt As DataTable = New DataTable
      sda.Fill(dt)
      If dt.Rows.Count > 0 Then
        quotes = (From x In dt.AsEnumerable()
                  Select date1 = x.Field(Of Int32)("Date"), high1 = x.Field(Of Decimal)("High"), low1 = x.Field(Of Decimal)("Low"),
                    open1 = x.Field(Of Decimal)("Open"), close1 = x.Field(Of Decimal)("Close"), volume1 = x.Field(Of Long)("Volume")
                  Select New Skender.Stock.Indicators.Quote With {
                  .[Date] = ConvertDate(date1),
                  .High = CDbl(high1),
                  .Low = CDbl(low1),
                  .Open = CDbl(open1),
                  .Close = CDbl(close1),
                  .Volume = CDbl(volume1)}
                 ).ToList
      End If
    Catch e As Exception
      MessageBox.Show(e.Message)
      Exit Function
    End Try
    GetQuotesForRange = quotes
  End Function

  Function ConvertDate$(date1&)
    Dim s1$
    s1 = date1.ToString.Trim
    Dim parsedDate = DateTime.ParseExact(s1, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture)
    Dim formattedDate = parsedDate.ToString("M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture)
    ConvertDate = formattedDate
  End Function

  'Function ConvertDate$(date1&)
  '  Dim year1$, month1$, day1$
  '  ConvertDate = ""
  '  If Len(date1) = 8 Then
  '    year1 = Mid$(date1, 1, 4)
  '    month1 = Mid$(date1, 5, 2)
  '    If Mid$(month1, 1, 1) = "0" Then month1 = Mid$(month1, 2, 1)
  '    day1 = Mid$(date1, 7, 2)
  '    If Mid$(day1, 1, 1) = "0" Then day1 = Mid$(day1, 2, 1)
  '    ConvertDate = month1 & "/" & day1 & "/" & year1
  '  End If
  'End Function
  Function GetQuotes1(max_num_points%, ticker$, data_source$) As List(Of Skender.Stock.Indicators.Quote)
    ' This returns the same data as GetQuotes but uses a SqlDataReader instead of a DataTable
    Dim date1&, n%
    Dim date2 As Date
    Dim market_price_db$ = "Data Source=" & data_source & ";Initial Catalog=market_data;Integrated Security=True;"
    Dim cn As New SqlConnection() ' Don't put this statement in a try block; it throws an exception!!!

    Dim quotes As New List(Of Skender.Stock.Indicators.Quote)
    quotes.Clear()
    GetQuotes1 = quotes
    cn.ConnectionString = market_price_db
    cn.Open()
    Dim cmd As New SqlCommand, dr As SqlDataReader
    cmd.Connection = cn

    ' for the skender calculations, I want the records in ascending order
    Try
      ' I want the BOTTOM records of the original table; I also want them in ascending order
      cmd.CommandText = "Select * FROM (Select TOP " & Trim$(Str$(max_num_points)) & " * FROM market_price t1 WHERE Ticker='" & ticker & "' ORDER BY t1.Date DESC) t2 ORDER BY t2.Date ASC"
      'cmd.CommandText = "Select Top " & Trim$(Str$(max_num_points)) & " * from market_price where Ticker='" & ticker & "' Order By Date DESC"
      dr = cmd.ExecuteReader
      n = 0
      If dr.HasRows Then
        Dim year1$, month1$, day1$, s2$
        While dr.Read()
          date1 = CLng(dr("Date"))
          s2 = ""
          If Len(date1) = 8 Then
            year1 = Mid$(date1, 1, 4)
            month1 = Mid$(date1, 5, 2)
            If Mid$(month1, 1, 1) = "0" Then month1 = Mid$(month1, 2, 1)
            day1 = Mid$(date1, 7, 2)
            If Mid$(day1, 1, 1) = "0" Then day1 = Mid$(day1, 2, 1)
            date2 = month1 & "/" & day1 & "/" & year1
          End If

          Dim value1 As New Skender.Stock.Indicators.Quote With {
          .[Date] = date2,
          .High = CDbl(dr("High")),
          .Low = CDbl(dr("Low")),
          .Open = CDbl(dr("Open")),
          .Close = CDbl(dr("Close")),
          .Volume = CDbl(dr("Volume"))
          }

          quotes.Add(value1)
          n += 1
          If n >= max_num_points Then Exit While
        End While
        dr.Close()
      End If
    Catch e As Exception
      cmd.Dispose()
      cn.Close()
      MessageBox.Show(e.Message)
      Exit Function
    End Try
    GetQuotes1 = quotes
  End Function

  Function FindRmi%(quotes As IEnumerable(Of Quote), n%, m%, ns%, ByRef lstDate As List(Of Date), ByRef lstRmi As List(Of Double), ByRef lstSignal As List(Of Double))
    ' n%...the number of periods used for the smoothing
    ' m%...the interval (expressed as number of periods) used to find the change in price
    ' ns%...the number of periods used for the EMA signal line calculation

    FindRmi = -1
    Dim i%, L%
    Dim up#(), down#(), rmi#(), diff#

    Dim lstDate1 As New List(Of Date)
    Dim lstClose As New List(Of Double)
    Call GetQuoteCloseLists(quotes, lstDate1, lstClose)
    Dim x As Array = lstClose.ToArray
    L = x.Length
    If L < 2 * n + ns + m Then
      MessageBox.Show("FindRmi --- Not enough points")
      Exit Function
    End If

    ReDim up#(0 To L - m - 1), down#(0 To L - m - 1)
    For i = m To L - 1
      up(i - m) = 0#
      down(i - m) = 0#
      diff = x(i) - x(i - m)
      If diff > 0 Then
        up(i - m) = diff
      ElseIf (diff < 0) Then
        down(i - m) = Math.Abs(diff)
      End If
    Next

    ' initialize with the SMA
    Dim sma_up#, sma_down#, smooth_up#, smooth_down#, dN#, multiplier#, dNs#, sma#, ema#
    sma_up = 0#
    sma_down = 0#

    dN = CDbl(n)
    For i = 0 To n - 1
      sma_up = sma_up + up(i)
      sma_down = sma_down + down(i)
    Next
    smooth_up = sma_up / dN
    smooth_down = sma_down / dN

    ' continue with the smooth

    ReDim rmi#(0 To L - m - n - 1)
    multiplier = 1.0# / dN
    For i = n + m To L - 1
      smooth_up = up(i - m) * multiplier + (1.0# - multiplier) * smooth_up
      smooth_down = down(i - m) * multiplier + (1.0# - multiplier) * smooth_down
      If smooth_up + smooth_down <= 0.0000000001 Then
        rmi(i - n - m) = 50.0#
      Else
        rmi(i - n - m) = 100.0# - 100.0# * smooth_down / (smooth_down + smooth_up)
      End If
    Next

    'signal line
    dNs = CDbl(ns)
    sma = 0.0
    For i = 0 To ns - 1
      sma = sma + rmi(i)
    Next
    ema = sma / dNs

    lstDate.Clear()
    lstRmi.Clear()
    lstSignal.Clear()
    multiplier = 2.0# / (dNs + 1.0)
    For i = n + ns + m To L - 1
      ema = rmi(i - n - m) * multiplier + (1.0# - multiplier) * ema
      lstDate.Add(lstDate1.ElementAt(i))
      lstRmi.Add(rmi(i - n - m))
      lstSignal.Add(ema)
    Next
    FindRmi = 0
  End Function
  Sub InitializeDefaults()
    With UserInput
      .data_source = "your data source name goes here"
      .ticker = ""
      .num_of_days = "0"
      .num_check_box_indices = 0
      .dates_selected_using = 0
      .start_date = ""
      .end_date = ""
    End With
  End Sub
  Function ReadDefaults(ByVal sFileName$)
    ReadDefaults = 0
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadDefaults = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        .num_check_box_indices = 0
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then Exit Function
          line = line.Trim
          If line.Length <= 0 Then Exit Function
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "ticker"
              .ticker = items(1).Trim
            Case "num_of_days"
              .num_of_days = items(1).Trim
            Case "check_box_indices"
              Dim line_items = line.Split(CType(",", Char()), 2)
              Dim indices = line_items(1).Split(",")
              .num_check_box_indices = indices.Count
              ReDim .check_box_indices(0 To .num_check_box_indices - 1)
              For i = 0 To .num_check_box_indices - 1
                .check_box_indices(i) = CInt(indices(i))
              Next
            Case "dates_selected_using"
              .dates_selected_using = CInt(items(1).Trim)
            Case "start_date"
              .start_date = items(1).Trim
            Case "end_date"
              .end_date = items(1).Trim
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadDefaults = -2
      Exit Function
    End Try
    ReadDefaults = 0
  End Function

  Function ReadDataSource(ByVal sFileName$)
    ReadDataSource = 0
    If (Dir(sFileName$) = "") Then Exit Function
    If Not File.Exists(sFileName) Then Exit Function
    Dim line$
    ReadDataSource = -1
    line = ""

    Try
      Dim reader As New StreamReader(sFileName)
      With UserInput
        While (Not reader.EndOfStream)
          line = reader.ReadLine()
          If (line Is Nothing) Then Exit Function
          line = line.Trim
          If line.Length <= 0 Then Exit Function
          Dim items = line.Split(",")
          Select Case (Trim$(items(0)))
            Case "data_source"
              .data_source = items(1).Trim
          End Select
        End While
      End With
      reader.Close()
    Catch e As Exception
      MessageBox.Show("Error in file " & sFileName & ": " & e.Message)
      ReadDataSource = -2
      Exit Function
    End Try
    ReadDataSource = 0
  End Function
  Function SaveDefaults(ByVal sFileName$)
    SaveDefaults = -1
    If File.Exists(sFileName) Then File.Delete(sFileName)
    Try
      Dim writer1 As New StreamWriter(sFileName)
      With UserInput
        writer1.WriteLine("ticker," & .ticker.Trim)
        writer1.WriteLine("num_of_days," & .num_of_days.Trim)
        If .num_check_box_indices > 0 Then
          Dim s$, i%
          s = "check_box_indices"
          For i = 0 To .num_check_box_indices - 1
            s = s & "," & .check_box_indices(i).ToString.Trim
          Next
          writer1.WriteLine(s)
        End If
        writer1.WriteLine("dates_selected_using," & .dates_selected_using.ToString.Trim)
        writer1.WriteLine("start_date," & .start_date.Trim)
        writer1.WriteLine("end_date," & .end_date.Trim)
      End With
      writer1.Close()
    Catch e As Exception
      MessageBox.Show("Error writing file " & sFileName & ": " & e.Message)
      SaveDefaults = -2
      Exit Function
    End Try
    SaveDefaults = 0
  End Function
  Function ResizeLists(num_for_chart%, ByRef lstDate As List(Of Date), ByRef list0 As List(Of Double), ByRef Optional list1 As List(Of Double) = Nothing,
                         ByRef Optional list2 As List(Of Double) = Nothing, ByRef Optional list3 As List(Of Double) = Nothing, ByRef Optional list4 As List(Of Double) = Nothing,
                           ByRef Optional list5 As List(Of Double) = Nothing, ByRef Optional list6 As List(Of Double) = Nothing, ByRef Optional list7 As List(Of Double) = Nothing)
    Dim min_num_points%
    ResizeLists = -1

    min_num_points = lstDate.Count
    If list0.Count < min_num_points Then min_num_points = list0.Count
    If Not IsNothing(list1) Then
      If list1.Count < min_num_points Then min_num_points = list1.Count
    End If
    If Not IsNothing(list2) Then
      If list2.Count < min_num_points Then min_num_points = list2.Count
    End If
    If Not IsNothing(list3) Then
      If list3.Count < min_num_points Then min_num_points = list3.Count
    End If
    If Not IsNothing(list4) Then
      If list4.Count < min_num_points Then min_num_points = list4.Count
    End If
    If Not IsNothing(list5) Then
      If list5.Count < min_num_points Then min_num_points = list5.Count
    End If
    If Not IsNothing(list6) Then
      If list6.Count < min_num_points Then min_num_points = list6.Count
    End If
    If Not IsNothing(list7) Then
      If list7.Count < min_num_points Then min_num_points = list7.Count
    End If
    If min_num_points < 10 Then Exit Function

    ResizeListOfDate(min_num_points, num_for_chart, lstDate)
    ResizeListOfDbl(min_num_points, num_for_chart, list0)
    If Not IsNothing(list1) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list1)
    End If
    If Not IsNothing(list2) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list2)
    End If
    If Not IsNothing(list3) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list3)
    End If
    If Not IsNothing(list4) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list4)
    End If
    If Not IsNothing(list5) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list5)
    End If
    If Not IsNothing(list6) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list6)
    End If
    If Not IsNothing(list7) Then
      ResizeListOfDbl(min_num_points, num_for_chart, list7)
    End If

    ResizeLists = 0
  End Function
  Sub ResizeListOfDate(min_num_points%, num_for_chart%, ByRef lstList As List(Of Date))
    Dim num_elements%
    num_elements = lstList.Count
    If num_elements > min_num_points Then
      lstList.RemoveRange(0, num_elements - min_num_points)
      num_elements = min_num_points
    End If
    If num_elements > num_for_chart Then
      lstList.RemoveRange(0, num_elements - num_for_chart)
    End If
  End Sub
  Sub ResizeListOfDbl(min_num_points%, num_for_chart%, ByRef lstList As List(Of Double))
    Dim num_elements%
    num_elements = lstList.Count
    If num_elements > min_num_points Then
      lstList.RemoveRange(0, num_elements - min_num_points)
      num_elements = min_num_points
    End If
    If num_elements > num_for_chart Then
      lstList.RemoveRange(0, num_elements - num_for_chart)
    End If
  End Sub

  Function DaysRisingOrFalling%(n%, lstOpen As List(Of Double), lstClose As List(Of Double), bCheckPreviousClose As Boolean)
    ' n = maximum number of elements to check
    Dim count%, i%, ii%, num_in_list%
    DaysRisingOrFalling = 0
    count = 0
    num_in_list = lstClose.Count
    If num_in_list < n Then
      n = num_in_list
    End If

    If bCheckPreviousClose Then
      For i = 0 To n - 2
        ii = num_in_list - 1 - i  ' ii decreases starting with the last element
        If lstClose.ElementAt(ii) > lstClose.ElementAt(ii - 1) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf lstClose.ElementAt(ii) < lstClose.ElementAt(ii - 1) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    Else
      For i = 0 To n - 1
        ii = num_in_list - 1 - i  ' ii decreases starting with the last element
        If lstClose.ElementAt(ii) > lstOpen.ElementAt(ii) Then
          If count >= 0 Then
            count += 1
          Else
            Exit For
          End If
        ElseIf lstClose.ElementAt(ii) < lstOpen.ElementAt(ii) Then
          If count <= 0 Then
            count -= 1
          Else
            Exit For
          End If
        Else
          Exit For
        End If
      Next
    End If
    DaysRisingOrFalling = count
  End Function

  Function DaysRisingOrFalling1%(n%, list1 As List(Of Double))
    ' n = maximum number of elements to check
    Dim count%, i%, ii%, ii_previous%, num_in_list%
    DaysRisingOrFalling1 = 0
    count = 0
    num_in_list = list1.Count
    If num_in_list < n Then
      n = num_in_list
    End If

    For i = 0 To n - 2
      ii = num_in_list - 1 - i  ' ii decreases starting with the last element
      ii_previous = num_in_list - 2 - i  ' previous in time
      If list1.ElementAt(ii) > list1.ElementAt(ii_previous) Then
        If count >= 0 Then
          count += 1
        Else
          Exit For
        End If
      ElseIf list1.ElementAt(ii) < list1.ElementAt(ii_previous) Then
        If count <= 0 Then
          count -= 1
        Else
          Exit For
        End If
      Else
        Exit For
      End If
    Next
    DaysRisingOrFalling1 = count
  End Function
End Module

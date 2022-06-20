' Modified on 1Jun22 to make a few minor changes that did not affect the operation; delete unneccessary lines, etc.
' Modified on 19Jun22 to check the quotes using quotes.Validate()
' Modifed on 20Jun22 to add "Panel1.AutoScrollPosition = New Point(0, 0)" at the begining of SetControlSizes. This
' keeps the chart controls at the top of the panel if the panel is scrolled downward when the form is resized.
' Last modified on 20Jun22

Option Strict Off
Option Explicit On
Imports System.IO
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Imports Skender.Stock.Indicators
Structure INPUTTYPE
  Dim data_source$
  Dim ticker$
  Dim num_for_chart$
  Dim num_check_box_indices%
  Dim check_box_indices%()
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

  Function ConvertDate$(date1&)
    Dim year1$, month1$, day1$
    ConvertDate = ""
    If Len(date1) = 8 Then
      year1 = Mid$(date1, 1, 4)
      month1 = Mid$(date1, 5, 2)
      If Mid$(month1, 1, 1) = "0" Then month1 = Mid$(month1, 2, 1)
      day1 = Mid$(date1, 7, 2)
      If Mid$(day1, 1, 1) = "0" Then day1 = Mid$(day1, 2, 1)
      ConvertDate = month1 & "/" & day1 & "/" & year1
    End If
  End Function
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
      ' I want the BOTTOM records of the original table; I also want them in ascending orser
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
  Sub InitializeDefaults()
    With UserInput
      .data_source = "your data source name goes here"
      .ticker = ""
      .num_for_chart = "0"
      .num_check_box_indices = 0
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
            Case "num_for_chart"
              .num_for_chart = items(1).Trim
            Case "check_box_indices"
              Dim line_items = line.Split(CType(",", Char()), 2)
              Dim indices = line_items(1).Split(",")
              .num_check_box_indices = indices.Count
              ReDim .check_box_indices(0 To .num_check_box_indices - 1)
              For i = 0 To .num_check_box_indices - 1
                .check_box_indices(i) = CInt(indices(i))
              Next
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
        writer1.WriteLine("num_for_chart," & .num_for_chart.Trim)
        If .num_check_box_indices > 0 Then
          Dim s$, i%
          s = "check_box_indices"
          For i = 0 To .num_check_box_indices - 1
            s = s & "," & .check_box_indices(i).ToString.Trim
          Next
          writer1.WriteLine(s)
        End If
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

Sub calc_ticker()

    Dim lastrow, sum_row, cnt, cntrlbrk, grow As Integer
    Dim cur_price, pre_price, yr_avg, start_price, gper(2) As Double
    Dim svol, gvol As LongLong
    Dim pre_ticker, cur_ticker, gticker(3) As String

    For Each ws In Worksheets
    
        'Initializing values
        gper(0) = 0
        gper(1) = 0
        gvol = 0
        'Counting number of rows for the worksheet...
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        'Summary start column -- counting columns plus 1 for summary data output...
        colcnt = ws.Cells(1, Columns.Count).End(xlToLeft).Column + 2
        
        'MsgBox ("Last Row: " & lastrow & "  No. of columns: " & colcnt)
        'resp = MsgBox("Want to Continue?", vbOKCancel)
        
        'If resp = 1 Then
                For r = 2 To lastrow
                        cnt = cnt + 1
                        'If cnt > 5 Then
                            'Exit For
                        'End If
                        cur_ticker = ws.Cells(r, 1).Value
                        cntrlbrk = StrComp(cur_ticker, pre_ticker)
                        If cntrlbrk = 0 Then
                                'MsgBox (" Inside Equal condition....pre ticker: " & pre_ticker & " current ticker:" & cur_ticker)
                                cur_price = ws.Cells(r, 6).Value
                                yr_avg = yr_avg + (cur_price - pre_price)
                                pre_price = cur_price
                                svol = svol + ws.Cells(r, 7).Value
                        Else
                                'MsgBox ("Inside <> condition.....pre ticker: " & pre_ticker & " current ticker:" & cur_ticker)
                                If r = 2 Then
                                        'Print column heading....
                                        sum_row = 1
                                        ws.Cells(sum_row, colcnt).Value = "Ticker"
                                        ws.Cells(sum_row, colcnt + 1).Value = "Yearly Change"
                                        ws.Cells(sum_row, colcnt + 2).Value = "Percent Change"
                                        ws.Cells(sum_row, colcnt + 3).Value = "Total Sock Volume"
                                        'Print column heading for hard_solution...
                                        ws.Cells(sum_row, colcnt + 7).Value = "Ticker"
                                        ws.Cells(sum_row, colcnt + 8).Value = "Value"
                                        sum_row = sum_row + 1
                                        grow = sum_row
                                Else
                                        ws.Cells(sum_row, colcnt).Value = pre_ticker
                                        ws.Cells(sum_row, colcnt + 1).Value = Format(yr_avg, "Standard")
                                         If yr_avg < 0 Then
                                                ws.Cells(sum_row, colcnt + 1).Interior.ColorIndex = 3
                                        ElseIf yr_avg > 0 Then
                                                ws.Cells(sum_row, colcnt + 1).Interior.ColorIndex = 4
                                        End If
                                        If start_price > 0 Then
                                                yr_per = (pre_price - start_price) / start_price
                                        Else
                                                yr_per = 0
                                        End If
                                        'for challenge part....
                                        If yr_per >= 0 And gper(0) < yr_per Then
                                            gper(0) = yr_per
                                            gticker(0) = pre_ticker
                                        ElseIf gper(1) > yr_per Then
                                            gper(1) = yr_per
                                            gticker(1) = pre_ticker
                                        End If
                                        '   format(yr_per, "Percent")
                                        ws.Cells(sum_row, colcnt + 2).Value = Format(yr_per, "0.00%")
                                        ws.Cells(sum_row, colcnt + 3).Value = svol
                                        'for challenge part....
                                        If svol > gvol Then
                                            gvol = svol
                                            gticker(2) = pre_ticker
                                        End If
                                        sum_row = sum_row + 1
                                End If
                                yr_avg = 0
                                pre_ticker = cur_ticker
                                pre_price = ws.Cells(r, 6).Value
                                start_price = pre_price
                                svol = ws.Cells(r, 7).Value
                        End If
                        
                        If r = lastrow Then
                                'page end print the summary...
                                'MsgBox ("Inside page end print")
                                ws.Cells(sum_row, colcnt).Value = pre_ticker
                                ws.Cells(sum_row, colcnt + 1).Value = Format(yr_avg, "Standard")
                                If yr_avg < 0 Then
                                                ws.Cells(sum_row, colcnt + 1).Interior.ColorIndex = 3
                                ElseIf yr_avg > 0 Then
                                                ws.Cells(sum_row, colcnt + 1).Interior.ColorIndex = 4
                                End If
                                If start_price > 0 Then
                                                yr_per = (pre_price - start_price) / start_price
                                Else
                                                yr_per = 0
                                End If
                                        '  also use format(yr_per, "Percent")
                                ws.Cells(sum_row, colcnt + 2).Value = Format(yr_per, "0.00%")
                                ws.Cells(sum_row, colcnt + 3).Value = svol
                                ws.Cells(sum_row, colcnt + 3).ColumnWidth = Len(ws.Cells(1, colcnt + 3).Value)
                                ' End of moderate_solution task...
                                '
                                ws.Cells(grow, colcnt + 6).Value = "Greatest % Increase"
                                ws.Cells(grow, colcnt + 6).ColumnWidth = Len(ws.Cells(grow, colcnt + 6).Value) + 5
                                ws.Cells(grow, colcnt + 7).Value = gticker(0)
                                ws.Cells(grow, colcnt + 8).Value = Format(gper(0), "Percent")
                                grow = grow + 1
                                ws.Cells(grow, colcnt + 6).Value = "Greatest % Decrease"
                                ws.Cells(grow, colcnt + 7).Value = gticker(1)
                                ws.Cells(grow, colcnt + 8).Value = Format(gper(1), "Percent")
                                grow = grow + 1
                                ws.Cells(grow, colcnt + 6).Value = "Greatest Total Volume"
                                ws.Cells(grow, colcnt + 7).Value = gticker(2)
                                'ws.Cells(grow, colcnt + 8).Value = Format(gvol, "Scientific")
                                'ws.Cells(grow, colcnt + 8).Value = Format(gvol, "0.####E+00")
                                'ws.Cells(grow, colcnt + 8).Value = Format(gvol, "0.####")
                                ws.Cells(grow, colcnt + 8).NumberFormat = "0.####E+00"
                                ws.Cells(grow, colcnt + 8).Value = Format(gvol, "0.####E+00")
                         End If
                  Next r
         'Else
            'MsgBox ("Exiting program...")
            'Exit For
        'End If

    Next ws

End Sub

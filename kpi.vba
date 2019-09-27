Private Sub ComboBox2_Change()

End Sub

Private Sub CommandButton1_Click()
    Dim pocz, kon As Date
    pocz = MonthView1.Value
    kon = MonthView2.Value
    najstarsza = Application.Min(Worksheets("Dane").Range("d:d"))
    najnowsza = Application.Max(Worksheets("Dane").Range("d:d"))
    
    If ComboBox1.Value = "" Then
        MsgBox "Nie wybrałeś grupy towarowej"
    ElseIf CLng(pocz) < CLng(najstarsza) Then
        MsgBox "Data początkowa poza zakresem danych"
    ElseIf CLng(kon) > CLng(najnowsza) Then
        MsgBox "Data końcowa poza zakresem danych"
    ElseIf CLng(kon) < CLng(pocz) Then
        MsgBox "Data końcowa jest starsza od początkowej"
    ElseIf ComboBox1.Value = "Wszystkie GT" Then
        Label1.Caption = Round(WorksheetFunction.Sum(Worksheets("Dane").Range("CY:CY")) / WorksheetFunction.Sum(Worksheets("Dane").Range("L:L")), 2)
    Else
        Label1.Caption = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("CY:CY"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(pocz), Worksheets("Dane").Range("d:d"), "<=" & CLng(kon)) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("L:L"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(pocz), Worksheets("Dane").Range("d:d"), "<=" & CLng(kon)), 2)
        Label9.Caption = ComboBox1.Value
        Label8.Caption = "Termin platnosci dla:"
        Label10.Caption = "wynosi:"
        Label11.Caption = "od:"
        Label12.Caption = pocz
        Label13.Caption = "do:"
        Label14.Caption = kon
        Label15.Caption = "dni"
    End If
End Sub


Private Sub CommandButton3_Click()
    Dim pocz, kon As Date
    Dim i, j, k, l, m, n, o As Long
    Dim oChObj As ChartObject, rngSourceData As Range, ws As Worksheet
    Dim s As String
    Dim nazwa As Boolean

    najstarsza = Application.Min(Worksheets("Dane").Range("d:d"))
    najnowsza = Application.Max(Worksheets("Dane").Range("d:d"))
    nazwa = True
    pocz = MonthView1.Value
    kon = MonthView2.Value
    ' Worksheets("Wykresy").Cells(1, 1).Value = DateDiff("m", pocz, kon)
    
    '----- dane do wybierania zakresu dat ------- |
    
    k = Year(pocz)
    j = Month(pocz)
    l = Day(pocz)
    
    
    m = Year(kon)
    n = Month(kon)
    o = Day(kon)
    
    '------ walidacja danych -------|
    
    If ComboBox1.Value = "" Then
        MsgBox "Nie wybrałeś grupy towarowej"
    ElseIf CLng(pocz) < CLng(najstarsza) Then
        MsgBox "Data początkowa poza zakresem danych"
    ElseIf CLng(kon) > CLng(najnowsza) Then
        MsgBox "Data końcowa poza zakresem danych"
    ElseIf CLng(kon) < CLng(pocz) Then
        MsgBox "Data końcowa jest starsza od początkowej"
    
    '--------------------- Terminy platnosci - krajowi ------------------
    
    ElseIf ComboBox2.Value = "Terminy płatności - krajowi" Then
        If ComboBox1.Value = "Wszystkie GT" Then ' --------- dla wszystkich GT ----------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminy płatności"
            Worksheets(s).Cells(1, 3).Value = "Obrót [kPLN]"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cy:cy"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / 1000, 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
        '--------------- rysowanie wykresu ------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Dim MySeries As Series
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerBackgroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
            
        Else ' --------------- dla konkretnych GT -------------------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminy płatności"
            Worksheets(s).Cells(1, 3).Value = "Obrót [kPLN]"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cy:cy"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "=" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / 1000, 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
        '------------------------ rysowanie wykresu ----------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        End If
        
    '--------------------- Terminy platnosci - zagraniczni ------------------
    
    ElseIf ComboBox2.Value = "Terminy płatności - zagraniczni" Then
        If ComboBox1.Value = "Wszystkie GT" Then ' --------- dla wszystkich GT ----------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminy płatności"
            Worksheets(s).Cells(1, 3).Value = "Obrót [kPLN]"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cy:cy"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / 1000, 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        Else ' --------------- dla konkretnych GT -----------------------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminy płatności"
            Worksheets(s).Cells(1, 3).Value = "Obrót [kPLN]"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cy:cy"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("l:l"), Worksheets("Dane").Range("df:df"), "<>" & "POLSKA", Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / 1000, 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        End If
        
    '------------------------- Terminowość -------------------------
    
    ElseIf ComboBox2.Value = "Terminowość" Then
        If ComboBox1.Value = "Wszystkie GT" Then ' --------- dla wszystkich GT ----------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminowość [%]"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cj:cj"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2) * 100
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        Else ' --------------- dla konkretnych GT -----------------------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Terminowość [%]"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cj:cj"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2) * 100
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        End If
        
    '--------------------------------Czas dostawy ------------------------------
            
    ElseIf ComboBox2.Value = "Czas dostawy" Then
        If ComboBox1.Value = "Wszystkie GT" Then ' --------- dla wszystkich GT ----------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Czas dostawy"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cg:cg"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        Else ' --------------- dla konkretnych GT -----------------------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Czas dostawy"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("cg:cg"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("D:D"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("D:D"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        End If
        
    '------------------- Jakość ------------------------------
    
    ElseIf ComboBox2.Value = "Jakość" Then
        If ComboBox1.Value = "Wszystkie GT" Then ' --------- dla wszystkich GT ----------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Jakość [%]"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("ck:ck"), Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2) * 100
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        Else ' --------------- dla konkretnych GT -----------------------
            s = TextBox1.Text
            For Each Sheet In Worksheets
                If s = Sheet.Name Then
                    MsgBox "Wybierz inną nazwę wykresu"
                    nazwa = False
                    Exit Sub
                End If
            Next Sheet
            With ThisWorkbook
                .Sheets.Add(After:=.Sheets(.Sheets.Count)).Name = s
            End With
            Worksheets(s).Cells(1, 2).Value = "Jakość [%]"
            Worksheets(s).Cells(1, 3).Value = "Ilość linii zamówieniowych"
            For i = 1 To DateDiff("m", pocz, kon) + 1
                Worksheets(s).Cells(i + 1, 1).Value = MonthName(j) & " " & k
                Worksheets(s).Cells(i + 1, 2).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("ck:ck"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))) / WorksheetFunction.SumIfs(Worksheets("Dane").Range("dl:dl"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2) * 100
                Worksheets(s).Cells(i + 1, 3).Value = Round(WorksheetFunction.SumIfs(Worksheets("Dane").Range("ck:ck"), Worksheets("Dane").Range("bv:bv"), ComboBox1.Value, Worksheets("Dane").Range("d:d"), ">=" & CLng(DateSerial(k, j, 0) + 1), Worksheets("Dane").Range("d:d"), "<=" & CLng(DateSerial(k, j + 1, 0))), 2)
                If j = 12 Then
                    j = 1
                    k = k + 1
                Else: j = j + 1
                End If
            Next i
             '---------------- rysowanie wykresu ---------------------------
            Set ws = Sheets(s)
            lastRow = ws.Range("B" & Rows.Count).End(xlUp).Row
            lastRow1 = ws.Range("A" & Rows.Count).End(xlUp).Row
            lastRow2 = ws.Range("C" & Rows.Count).End(xlUp).Row
            Set rngSourceData = ws.Range("c1:c" & lastRow)

            For Each oChObj In ws.ChartObjects
                oChObj.Delete
            Next
    
            Set oChObj = ws.ChartObjects.Add(Left:=ws.Columns("I").Left, Width:=670, Top:=ws.Rows(2).Top, Height:=410)

            With oChObj.Chart
                .ChartType = xlColumnClustered
                .SetSourceData Source:=rngSourceData, PlotBy:=xlColumns
                .HasTitle = True
        
                With .Axes(xlCategory, xlPrimary)
                    .CategoryNames = ws.Range("A2:A" & lastRow)
                    .TickLabels.Font.Bold = True
                End With
        
                Set MySeries = .SeriesCollection.NewSeries
        
                With MySeries
                    .Type = xlLine
                    .AxisGroup = xlSecondary
                    .MarkerStyle = xlMarkerStyleDiamond
                    .MarkerSize = 7
                    .Name = ws.Range("b1")
                    .Values = ws.Range("b2:b" & lastRow)
                    .Border.ColorIndex = 46
                    .MarkerForegroundColor = RGB(255, 140, 0)
                    .MarkerForegroundColor = RGB(255, 140, 0)
                End With
        
                .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
        
                With .ChartTitle
                    .Caption = ComboBox2.Value & " " & ComboBox1.Value
                    .Font.Size = 12
                    .Font.Bold = True
                    .Font.Name = "Arial"
                    .Font.Color = RGB(0, 0, 0)
                    .Format.Fill.ForeColor.RGB = RGB(245, 245, 245)
                    .Border.Color = RGB(0, 0, 0)
                End With
            End With
            MsgBox "Wykres dodany do nowego arkusza"
        End If
    ' ----------- inne błędy -------------
    Else
        MsgBox "Fatal Error"
    End If
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Initialize()

Dim v, e

With Sheets("Dane").Range("bv2:bv100000")
    v = .Value
End With


Set dict = CreateObject("Scripting.Dictionary")
Set dict1 = CreateObject("Scripting.Dictionary")

With dict
    .comparemode = 1
    dict.Add "Wszystkie GT", 0
    For Each e In v
        If Not .exists(e) Then .Add e, Nothing
    Next
    If .Count Then Me.ComboBox1.List = Application.Transpose(.keys)

End With

With dict1
    .comparemode = 1
    dict1.Add "Terminy płatności - krajowi", 0
    dict1.Add "Terminy płatności - zagraniczni", 0
    dict1.Add "Czas dostawy", 0
    dict1.Add "Terminowość", 0
    dict1.Add "Jakość", 0
    If .Count Then Me.ComboBox2.List = Application.Transpose(.keys)

End With


End Sub


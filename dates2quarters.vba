Sub dates2quarters()

'-----------------------------------------'
sheet_name = "collection"
indicator = "release" ' first column
firstyear = 2023 ' year to start with
lastyear = 2029 ' year to end with
'-----------------------------------------'

Dim ix

For ix = 1 To Worksheets.Count
    If Sheets(ix).Name = "quarters" Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("quarters").Delete
        Application.DisplayAlerts = True
        Exit For ' exit loop for less trouble
    End If
Next

Sheets.Add after:=Sheets(Worksheets.Count)
ActiveSheet.Name = "quarters"

Sheets("quarters").Select
ActiveWindow.DisplayGridlines = False
Columns(2).ColumnWidth = 4
Columns(3).ColumnWidth = 33.5
Columns(4).ColumnWidth = 9
Columns(5).ColumnWidth = 4
Columns(6).ColumnWidth = 4
Columns(7).ColumnWidth = 11
Columns(8).ColumnWidth = 9
Columns(9).ColumnWidth = 5
Columns(10).ColumnWidth = 4
Columns(11).ColumnWidth = 10 ' K
Columns(12).ColumnWidth = 10 ' L
Columns(13).ColumnWidth = 10 ' M
Columns(14).ColumnWidth = 10 ' N
Columns(15).ColumnWidth = 4
Columns(18).ColumnWidth = 4

'Frame
Range("F6:F18").Borders(xlEdgeLeft).LineStyle = xlContinuous
Range("J6:J18").Borders(xlEdgeLeft).LineStyle = xlContinuous

Worksheets("quarters").Cells(2, 3).Value = "Timestamp: " & Date & " | " & Time & " Uhr"
Worksheets("quarters").Cells(2, 3).Font.Bold = True

Worksheets("quarters").Cells(4, 3).Value = "Indicator (term to count): " & indicator

If Worksheets(sheet_name).FilterMode Then Worksheets(sheet_name).ShowAllData ' reset filter
Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
nSchecks = Sheets(sheet_name).AutoFilter.Range.Columns(1).SpecialCells(xlCellTypeVisible).Cells.Count - 1
If Worksheets(sheet_name).FilterMode Then Worksheets(sheet_name).ShowAllData ' reset filter

Worksheets("quarters").Cells(6, 3).Value = "Indicator count: "
Worksheets("quarters").Cells(6, 4).Value = nSchecks
Worksheets("quarters").Cells(6, 3).Font.Bold = True
Worksheets("quarters").Cells(6, 4).Font.Bold = True

'-----------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------'

Worksheets("quarters").Cells(6, 7).Value = "Indicator per "
Worksheets("quarters").Cells(6, 7).Font.Bold = True
Worksheets("quarters").Cells(7, 7).Value = "year (n = " & nSchecks & ")"
Worksheets("quarters").Cells(7, 7).Font.Bold = True

Worksheets(sheet_name).EnableAutoFilter = True
Worksheets(sheet_name).Protect contents:=True, userInterfaceOnly:=True

nSonst = 0
nCCJSum = 0
zzz = 0

For i = firstyear To lastyear

startdatum = "01/01/" & i
enddatum = "12/31/" & i


    With Worksheets(sheet_name)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
        If Not .AutoFilterMode Then .Range("B1").AutoFilter
        .Range("N1").AutoFilter Field:=2, Criteria1:=">=" & startdatum, Operator:=xlAnd, Criteria2:="<=" & enddatum
    End With

    nCompJahr = Worksheets(sheet_name).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
        If nCompJahr > 0 Then
            Worksheets("quarters").Cells(9 + zzz, 7).Value = i
            Worksheets("quarters").Cells(9 + zzz, 8).Value = nCompJahr
            zzz = zzz + 1
            nCCJSum = nCCJSum + nCompJahr ' Sum for unknown
        End If
        
        ' future years are 0
        If nCompJahr = 0 Then
            Worksheets("quarters").Cells(9 + zzz, 7).Value = i
            Worksheets("quarters").Cells(9 + zzz, 8).Value = nCompJahr
            zzz = zzz + 1
        End If
Next

Worksheets(sheet_name).ShowAllData ' reset filter


' Unbekannt/Summe Schecks
nSonst = nSchecks - nCCJSum
Worksheets("quarters").Cells(9 + zzz, 7).Value = "unknown"
Worksheets("quarters").Cells(9 + zzz, 8).Value = nSonst
Worksheets("quarters").Cells(10 + zzz, 7).Value = "Sum"
    Range(Cells(10 + zzz, 7), Cells(10 + zzz, 8)).Borders(xlEdgeTop).LineStyle = xlContinuous
Worksheets("quarters").Cells(10 + zzz, 8).Value = nSchecks

'-----------------------------------------------------------------------------------------------------------------'
'-----------------------------------------------------------------------------------------------------------------'

Worksheets("quarters").Cells(6, 14).Value = "Indicator per quarter (n = " & nSchecks & ")"
Worksheets("quarters").Cells(6, 14).Font.Bold = True
Worksheets("quarters").Cells(6, 14).HorizontalAlignment = xlRight
Worksheets("quarters").Cells(7, 11).Value = "1st quarter"
Worksheets("quarters").Cells(7, 11).HorizontalAlignment = xlRight
Worksheets("quarters").Cells(7, 12).Value = "2nd quarter"
Worksheets("quarters").Cells(7, 12).HorizontalAlignment = xlRight
Worksheets("quarters").Cells(7, 13).Value = "3rd quarter"
Worksheets("quarters").Cells(7, 13).HorizontalAlignment = xlRight
Worksheets("quarters").Cells(7, 14).Value = "4th quarter"
Worksheets("quarters").Cells(7, 14).HorizontalAlignment = xlRight


Worksheets(sheet_name).EnableAutoFilter = True
Worksheets(sheet_name).Protect contents:=True, userInterfaceOnly:=True

If Worksheets(sheet_name).FilterMode Then Worksheets(sheet_name).ShowAllData ' reset filter


jahr = "2023"
startdatum = "01/01/" & jahr
enddatum = "12/31/" & jahr

Dim zz1
Dim zz2
Dim zz3
Dim zz4

nSonst = 0
nCCJSum = 0
zzz = 0
zz1 = 0
zz2 = 0
zz3 = 0
zz4 = 0

For i = 2023 To 2027


'------------------------------------------------------------> quarter 01 <---
startdatum = "01/01/" & i
enddatum = "03/31/" & i

    With Worksheets(sheet_name)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
        If Not .AutoFilterMode Then .Range("B1").AutoFilter
        .Range("N1").AutoFilter Field:=2, Criteria1:=">=" & startdatum, Operator:=xlAnd, Criteria2:="<=" & enddatum
    End With

    nCompJahr = Worksheets(sheet_name).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
        If nCompJahr > 0 Then
            Worksheets("quarters").Cells(9 + zz1, 11).Value = nCompJahr
            zz1 = zz1 + 1
            nCCJSum = nCCJSum + nCompJahr ' Sum for unknown
        End If
        
        ' future years are 0
        If nCompJahr = 0 Then
            Worksheets("quarters").Cells(9 + zz1, 11).Value = nCompJahr
            zz1 = zz1 + 1
        End If

'------------------------------------------------------------> quarter 02 <---
startdatum = "04/01/" & i
enddatum = "06/30/" & i

    With Worksheets(sheet_name)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
        If Not .AutoFilterMode Then .Range("B1").AutoFilter
        .Range("N1").AutoFilter Field:=2, Criteria1:=">=" & startdatum, Operator:=xlAnd, Criteria2:="<=" & enddatum
    End With

    nCompJahr = Worksheets(sheet_name).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
        If nCompJahr > 0 Then
            Worksheets("quarters").Cells(9 + zz2, 12).Value = nCompJahr
            zz2 = zz2 + 1
            nCCJSum = nCCJSum + nCompJahr ' Sum for unknown
        End If
        
        ' future years are 0
        If nCompJahr = 0 Then
            Worksheets("quarters").Cells(9 + zz2, 12).Value = nCompJahr
            zz2 = zz2 + 1
        End If

'------------------------------------------------------------> quarter 03 <---
startdatum = "07/01/" & i
enddatum = "09/30/" & i

    With Worksheets(sheet_name)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
        If Not .AutoFilterMode Then .Range("B1").AutoFilter
        .Range("N1").AutoFilter Field:=2, Criteria1:=">=" & startdatum, Operator:=xlAnd, Criteria2:="<=" & enddatum
    End With

    nCompJahr = Worksheets(sheet_name).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
        If nCompJahr > 0 Then
            Worksheets("quarters").Cells(9 + zz3, 13).Value = nCompJahr
            zz3 = zz3 + 1
            nCCJSum = nCCJSum + nCompJahr ' Sum for unknown
        End If
        
        ' future years are 0
        If nCompJahr = 0 Then
            Worksheets("quarters").Cells(9 + zz3, 13).Value = nCompJahr
            zz3 = zz3 + 1
        End If

'------------------------------------------------------------> quarter 04 <---
startdatum = "10/01/" & i
enddatum = "12/31/" & i

    With Worksheets(sheet_name)
        If Not .AutoFilterMode Then .Range("A1").AutoFilter
        Worksheets(sheet_name).UsedRange.AutoFilter Field:=1, Criteria1:=indicator
        If Not .AutoFilterMode Then .Range("B1").AutoFilter
        .Range("N1").AutoFilter Field:=2, Criteria1:=">=" & startdatum, Operator:=xlAnd, Criteria2:="<=" & enddatum
    End With

    nCompJahr = Worksheets(sheet_name).AutoFilter.Range.Columns(2).SpecialCells(xlCellTypeVisible).Cells.Count - 1
    
        If nCompJahr > 0 Then
            Worksheets("quarters").Cells(9 + zz4, 14).Value = nCompJahr
            zz4 = zz4 + 1
            nCCJSum = nCCJSum + nCompJahr ' Sum for unknown
        End If
        
        ' future years are 0
        If nCompJahr = 0 Then
            Worksheets("quarters").Cells(9 + zz4, 14).Value = nCompJahr
            zz4 = zz4 + 1
        End If

Next

Worksheets(sheet_name).ShowAllData ' reset filter


'-----------------------------------------'
'-----------------------------------------'

If Worksheets(sheet_name).FilterMode Then Worksheets(sheet_name).ShowAllData ' reset filter
Worksheets(sheet_name).Unprotect

End Sub


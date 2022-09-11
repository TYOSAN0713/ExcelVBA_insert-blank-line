Attribute VB_Name = "form"
Sub meibo()
  Dim gakunen
  Dim LastRow As Long
  Dim LastColumn As Long
  LastRow = Range("A1").End(xlDown).Row
  LastColumn = Cells(1, Columns.Count).End(xlToLeft).Column
  Dim class1(1 To 50, 1 To 100)
  Dim class2(1 To 50, 1 To 100)
  Dim class3(1 To 50, 1 To 100)
  Dim class4(1 To 50, 1 To 100)
  Dim class5(1 To 50, 1 To 100)
  Dim class6(1 To 50, 1 To 100)
  Dim class7(1 To 50, 1 To 100)
  Dim class8(1 To 50, 1 To 100)
  Dim cn1 As Integer, cn2 As Integer, cn3 As Integer, cn4 As Integer, cn5 As Integer, cn6 As Integer, cn7 As Integer, cn8 As Integer
  Dim i As Long, j As Long
  Dim Cnt As Long

    '  make sheet start
    Worksheets.Add
    ActiveSheet.Name = "A"
    Worksheets("A").Move , Worksheets(Worksheets.Count)
    Worksheets("Sheet1").Rows(1).Copy
    Worksheets("A").Rows(1).PasteSpecial
    Worksheets("A").Range("B:H").ColumnWidth = 20
    
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "B"
    Worksheets("B").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "C"
    Worksheets("C").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "D"
    Worksheets("D").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "E"
    Worksheets("E").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "F"
    Worksheets("F").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "G"
    Worksheets("G").Range("B:H").ColumnWidth = 20
    Worksheets("A").Copy After:=Worksheets(Worksheets.Count)
    ActiveSheet.Name = "IB"
    Worksheets("IB").Range("B:H").ColumnWidth = 20
    'make sheet end
    
    'devide start
    gakunen = Worksheets("Sheet1").Range("A1").CurrentRegion.Value

    For i = 1 To UBound(gakunen)
        Select Case gakunen(i, 9)  'change from 9 to class number
            Case "A"
                cn1 = cn1 + 1
                For j = 1 To LastColumn
                    class1(cn1, j) = gakunen(i, j)
                Next j
            Case "B"
                cn2 = cn2 + 1
                For j = 1 To LastColumn
                    class2(cn2, j) = gakunen(i, j)
                Next j
            Case "C"
                cn3 = cn3 + 1
                For j = 1 To LastColumn
                    class3(cn3, j) = gakunen(i, j)
                Next j
            Case "D"
                cn4 = cn4 + 1
                For j = 1 To LastColumn
                    class4(cn4, j) = gakunen(i, j)
                Next j
            Case "E"
                cn5 = cn5 + 1
                For j = 1 To LastColumn
                    class5(cn5, j) = gakunen(i, j)
                Next j
            Case "F"
                cn6 = cn6 + 1
                For j = 1 To LastColumn
                    class6(cn6, j) = gakunen(i, j)
                Next j
            Case "G"
                cn7 = cn7 + 1
                For j = 1 To LastColumn
                    class7(cn7, j) = gakunen(i, j)
                Next j
            Case "IB"
                cn8 = cn8 + 1
                For j = 1 To LastColumn
                    class8(cn8, j) = gakunen(i, j)
                Next j
        End Select
    Next i

    Worksheets("A").Range("A2").Resize(UBound(class1), LastColumn).Value = class1
    Worksheets("B").Range("A2").Resize(UBound(class2), LastColumn).Value = class2
    Worksheets("C").Range("A2").Resize(UBound(class3), LastColumn).Value = class3
    Worksheets("D").Range("A2").Resize(UBound(class4), LastColumn).Value = class4
    Worksheets("E").Range("A2").Resize(UBound(class5), LastColumn).Value = class5
    Worksheets("F").Range("A2").Resize(UBound(class6), LastColumn).Value = class6
    Worksheets("G").Range("A2").Resize(UBound(class7), LastColumn).Value = class7
    Worksheets("IB").Range("A2").Resize(UBound(class8), LastColumn).Value = class8
    'devide end

    'insert start
    i = 0
    LastRow = 0
    Sheets("A").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        '最終行を取得
        LastRow = Range("A1").End(xlDown).Row
        '連番チェック用
        Cnt = "1"
        '2行目からループ
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then  'change from 12 to class number
            Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If

    Sheets("B").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If

    Sheets("C").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If
 
    Sheets("D").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If

    Sheets("E").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If

    Sheets("F").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If
        
    Sheets("G").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If

    Sheets("IB").Select
    If Range("A2").Value <> 0 Then
        Range("A1").CurrentRegion.Sort _
        Key1:=Range("L1"), Order1:=xlAscending, _
        Orientation:=xlTopToBottom, Header:=xlGuess
        LastRow = Range("A1").End(xlDown).Row
        Cnt = "1"
        For i = 2 To LastRow
        If Cells(i, 12).Value <> Cnt Then
           Rows(i).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
        End If
        Cnt = Cnt + 1
        Next i
    End If
    'insert end

End Sub

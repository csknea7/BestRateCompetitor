Sub ExportFiltersOnlyToNewWorkbook()

    Dim wsSource As Worksheet
    Set wsSource = ThisWorkbook.Sheets("Sheet1") ' Kaynak sayfa

    Dim wsNew As Worksheet

    ' Eğer "Report" sayfası varsa sil (eski rapor temizlenir)
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Report").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Yeni sayfa oluştur ve "Report" olarak isimlendir
    Set wsNew = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsNew.Name = "Report"

    Dim outputRow As Long: outputRow = 1

    ' 1. satır: Başlık
    wsNew.Cells(outputRow, 1).Value = "BEST RATE COMPETITOR AVAILABLE"
    wsNew.Cells(outputRow, 1).Font.Bold = True
    wsNew.Cells(outputRow, 1).Font.Size = 12
    wsNew.Range(wsNew.Cells(outputRow, 1), wsNew.Cells(outputRow, 2)).Merge
    wsNew.Range(wsNew.Cells(outputRow, 1), wsNew.Cells(outputRow, 2)).HorizontalAlignment = xlCenter
    outputRow = outputRow + 2 ' 1 satır boşluk

    ' 2. satır: FILTERS başlığı
    wsNew.Cells(outputRow, 1).Value = "FILTERS"
    wsNew.Cells(outputRow, 1).Font.Bold = True
    outputRow = outputRow + 1

    Dim filters As Variant
    filters = Array("Market", "From", "To", "Price type", "Convert amount to", "Nationality")

    Dim i As Long, j As Long
    For i = 1 To 50
        For j = LBound(filters) To UBound(filters)
            If Trim(LCase(wsSource.Cells(i, 2).Value)) = LCase(filters(j)) Then
                wsNew.Cells(outputRow, 1).Value = filters(j)
                wsNew.Cells(outputRow, 1).Font.Bold = True
                wsNew.Cells(outputRow, 2).Value = wsSource.Cells(i, 3).Value
                outputRow = outputRow + 1
            End If
        Next j
    Next i

    ' FILTERS tablosuna kenarlık ekle
    With wsNew.Range(wsNew.Cells(outputRow - (UBound(filters) + 1), 1), wsNew.Cells(outputRow - 1, 2)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' --- Dinamik başlıkları bul ---
    Dim listHeaderRow As Long: listHeaderRow = 0
    Dim jpCol As Long, hotelCol As Long, roomCol As Long
    Dim currencyCol As Long, refundableCol As Long, baseRateCol As Long
    jpCol = 0: hotelCol = 0: roomCol = 0: currencyCol = 0: refundableCol = 0: board = 0: room = 0: baseRateCol = 0: feeCol = 0

    Dim rowCheck As Long, colCheck As Long

    ' Başlıkları bul (ilk 200 satır, 30 sütun)
    For rowCheck = 1 To 200
        For colCheck = 1 To 30
            Select Case LCase(Trim(wsSource.Cells(rowCheck, colCheck).Value))
                Case "jp code": jpCol = colCheck: listHeaderRow = rowCheck
                Case "hotel name": hotelCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "room type": roomCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "currency": currencyCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "refundable": refundableCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "board": boardCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "room": rCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "base rate": baseRateCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
                Case "fee": feeCol = colCheck: If listHeaderRow = 0 Then listHeaderRow = rowCheck
            End Select
            If jpCol * hotelCol * roomCol * currencyCol * refundableCol * boardCol * rCol * baseRateCol * feeCol > 0 Then Exit For
        Next colCheck
        If jpCol * hotelCol * roomCol * currencyCol * refundableCol * boardCol * rCol * baseRateCol * feeCol > 0 Then Exit For
    Next rowCheck

    If jpCol = 0 Or hotelCol = 0 Or roomCol = 0 Or currencyCol = 0 Or refundableCol = 0 Or boardCol = 0 Or rCol = 0 Or baseRateCol = 0 Or feeCol = 0 Then
        MsgBox "Gerekli basliklardan biri eksik: JP Code, Hotel Name, Room Type, Currency, Refundable, Board, Room,  Base Rate, Fee", vbExclamation
        Exit Sub
    End If

    ' FILTERS sonrası 2 satır boşluk bırak
    outputRow = outputRow + 2

    ' Başlıkları yaz (Report sayfasında)
    wsNew.Cells(outputRow, 1).Value = "JP Code"
    wsNew.Cells(outputRow, 2).Value = "Hotel Name"
    wsNew.Cells(outputRow, 3).Value = "Room Type"
    wsNew.Cells(outputRow, 4).Value = "Currency"
    wsNew.Cells(outputRow, 5).Value = "Refundable"
    wsNew.Cells(outputRow, 6).Value = "Board"
    wsNew.Cells(outputRow, 7).Value = "Room"
    wsNew.Rows(outputRow).Font.Bold = True
    outputRow = outputRow + 1

    ' Verileri satır satır kopyala
    Dim srcRow As Long: srcRow = listHeaderRow + 1
    Do While wsSource.Cells(srcRow, jpCol).Value <> "" Or wsSource.Cells(srcRow, hotelCol).Value <> "" _
        Or wsSource.Cells(srcRow, roomCol).Value <> "" Or wsSource.Cells(srcRow, currencyCol).Value <> "" _
        Or wsSource.Cells(srcRow, refundableCol).Value <> ""

        wsNew.Cells(outputRow, 1).Value = wsSource.Cells(srcRow, jpCol).Value
        wsNew.Cells(outputRow, 2).Value = wsSource.Cells(srcRow, hotelCol).Value
        wsNew.Cells(outputRow, 3).Value = wsSource.Cells(srcRow, roomCol).Value
        wsNew.Cells(outputRow, 4).Value = wsSource.Cells(srcRow, currencyCol).Value
        wsNew.Cells(outputRow, 5).Value = wsSource.Cells(srcRow, refundableCol).Value
            
        wsNew.Cells(outputRow, 6).Value = wsSource.Cells(srcRow, boardCol).Value
        wsNew.Cells(outputRow, 7).Value = wsSource.Cells(srcRow, rCol).Value

        wsNew.Cells(outputRow, 8).Value = wsSource.Cells(srcRow, baseRateCol).Value ' H sütununa (PRIME/BEDSOPIA) veriyi ekle
        wsNew.Cells(outputRow, 9).Value = wsSource.Cells(srcRow, feeCol).Value

        outputRow = outputRow + 1
        srcRow = srcRow + 1
    Loop

    ' Kenarlık ekle (veri tablosu)
    With wsNew.Range(wsNew.Cells(outputRow - (srcRow - listHeaderRow - 1 + 1), 1), wsNew.Cells(outputRow - 1, 11)).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    ' Sütun genişliklerini ayarla
    wsNew.Columns("A:K").AutoFit

    ' JP Code'a göre renklendirme
    Dim colorIndex As Long: colorIndex = 2 ' Başlangıç rengi
    Dim lastRow As Long: lastRow = outputRow - 1
    Dim startRow As Long: startRow = outputRow - (srcRow - listHeaderRow - 1)

    Dim currentJP As String
    Dim previousJP As String

    previousJP = wsNew.Cells(startRow, 1).Value ' İlk veri satırındaki JP Code

    For i = startRow To lastRow
        currentJP = wsNew.Cells(i, 1).Value

        If currentJP <> previousJP Then
            colorIndex = colorIndex + 1
            If colorIndex = 2 Or colorIndex = 15 Then colorIndex = colorIndex + 1 ' Renk çakışmalarını önle
            previousJP = currentJP
        End If

        wsNew.Range(wsNew.Cells(i, 1), wsNew.Cells(i, 11)).Interior.ColorIndex = colorIndex
    Next i

    ' 12. satır soft mavi arka plan
    With wsNew.Range("A12:K12").Interior
        .Color = RGB(198, 223, 249)
    End With

    ' 12. satırın F-J sütunlarına ek başlıkları yaz
    Dim extraHeaders As Variant
    extraHeaders = Array("Board", "Room", "PRIME/BEDSOPIA", "Competitor", "Competitor_", "%Needed Disc.")

    For i = 0 To UBound(extraHeaders)
        wsNew.Cells(12, 6 + i).Value = extraHeaders(i)
    Next i

    call HesaplaCompetitorCarp93
    call HesaplaNeededDiscount

    wsNew.Columns("I").EntireColumn.Hidden = True


    MsgBox "Rapor basariyla olusturuldu.", vbInformation

End Sub

Sub HesaplaCompetitorCarp93()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Report") ' Sayfa adı

    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row ' Competitor sütunu (I sütunu)

    For i = 13 To lastRow
        If IsNumeric(ws.Cells(i, 9).Value) Then ' I sütunu = 9. sütun
            ws.Cells(i, 10).Value = ws.Cells(i, 9).Value * 0.93 ' K sütununa (11) yaz
        Else
            ws.Cells(i, 10).Value = "N/A"
        End If
    Next i

End Sub

Sub HesaplaNeededDiscount()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Report") ' Sayfa adı burada "Report"

    Dim i As Long, lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row ' H sütunu: PRIME/BEDSOPIA

    For i = 13 To lastRow
        Dim primeVal As Variant
        Dim compVal As Variant
        Dim sonuc As Variant

        primeVal = ws.Cells(i, 8).Value   ' H sütunu: PRIME/BEDSOPIA
        compVal = ws.Cells(i, 10).Value   ' K sütunu: Competitor_

        If UCase(Trim(primeVal)) = "ND" Then
            ws.Cells(i, 11).Value = "ND" ' L sütunu: %Needed Disc.
        ElseIf IsNumeric(primeVal) And IsNumeric(compVal) And compVal <> 0 Then
            sonuc = (primeVal - compVal) / compVal
            ws.Cells(i, 11).Value = Format(sonuc, "0.00%")
        Else
            ws.Cells(i, 11).Value = "N/A"
        End If
    Next i


End Sub

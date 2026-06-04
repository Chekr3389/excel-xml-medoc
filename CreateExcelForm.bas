Attribute VB_Name = "CreateExcelForm"
'=============================================================================
' VBA Module: Create Excel Form Template
' Purpose: Generate Excel template as .xls file for M.E.Doc data entry
' Version: 1.0
'=============================================================================

Sub CreateTransferPricingForm()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsRef As Worksheet
    Dim filePath As String
    Dim row As Long
    Dim col As Long
    
    ' Get save path
    filePath = InputBox("Enter path to save Excel form (include filename):", _
        "Save Form", ThisWorkbook.Path & "\Form_J0147107.xls")
    If filePath = "" Then Exit Sub
    
    ' Create new workbook
    Set wb = Workbooks.Add
    Set ws = wb.Sheets(1)
    ws.Name = "Дані для конвертації"
    
    ' Add second worksheet for reference codes
    Set wsRef = wb.Sheets.Add
    wsRef.Name = "Довідка - Коди"
    
    ' ===== MAIN SHEET SETUP =====
    
    ' Set column widths
    ws.Columns("A").ColumnWidth = 8
    ws.Columns("B").ColumnWidth = 12
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 15
    ws.Columns("F").ColumnWidth = 15
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 12
    ws.Columns("I").ColumnWidth = 12
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 12
    ws.Columns("L").ColumnWidth = 12
    ws.Columns("M").ColumnWidth = 15
    ws.Columns("N").ColumnWidth = 12
    ws.Columns("O").ColumnWidth = 15
    ws.Columns("P").ColumnWidth = 12
    ws.Columns("Q").ColumnWidth = 15
    ws.Columns("R").ColumnWidth = 18
    ws.Columns("S").ColumnWidth = 12
    ws.Columns("T").ColumnWidth = 15
    ws.Columns("U").ColumnWidth = 12
    ws.Columns("V").ColumnWidth = 12
    ws.Columns("Z").ColumnWidth = 20
    
    ' Create header row (row 1)
    row = 1
    ws.Rows(row).RowHeight = 30
    ws.Range("A" & row & ":AA" & row).Interior.Color = RGB(68, 114, 196) ' Blue
    ws.Range("A" & row & ":AA" & row).Font.Color = RGB(255, 255, 255) ' White
    ws.Range("A" & row & ":AA" & row).Font.Bold = True
    ws.Range("A" & row & ":AA" & row).Font.Size = 11
    ws.Range("A" & row & ":AA" & row).HorizontalAlignment = xlCenter
    ws.Range("A" & row & ":AA" & row).VerticalAlignment = xlCenter
    ws.Range("A" & row & ":AA" & row).WrapText = True
    
    ' Add header text
    Dim headers As Variant
    headers = Array("№ з/п", "Код операції", "Тип предмета*", "Опис предмета", _
        "Код УКТ ЗЕД", "Код послуги", "Контракт/Договір*", "Код сторони*", _
        "Код країни*", "Інкотермс", "Дата операції*", "Дата до", _
        "Ціна/ум.*", "Кількість*", "Кількість контракт", "Код валюти*", _
        "Курс валюти*", "Вартість (грн)*", "Код методу", "Показник", _
        "Значення", "Сторона", "Код джерела", "Назва джерела", "Резерв 1", _
        "Резерв 2", "Примітки")
    
    For col = LBound(headers) To UBound(headers)
        ws.Cells(row, col + 1).Value = headers(col)
    Next col
    
    ' Add example rows (rows 2-3)
    Call AddExampleRow(ws, 2, "1", "035", "205", "Конвертація валюти", "07.01.02", "", _
        "№РКО-1121688/23032022/0048", "144", "826", "02", "2025-01-02", "2025-01-02", _
        689.89, 1, "2454 або 2009", "840", 41.97, 28956.41, "301", "", "", "", "601", _
        "Інформація про зіставні неконтрольовані операції", "", "", "Приклад заповнення")
    
    Call AddExampleRow(ws, 3, "2", "035", "205", "Конвертація валюти", "07.01.02", "", _
        "№РКО-1121688/23032022/0048", "144", "826", "02", "2025-01-02", "2025-01-02", _
        649, 1, "2454 або 2009", "978", 43.61, 28299.77, "301", "", "", "", "601", _
        "Інформація про зіставні неконтрольовані операції", "", "", "Приклад заповнення")
    
    ' Add empty rows (4-11)
    For row = 4 To 11
        ws.Rows(row).RowHeight = 25
        Call FormatDataRow(ws, row, CLng(row - 2))
    Next row
    
    ' ===== REFERENCE SHEET SETUP =====
    
    ' Set up reference sheet
    wsRef.Columns("A").ColumnWidth = 25
    wsRef.Columns("B").ColumnWidth = 40
    wsRef.Columns("C").ColumnWidth = 25
    wsRef.Columns("D").ColumnWidth = 40
    
    ' Add header
    row = 1
    wsRef.Range("A" & row & ":D" & row).Interior.Color = RGB(68, 114, 196)
    wsRef.Range("A" & row & ":D" & row).Font.Color = RGB(255, 255, 255)
    wsRef.Range("A" & row & ":D" & row).Font.Bold = True
    
    wsRef.Cells(row, 1).Value = "КОДИ ОПЕРАЦІЙ"
    wsRef.Cells(row, 2).Value = "Опис"
    wsRef.Cells(row, 3).Value = "КОДИ ВАЛЮТ"
    wsRef.Cells(row, 4).Value = "Опис"
    
    ' Add operation codes
    row = 2
    Call AddRefRow(wsRef, row, "035", "Послуги", "840", "USD (Доларі США)")
    Call AddRefRow(wsRef, row + 1, "201", "Товари", "978", "EUR (Євро)")
    Call AddRefRow(wsRef, row + 2, "202", "Права ІВ", "826", "GBP (Фунти)")
    Call AddRefRow(wsRef, row + 3, "203", "Гарантії", "392", "JPY (Єни)")
    Call AddRefRow(wsRef, row + 4, "204", "Кредити", "756", "CHF (Франк)")
    Call AddRefRow(wsRef, row + 5, "205", "Конвертація", "208", "DKK (Крони)")
    Call AddRefRow(wsRef, row + 6, "206", "Страхування", "380", "EUR (Італія)")
    Call AddRefRow(wsRef, row + 7, "207", "Доставка", "643", "RUB (Рублі)")
    Call AddRefRow(wsRef, row + 8, "208", "Обслуговування", "985", "PLN (Злоті)")
    Call AddRefRow(wsRef, row + 9, "209", "Інші", "124", "CAD (Доларі)")
    
    ' Add country codes section
    row = 13
    wsRef.Cells(row, 1).Value = "КОДИ КРАЇН"
    wsRef.Cells(row, 2).Value = "Опис"
    wsRef.Cells(row, 3).Value = "ІНКОТЕРМС"
    wsRef.Cells(row, 4).Value = "Опис"
    wsRef.Range("A" & row & ":D" & row).Interior.Color = RGB(68, 114, 196)
    wsRef.Range("A" & row & ":D" & row).Font.Color = RGB(255, 255, 255)
    wsRef.Range("A" & row & ":D" & row).Font.Bold = True
    
    ' Add country codes
    row = 14
    Call AddRefRow(wsRef, row, "826", "США", "02", "CIF (Вартість, страхування)")
    Call AddRefRow(wsRef, row + 1, "144", "Туреччина", "10", "DAP (Доставлено)")
    Call AddRefRow(wsRef, row + 2, "040", "Австрія", "13", "DDP (Доставлено, мито)")
    Call AddRefRow(wsRef, row + 3, "056", "Бельгія", "04", "FOB (На борту)")
    Call AddRefRow(wsRef, row + 4, "756", "Швейцарія", "05", "CFR (Вартість, фрахт)")
    Call AddRefRow(wsRef, row + 5, "250", "Франція", "01", "EXW (Ex Works)")
    Call AddRefRow(wsRef, row + 6, "276", "Німеччина", "03", "CPT (Фрахт оплачено)")
    Call AddRefRow(wsRef, row + 7, "372", "Ірландія", "06", "CIP (Фрахт, страхування)")
    Call AddRefRow(wsRef, row + 8, "380", "Італія", "07", "FCA (Франко перевізнику)")
    Call AddRefRow(wsRef, row + 9, "528", "Нідерланди", "08", "FAS (Франко борту)")
    
    ' Save workbook
    wb.SaveAs filePath, xlExcel8
    wb.Close
    
    MsgBox "Excel form successfully created at: " & filePath, vbInformation, "Success"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Sub AddExampleRow(ws As Worksheet, rowNum As Long, ParamArray values() As Variant)
    Dim col As Long
    Dim cell As Range
    
    ws.Rows(rowNum).RowHeight = 25
    
    For col = LBound(values) To UBound(values)
        Set cell = ws.Cells(rowNum, col + 1)
        cell.Value = values(col)
        
        ' Format required fields (yellow)
        If col = 0 Or col = 1 Or col = 2 Or col = 6 Or col = 7 Or col = 8 Or _
           col = 10 Or col = 13 Or col = 15 Or col = 16 Or col = 17 Then
            cell.Interior.Color = RGB(242, 242, 242)
        Else
            cell.Interior.Color = RGB(255, 255, 255)
        End If
        
        ' Apply borders
        With cell.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(217, 217, 217)
        End With
        
        cell.Font.Size = 11
        cell.VerticalAlignment = xlCenter
    Next col
End Sub

Sub FormatDataRow(ws As Worksheet, rowNum As Long, serialNum As Long)
    Dim col As Long
    Dim cell As Range
    
    ' Add serial number
    Set cell = ws.Cells(rowNum, 1)
    cell.Value = serialNum
    cell.Interior.Color = RGB(255, 242, 204) ' Yellow for required
    
    ' Format remaining cells
    For col = 2 To 27
        Set cell = ws.Cells(rowNum, col)
        
        ' Required fields - yellow background
        If col = 2 Or col = 3 Or col = 7 Or col = 8 Or col = 9 Or _
           col = 11 Or col = 14 Or col = 16 Or col = 17 Or col = 18 Then
            cell.Interior.Color = RGB(255, 242, 204)
        Else
            cell.Interior.Color = RGB(242, 242, 242)
        End If
        
        ' Apply borders
        With cell.Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(217, 217, 217)
        End With
        
        cell.Font.Size = 11
        cell.VerticalAlignment = xlCenter
    Next col
End Sub

Sub AddRefRow(ws As Worksheet, rowNum As Long, code1 As String, desc1 As String, _
    code2 As String, desc2 As String)
    
    ws.Cells(rowNum, 1).Value = code1
    ws.Cells(rowNum, 2).Value = desc1
    ws.Cells(rowNum, 3).Value = code2
    ws.Cells(rowNum, 4).Value = desc2
    
    ws.Range("A" & rowNum & ":D" & rowNum).Interior.Color = RGB(242, 242, 242)
    
    Dim col As Long
    For col = 1 To 4
        With ws.Cells(rowNum, col).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(217, 217, 217)
        End With
    Next col
End Sub

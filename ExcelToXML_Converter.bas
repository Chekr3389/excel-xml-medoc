Attribute VB_Name = "ExcelToXML_Converter"
'=============================================================================
' VBA Macro: Excel to XML Converter for M.E.Doc (Form J0147107)
' Purpose: Convert transfer pricing report data from Excel to XML format
' Version: 1.0
' Author: GitHub Copilot
' Date: 2026-06-04
'=============================================================================

Option Explicit

Sub ExportToXML()
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Dim xmlDoc As Object
    Dim rootElement As Object
    Dim headElement As Object
    Dim bodyElement As Object
    Dim tableElement As Object
    Dim rowElement As Object
    Dim rowNum As Long
    Dim lastRow As Long
    Dim i As Long
    Dim filePath As String
    Dim fileName As String
    Dim TIN As String
    Dim docDate As String
    
    ' Create XML DOM
    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.validateOnParse = False
    
    ' Get data from active sheet
    Set ws = ActiveSheet
    
    ' Input dialog for TIN (10 digits)
    TIN = InputBox("Enter TIN (10 digits):", "TIN Input", "0123456789")
    If TIN = "" Then Exit Sub
    If Len(TIN) <> 10 Or Not IsNumeric(TIN) Then
        MsgBox "Invalid TIN format! Must be 10 digits.", vbCritical
        Exit Sub
    End If
    
    ' Input dialog for file path
    filePath = InputBox("Enter file path to save XML:", "File Path", ThisWorkbook.Path & "\Report.xml")
    If filePath = "" Then Exit Sub
    
    ' Create XML declaration
    Dim xmlDecl As Object
    Set xmlDecl = xmlDoc.createProcessingInstruction("xml", "version=""1.0"" encoding=""windows-1251""")
    xmlDoc.appendChild xmlDecl
    
    ' Create root element
    Set rootElement = xmlDoc.createElement("DECLAR")
    rootElement.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
    rootElement.setAttribute "xsi:noNamespaceSchemaLocation", "J0147107.xsd"
    xmlDoc.appendChild rootElement
    
    ' Create DECLARHEAD
    Set headElement = xmlDoc.createElement("DECLARHEAD")
    rootElement.appendChild headElement
    
    Call AddElement(xmlDoc, headElement, "TIN", TIN)
    Call AddElement(xmlDoc, headElement, "C_DOC", "J01")
    Call AddElement(xmlDoc, headElement, "C_DOC_SUB", "471")
    Call AddElement(xmlDoc, headElement, "C_DOC_VER", "7")
    Call AddElement(xmlDoc, headElement, "C_DOC_TYPE", "0")
    Call AddElement(xmlDoc, headElement, "C_DOC_CNT", "1")
    Call AddElement(xmlDoc, headElement, "C_REG", "26")
    Call AddElement(xmlDoc, headElement, "C_RAJ", "90")
    Call AddElement(xmlDoc, headElement, "PERIOD_MONTH", "12")
    Call AddElement(xmlDoc, headElement, "PERIOD_TYPE", "1")
    Call AddElement(xmlDoc, headElement, "PERIOD_YEAR", Year(Now()))
    Call AddElement(xmlDoc, headElement, "C_STI_ORIG", "2600")
    Call AddElement(xmlDoc, headElement, "C_DOC_STAN", "1")
    Call AddElement(xmlDoc, headElement, "D_FILL", Format(Now(), "YYYY-MM-DD"))
    Call AddElement(xmlDoc, headElement, "SOFTWARE", "ExcelToXML v1.0")
    
    ' Create DECLARBODY
    Set bodyElement = xmlDoc.createElement("DECLARBODY")
    rootElement.appendChild bodyElement
    
    ' Add header information
    Call AddElement(xmlDoc, bodyElement, "HZ", "1")
    Call AddElement(xmlDoc, bodyElement, "HNUM", "1")
    Call AddElement(xmlDoc, bodyElement, "HZY", Year(Now()))
    Call AddElement(xmlDoc, bodyElement, "R00G1", "01")
    Call AddElement(xmlDoc, bodyElement, "R01G1S", "Company Name")
    Call AddElement(xmlDoc, bodyElement, "R01G2S", TIN)
    Call AddElement(xmlDoc, bodyElement, "R01G3S", "001")
    Call AddElement(xmlDoc, bodyElement, "R01G4S", "Main Activity")
    Call AddElement(xmlDoc, bodyElement, "R01G5S", "Ukraine")
    Call AddElement(xmlDoc, bodyElement, "M01", "1")
    
    ' Find last row with data (skip header rows)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Process data rows (assuming data starts from row 5)
    Dim dataStartRow As Long
    dataStartRow = 5
    
    ' Add table rows
    For rowNum = dataStartRow To lastRow
        If ws.Cells(rowNum, 1).Value <> "" Then
            
            ' T1RXXXXG2S - Operation code (col 2)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG2S", _
                CStr(ws.Cells(rowNum, 2).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG3S - Operation subject type (col 3)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG3S", _
                CStr(ws.Cells(rowNum, 5).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG4S - Description (col 4)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG4S", _
                CStr(ws.Cells(rowNum, 4).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG5S - UKT ZED code (col 6)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG5S", _
                CStr(ws.Cells(rowNum, 6).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG7D - Operation date (col 11)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG7D", _
                Format(ws.Cells(rowNum, 11).Value, "YYYY-MM-DD"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG8S - Contract number (col 7)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG8S", _
                CStr(ws.Cells(rowNum, 7).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG9S - Side code (col 8)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG9S", _
                CStr(ws.Cells(rowNum, 8).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG10S - Country code (col 9)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG10S", _
                CStr(ws.Cells(rowNum, 9).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG11S - Incoterms code (col 10)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG11S", _
                CStr(ws.Cells(rowNum, 10).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG14D - Date from (col 11)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG14D", _
                Format(ws.Cells(rowNum, 11).Value, "YYYY-MM-DD"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG15D - Date to (col 12)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG15D", _
                Format(ws.Cells(rowNum, 12).Value, "YYYY-MM-DD"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG16 - Price (col 13)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG16", _
                Format(ws.Cells(rowNum, 13).Value, "0.00"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG17 - Quantity (col 14)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG17", _
                Format(ws.Cells(rowNum, 14).Value, "0.00"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG18S - Currency code (col 16)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG18S", _
                CStr(ws.Cells(rowNum, 16).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG19S - Exchange rate code (col 17)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG19S", _
                CStr(ws.Cells(rowNum, 17).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG20 - Total amount (col 18)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG20", _
                Format(ws.Cells(rowNum, 18).Value, "0.00"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG21S - Method code (col 19)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG21S", _
                CStr(ws.Cells(rowNum, 19).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG22S - Profitability indicator (col 20)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG22S", _
                CStr(ws.Cells(rowNum, 20).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG23 - Profitability value (col 21)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG23", _
                Format(ws.Cells(rowNum, 21).Value, "0.00"), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG26S - Source type (col 26)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG26S", _
                CStr(ws.Cells(rowNum, 26).Value), rowNum - dataStartRow + 1)
            
            ' T1RXXXXG27S - Source name (col 27)
            Call AddTableElement(xmlDoc, bodyElement, "T1RXXXXG27S", _
                CStr(ws.Cells(rowNum, 27).Value), rowNum - dataStartRow + 1)
        End If
    Next rowNum
    
    ' Add footer
    Call AddElement(xmlDoc, bodyElement, "R01G20", "0")
    
    ' Save XML file
    xmlDoc.Save filePath
    
    MsgBox "XML file successfully exported to: " & filePath, vbInformation, "Export Complete"
    
    Exit Sub
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical, "Error"
End Sub

Sub AddElement(xmlDoc As Object, parentElement As Object, tagName As String, tagValue As String)
    Dim newElement As Object
    Set newElement = xmlDoc.createElement(tagName)
    newElement.Text = CStr(tagValue)
    parentElement.appendChild newElement
End Sub

Sub AddTableElement(xmlDoc As Object, parentElement As Object, tagName As String, tagValue As String, rowNum As Long)
    Dim newElement As Object
    Set newElement = xmlDoc.createElement(tagName)
    newElement.Text = CStr(tagValue)
    newElement.setAttribute "ROWNUM", CLng(rowNum)
    parentElement.appendChild newElement
End Sub

Function IsNumeric(inputValue As String) As Boolean
    IsNumeric = Not IsError(CDbl(inputValue))
End Function

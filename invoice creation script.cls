VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SubmitBilling()

    Dim wsForm As Worksheet
    Dim ws As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim nextRow As Long
    Dim checkRange As Range
    

    Set wsForm = ThisWorkbook.Sheets("Billings")
    Set ws = ThisWorkbook.Sheets("Billings retrieval and update")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Billings")
    
    
    ' NO NULL VALUES
    
If wsForm.Range("F9").value = "" Or _
   wsForm.Range("H9").value = "" Or _
   wsForm.Range("F13").value = "" Or _
   wsForm.Range("H13").value = "" Then

    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Please fill all required fields!", vbExclamation
    Exit Sub
End If

    'check for duplicates
    Set checkRange = tWS.Range("A:F")
    
If Application.WorksheetFunction.CountIfs( _
    tWS.Range("B:B"), wsForm.Range("F9").value, _
    tWS.Range("C:C"), ">=" & Date, _
    tWS.Range("C:C"), "<" & Date + 1, _
    tWS.Range("D:D"), wsForm.Range("H9").value, _
    tWS.Range("E:E"), wsForm.Range("F13").value) > 0 Then
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "EXISTING RECORD!!!", vbExclamation
    Exit Sub

End If
    With tWS.Columns("A").NumberFormat = "@"
    tWS.Columns("B").NumberFormat = "@"
    tWS.Columns("C").NumberFormat = "dd/mm/yyy"
    tWS.Columns("E").NumberFormat = "$#,###.0"
    tWS.Columns("F").NumberFormat = "@"
    End With
    
    ' GENERATE ID
    nextRow = tWS.Cells(tWS.Rows.Count, "A").End(xlUp).Row + 1
    newID = "B_" & Format(nextRow - 1, "00000")

    ' show bill id
    ws.Range("K6") = newID
    
    ' Save ID
    tWS.Cells(nextRow, 1).value = newID

    ' Transfer data
    tWS.Cells(nextRow, 2).value = wsForm.Range("F9") ' Patient ID
    tWS.Cells(nextRow, 3).value = Date ' Date
    tWS.Cells(nextRow, 4).value = wsForm.Range("H9").value ' Service
    tWS.Cells(nextRow, 5).value = wsForm.Range("F13").value ' Amount
    tWS.Cells(nextRow, 6).value = wsForm.Range("H13").value ' Payment status

      
    ' Clear form after submit
    wsForm.Range("F9,H9,F13,H13").ClearContents
    
    'BACK TO ENTRY PAGE
    ws.Activate
    
     patientID = ws.Range("K6").value

    Set foundcell = tWS.Range("A:A").Find(patientID, lookat:=xlWhole)
    
    'LOAD BIILLS
     ws.Range("K12:P12").value = tWS.Cells(foundcell.Row, 2).value ' PATIENT ID
    ws.Range("K20:P20").value = tWS.Cells(foundcell.Row, 3).value ' DATE
    ws.Range("K14:P14").value = tWS.Cells(foundcell.Row, 4).value ' SERVICE
    ws.Range("K16:P16").value = tWS.Cells(foundcell.Row, 5).value ' AMOUNT
    ws.Range("K18:P18").value = tWS.Cells(foundcell.Row, 6).value ' PAMENT STATUS
    ws.Range("K22:P22").value = tWS.Cells(foundcell.Row, 1).value ' BILL ID
    ws.Range("K24:P24").value = "QUADE_CARE HEALTH CENTER" ' BILLER
    
    MsgBox "Invoice Created successfully!"
    
    'Save the wprkbook
    tWB.Save
    
    ' CLOSE THE WORKBOOK
    tWB.Close
End Sub





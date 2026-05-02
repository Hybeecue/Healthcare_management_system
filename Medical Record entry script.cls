VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Submitmedicalrecord()

    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim nextRow As Long
    Dim checkRange As Range
    

    Set wsForm = ThisWorkbook.Sheets("Medical Records")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Medical Records")
    
    
    ' NO NULL VALUES
    
If wsForm.Range("F10").value = "" Or _
   wsForm.Range("F18").value = "" Or _
   wsForm.Range("H18").value = "" Or _
   wsForm.Range("F22").value = "" Then

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
    tWS.Range("B:B"), wsForm.Range("F13").value, _
    tWS.Range("C:C"), ">=" & Date, _
    tWS.Range("C:C"), "<" & Date + 1, _
    tWS.Range("D:D"), wsForm.Range("H13").value, _
    tWS.Range("E:E"), wsForm.Range("F17").value) > 0 Then
    
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
    tWS.Columns("F").NumberFormat = "@"
    End With
    
    ' GENERATE ID
    nextRow = tWS.Cells(tWS.Rows.Count, "A").End(xlUp).Row + 1
    newID = "R_" & Format(nextRow - 1, "000")

    ' Save ID
    tWS.Cells(nextRow, 1).value = newID

    ' Transfer data
    tWS.Cells(nextRow, 2).value = wsForm.Range("F10") ' Patient ID
    tWS.Cells(nextRow, 3).value = Date ' Date
    tWS.Cells(nextRow, 4).value = wsForm.Range("H10").value ' Diagnosis
    tWS.Cells(nextRow, 5).value = wsForm.Range("F14").value ' Treatment
    tWS.Cells(nextRow, 6).value = wsForm.Range("H14").value ' Notes
    tWS.Cells(nextRow, 7).value = wsForm.Range("F18").value ' Heights
    tWS.Cells(nextRow, 8).value = wsForm.Range("H18").value ' Weights
    tWS.Cells(nextRow, 9).value = wsForm.Range("F22").value ' Blood Pressure
    tWS.Cells(nextRow, 10).value = wsForm.Range("H22").value ' Doctor
    
    ' Clear form after submit
    wsForm.Range("F10,H10,F14,H14,F18,H18,F22,H22").ClearContents
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    'Save the wprkbook
    tWB.Save
    
    MsgBox "Medical Record saved successfully!"
End Sub



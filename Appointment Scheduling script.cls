VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub SubmitAppointmentData()

    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim nextRow As Long
    Dim checkRange As Range
    

    Set wsForm = ThisWorkbook.Sheets("Appointments")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Appointment")
    
    
    ' NO NULL VALUES
    
If wsForm.Range("F8").value = "" Or _
   wsForm.Range("H8").value = "" Or _
   wsForm.Range("F12").value = "" Or _
   wsForm.Range("H12").value = "" Then
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Please fill all required fields!", vbExclamation
    Exit Sub
End If

If Application.WorksheetFunction.CountIfs( _
    tWS.Range("D:D"), wsForm.Range("F12").value, _
    tWS.Range("E:E"), wsForm.Range("H12").value) > 5 Then

    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Appointment not Available!!!", vbExclamation
    Exit Sub

End If

    With tWS.Columns("A").NumberFormat = "@"
    tWS.Columns("B").NumberFormat = "@"
    tWS.Columns("D").NumberFormat = "dd/mm/yyy"
    tWS.Columns("E").NumberFormat = "hh:mm:ss"
    End With
    
    ' GENERATE ID
    nextRow = tWS.Cells(tWS.Rows.Count, "A").End(xlUp).Row + 1
    newID = "AP" & Format(nextRow - 1, "000")

    ' Save ID
    tWS.Cells(nextRow, 1).value = newID

    ' Transfer data
    tWS.Cells(nextRow, 2).value = wsForm.Range("F8").value ' Patient_ID
    tWS.Cells(nextRow, 3).value = wsForm.Range("H8").value ' Doctor
    tWS.Cells(nextRow, 4).value = wsForm.Range("F12").value ' Date
    tWS.Cells(nextRow, 5).value = wsForm.Range("H12").value ' Time
    tWS.Cells(nextRow, 6).value = "Scheduled" ' status

    ' Clear form after submit
    wsForm.Range("F8,H8,F12,H12").ClearContents

    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    'Save the wprkbook
    tWB.Save
    
    MsgBox "Appointment Scheduled successfully!"
End Sub





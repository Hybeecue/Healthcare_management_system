Attribute VB_Name = "enterpatientdetails"
Sub SubmitpatientData()

    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim nextRow As Long
    Dim checkRange As Range
    

    Set wsForm = ThisWorkbook.Sheets("PatientS")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Patients_data")
    

    ' NO NULL VALUES
    
If wsForm.Range("F7").value = "" Or _
   wsForm.Range("H7").value = "" Or _
   wsForm.Range("F11").value = "" Or _
   wsForm.Range("H11").value = "" Or _
   wsForm.Range("F15").value = "" Or _
   wsForm.Range("H15").value = "" Or _
   wsForm.Range("F19").value = "" Or _
   wsForm.Range("H19").value = "" Then

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
    Set checkRange = tWS.Range("A:H")
    
If Application.WorksheetFunction.CountIfs( _
    tWS.Range("B:B"), wsForm.Range("F7").value, _
    tWS.Range("C:C"), wsForm.Range("H7").value, _
    tWS.Range("D:D"), wsForm.Range("F11").value, _
    tWS.Range("E:E"), wsForm.Range("H11").value, _
    tWS.Range("F:F"), wsForm.Range("F15").value, _
    tWS.Range("G:G"), wsForm.Range("H15").value, _
    tWS.Range("H:H"), wsForm.Range("F19").value, _
    tWS.Range("I:I"), wsForm.Range("H19").value) > 0 Then

    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Patient Already Registered!!!", vbExclamation
    Exit Sub

End If
    With tWS.Columns("A").NumberFormat = "@"
    tWS.Columns("F").NumberFormat = "@"
    tWS.Columns("E").NumberFormat = "dd/mm/yyy"
    End With
    
    ' GENERATE ID
    nextRow = tWS.Cells(tWS.Rows.Count, "A").End(xlUp).Row + 1
    newID = "P" & Format(nextRow - 1, "000")

    ' Save ID
    tWS.Cells(nextRow, 1).value = newID

    ' Transfer data
    tWS.Cells(nextRow, 2).value = wsForm.Range("F7").value ' FIRST Name
    tWS.Cells(nextRow, 3).value = wsForm.Range("H7").value ' Last Name
    tWS.Cells(nextRow, 4).value = wsForm.Range("F11").value ' Gender
    tWS.Cells(nextRow, 5).value = wsForm.Range("H11").value ' Date of Birth
    tWS.Cells(nextRow, 6).value = wsForm.Range("F15").value ' Phone
    tWS.Cells(nextRow, 7).value = wsForm.Range("H15").value ' Address
    tWS.Cells(nextRow, 8).value = wsForm.Range("F19").value ' Blood Group
    tWS.Cells(nextRow, 9).value = wsForm.Range("H19").value ' Blood Group
    tWS.Cells(nextRow, 10).value = Now

    ' Clear form after submit
    wsForm.Range("F7,H7,F11,H11,F15,H15,F19,H19").ClearContents

    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    'Save the wprkbook
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Patient record saved successfully!"
End Sub


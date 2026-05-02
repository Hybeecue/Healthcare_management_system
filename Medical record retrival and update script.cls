VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Medicalretrieval()

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim pid As String
    Dim lastrow As Long
    Dim i As Long
    Dim outputrow As Long
    Dim filepath As String
    Dim foundcell As Range
        
    Set ws = ThisWorkbook.Sheets("Medical_Record")
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    Set wb = Workbooks.Open(filepath)
    Set ws2 = wb.Sheets("Medical Records")
    Set ws3 = wb.Sheets("Patients_data")
    
    ' patient id
    pid = ws.Range("G12").value
    
    ' clear contents first
    ws.Range("A15:K1000").ClearContents
    
    Set foundcell = ws3.Range("A:A").Find(pid, lookat:=xlWhole)
    
    lastrow = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    outputrow = 15
    
    ' loop all record
        
    For i = 2 To lastrow
    
    If ws2.Cells(i, 2) = pid Then
    
    'load all data
     ws.Cells(outputrow, 1).value = ws2.Cells(i, 1).value ' REC_ID
    ws.Cells(outputrow, 2).value = ws2.Cells(i, 3).value ' date
    ws.Cells(outputrow, 5).value = ws2.Cells(i, 4).value 'diagnosis
    ws.Cells(outputrow, 6).value = ws2.Cells(i, 5).value 'treatment
    ws.Cells(outputrow, 7).value = ws2.Cells(i, 6).value 'notes
    ws.Cells(outputrow, 8).value = ws2.Cells(i, 7).value 'height
    ws.Cells(outputrow, 9).value = ws2.Cells(i, 8).value 'weight
    ws.Cells(outputrow, 10).value = ws2.Cells(i, 9).value 'Blood pressue
    ws.Cells(outputrow, 11).value = ws2.Cells(i, 10).value 'doctor
    ws.Cells(outputrow, 3).value = ws3.Cells(foundcell.Row, 2).value ' first name
    ws.Cells(outputrow, 4).value = ws3.Cells(foundcell.Row, 3).value 'last name
    
            outputrow = outputrow + 1
        End If
    Next i
    
    If outputrow > 15 Then

    With ws.Sort
        .SortFields.Clear
        .SortFields.Add Key:=ws.Range("B15:B" & outputrow - 1), _
            Order:=xlDescending
        
        .SetRange ws.Range("A15:K" & outputrow - 1)
        .Header = xlNo
        .Apply
    End With

End If

    If outputrow = 15 Then
    
        ws.Activate
         
        wb.Save
        
        wb.Close
         
        MsgBox "No Medical Record Found", vbInformation
        
        Else
        
        ws.Activate
        
        wb.Save
        
        wb.Close
        
        MsgBox "Medical Records Loaded", vbInformation
        
        End If
        Exit Sub
    
End Sub

Sub RecordLoad()

Dim ws As Worksheet
    Dim wb As Workbook
    Dim ws2 As Worksheet
    Dim rec_id As String
    Dim filepath As String
    Dim recell As Range
    Dim foundcell As Range
        
    Set ws = ThisWorkbook.Sheets("Medical_Record")
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    Set wb = Workbooks.Open(filepath)
    Set ws2 = wb.Sheets("Medical Records")

    ' record id
    rec_id = ws.Range("C12").value
    
   Set recell = ws.Range("A:A").Find(rec_id, lookat:=xlWhole)
        
   Set foundcell = ws2.Range("A:A").Find(rec_id, lookat:=xlWhole)
   
   If foundcell = "" Then
   
   wb.Activate
   
   MsgBox "record not found"
   
   wb.Save
   
   wb.Close
   
   Exit Sub
End If
   
   ' load data
   
   ws2.Cells(foundcell.Row, 3).value = ws.Cells(recell.Row, 2).value 'date
   ws2.Cells(foundcell.Row, 4).value = ws.Cells(recell.Row, 5).value 'diagnosis
   ws2.Cells(foundcell.Row, 5).value = ws.Cells(recell.Row, 6).value 'treatment
   ws2.Cells(foundcell.Row, 6).value = ws.Cells(recell.Row, 7).value 'notes
   ws2.Cells(foundcell.Row, 7).value = ws.Cells(recell.Row, 8).value 'heights
   ws2.Cells(foundcell.Row, 8).value = ws.Cells(recell.Row, 9).value 'weight
   ws2.Cells(foundcell.Row, 9).value = ws.Cells(recell.Row, 10).value 'B.P
   ws2.Cells(foundcell.Row, 10).value = ws.Cells(recell.Row, 11).value 'Doctor
   
   wb.Activate
   
   wb.Save
   
   wb.Close
   
   MsgBox "record updated", vbInformation
   
   ws.Range("C12").ClearContents
   
End Sub


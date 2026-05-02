VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub appointmentload()

    Dim ws As Worksheet
    Dim wb As Workbook
    Dim ws2 As Worksheet
    Dim ws3 As Worksheet
    Dim lastrow As Long
    Dim i As Long
    
    Dim patientID As String
    Dim pid As String
    
    Dim bestRow As Long
    Dim bestDate As Date
    Dim currentDate As Variant
    
    Dim todayDate As Date
    Dim foundcell As Range
    todayDate = Date
    
    Set ws = ThisWorkbook.Sheets("Appointment retrival")
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set wb = Workbooks.Open(filepath)
    Set ws2 = wb.Sheets("Appointment")
    Set ws3 = wb.Sheets("Patients_data")
    
      
    ' Extract Patient_ID
    pid = ws.Range("K14").value
    patientID = pid
    
    lastrow = ws2.Cells(Rows.Count, 1).End(xlUp).Row
    
    bestRow = 0
    bestDate = 0
    
    Set foundcell = ws3.Range("A:A").Find(pid, lookat:=xlWhole)
    
    ' ?? FIRST PASS: Upcoming appointments
    For i = 2 To lastrow
        
        If ws2.Cells(i, 2).value = patientID Then
            
            currentDate = ws2.Cells(i, 4).value
            
            If True Then
                
                If currentDate >= todayDate Then
                    
                    If bestRow = 0 Or currentDate < bestDate Then
                        bestRow = i
                        bestDate = currentDate
                    End If
                    
                End If
                
            End If
            
        End If
        
    Next i
    
    ' ?? SECOND PASS: Latest past appointment
    If bestRow = 0 Then
        
        For i = 2 To lastrow
            
            If ws2.Cells(i, 2).value = patientID Then
                
                currentDate = ws2.Cells(i, 4).value
                
                If True Then
                    
                    If currentDate <= todayDate Then
                        
                        If bestRow = 0 Or currentDate > bestDate Then
                            bestRow = i
                            bestDate = currentDate
                        End If
                        
                    End If
                    
                End If
                
            End If
            
        Next i
        
    End If
    
    If bestRow = 0 Then
        
       
        ' SAVE THE WB
    wb.Save
    
     ' CLOSE THE WORKBOOK
    wb.Close
    
        MsgBox "No appointment found", vbExclamation
            
        Exit Sub
    End If

    ' ?? LOAD INTO FORM
       ws.Range("K16:P16").value = ws3.Cells(foundcell.Row, 2).value ' first name
    ws.Range("K18:P18").value = ws3.Cells(foundcell.Row, 3).value ' last name
    ws.Range("K20:P20").value = ws2.Cells(bestRow, 3).value ' DOCTOR
    ws.Range("K22:P22").value = ws2.Cells(bestRow, 4).value ' DATE
    ws.Range("K24:P24").value = ws2.Cells(bestRow, 5).value ' TIME
    ws.Range("K26:P26") = ws2.Cells(bestRow, 6).value ' STATUS
    
    
    ' activatating form
    ws.Activate
    
    ' SAVE THE WB
    wb.Save
    
     ' CLOSE THE WORKBOOK
    wb.Close
    
    MsgBox "Appointment Loaded!", vbExclamation
  End Sub


Sub UpdateAPPOINTMENT()

    Dim ws As Worksheet
    Dim ws2 As Worksheet
    Dim wb As Workbook
    Dim pid As String
    Dim foundcell As Range
    
    Set ws = ThisWorkbook.Sheets("Appointment retrival")
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set wb = Workbooks.Open(filepath)
    Set ws2 = wb.Sheets("Appointment")
    
    pid = ws.Range("k14").value
    Set foundcell = ws2.Range("B:B").Find(pid, lookat:=xlWhole)
    
    ' update data
    If foundcell Is Nothing Then
    
    'BACK TO ENTRY PAGE
    ws.Activate
    
        MsgBox "Patient not found!", vbExclamation
        Exit Sub
    End If
    
     ' update data
     ws2.Cells(foundcell.Row, 3).value = ws.Range("K20:P20").value  ' Doctor
     ws2.Cells(foundcell.Row, 4).value = ws.Range("K22:P22").value ' date
     ws2.Cells(foundcell.Row, 5).value = ws.Range("K24:P24").value ' time
     ws2.Cells(foundcell.Row, 6).value = ws.Range("K26:P26").value ' Status
     
    ' Activate sheet
     ws.Activate
      
      ' SAVE THE WRK BOOK
    wb.Save
    
    ' close the workbook
    wb.Close
    
    ' clear form
    ws.Range("K12:P12,K20:P20,K14:P14,K16:P16,K18:P18,K22:P22,K24:P24,K26:P26").ClearContents
      
    MsgBox "Appointment Updated", vbExclamation

End Sub

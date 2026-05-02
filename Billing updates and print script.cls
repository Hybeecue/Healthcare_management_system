VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet90"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Loadbills()
    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim patientID As String
    Dim foundcell As Range


    Set wsForm = ThisWorkbook.Sheets("Billings retrieval and update")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Billings")
    
    patientID = wsForm.Range("K6").value

    Set foundcell = tWS.Range("A:A").Find(patientID, lookat:=xlWhole)

    If foundcell Is Nothing Then
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
        MsgBox "INVOICE NOT FOUND!!!", vbExclamation
        Exit Sub
    End If

    ' Load data into form
    wsForm.Range("K12:P12").value = tWS.Cells(foundcell.Row, 2).value ' PATIENT ID
    wsForm.Range("K20:P20").value = tWS.Cells(foundcell.Row, 3).value ' DATE
    wsForm.Range("K14:P14").value = tWS.Cells(foundcell.Row, 4).value ' SERVICE
    wsForm.Range("K16:P16").value = tWS.Cells(foundcell.Row, 5).value ' AMOUNT
    wsForm.Range("K18:P18").value = tWS.Cells(foundcell.Row, 6).value ' PAMENT STATUS
    wsForm.Range("K22:P22").value = tWS.Cells(foundcell.Row, 1).value ' BILL ID
    wsForm.Range("K24:P24").value = "QUADE_CARE HEALTH CENTER" ' BILLER
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Invoice Loaded!"

End Sub


Sub UpdateBILLS()

    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim patientID As String
    Dim foundcell As Range

    Set wsForm = ThisWorkbook.Sheets("Billings retrieval and update")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Billings")
    
    patientID = wsForm.Range("K6").value

    ' Find patient
    Set foundcell = tWS.Range("A:A").Find(patientID, lookat:=xlWhole)
    
    If foundcell Is Nothing Then
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Patient not found!", vbExclamation
        Exit Sub
    End If

    ' Update values
    tWS.Cells(foundcell.Row, 2).value = wsForm.Range("K12:P12").value ' PATIENT ID
    tWS.Cells(foundcell.Row, 3).value = wsForm.Range("K20:P20").value ' DATE
    tWS.Cells(foundcell.Row, 4).value = wsForm.Range("K14:P14").value ' SERVICE
    tWS.Cells(foundcell.Row, 5).value = wsForm.Range("K16:P16").value ' AMOUNT
    tWS.Cells(foundcell.Row, 6).value = wsForm.Range("K18:P18").value ' PAYMENT STATUS
         
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    ' clear form
    wsForm.Range("K12:P12,K20:P20,K14:P14,K16:P16,K18:P18,K22:P22,K24:P24").ClearContents
    
    MsgBox "INVOICE UPDATED !!!", vbInformation

End Sub

Sub printinvoice()
    Dim wsForm As Worksheet
    Dim ans As VbMsgBoxResult
    
    Set wsForm = ThisWorkbook.Sheets("Billings retrieval and update")
    
        
    ans = MsgBox("Print Invoice Preview?", vbYesNo + vbQuestion, "Print Otion")
    
    If ans = vbYes Then
    
    'show preview
            wsForm.Range("H8:R26").PrintPreview
        Else
        
        'print out
        wsForm.Range("H8:R26").PrintOut
    End If
    
End Sub


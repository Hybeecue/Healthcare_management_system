VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub LoadPatient()
    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim patientID As String
    Dim foundcell As Range


    Set wsForm = ThisWorkbook.Sheets("Patients Update")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Patients_data")
    
    patientID = wsForm.Range("N12").value

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

    ' Load data into form
     wsForm.Range("N16:S16").value = tWS.Cells(foundcell.Row, 2).value ' FIRST NAME
    wsForm.Range("N18:S18").value = tWS.Cells(foundcell.Row, 3).value ' LAST NAME
    wsForm.Range("N20:S20").value = tWS.Cells(foundcell.Row, 4).value ' GENDER
    wsForm.Range("N22:S22").value = tWS.Cells(foundcell.Row, 5).value ' DOB
    wsForm.Range("N24:S24").value = tWS.Cells(foundcell.Row, 6).value ' PHONE
    wsForm.Range("N26:S26").value = tWS.Cells(foundcell.Row, 7).value ' ADDRESS
    wsForm.Range("N28:S28").value = tWS.Cells(foundcell.Row, 8).value ' BLOOD GROUP
     wsForm.Range("N30:S30").value = tWS.Cells(foundcell.Row, 9).value ' GENOTYPE
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
     ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Patient Record Loaded!"

End Sub


Sub UpdatePatient()

    Dim wsForm As Worksheet
    Dim tWB As Workbook
    Dim tWS As Worksheet
    Dim patientID As String
    Dim foundcell As Range


    Set wsForm = ThisWorkbook.Sheets("Patients Update")
    ' Path to the other workbook
    filepath = "C:\Users\IBRAHIM\Desktop\Moyin\IB\excel\DATABASE.xlsx"
    ' open workbook
    Set tWB = Workbooks.Open(filepath)
    Set tWS = tWB.Sheets("Patients_data")
    
    patientID = wsForm.Range("N12").value

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
    tWS.Cells(foundcell.Row, 2).value = wsForm.Range("N16:S16").value ' FIRST NAME
    tWS.Cells(foundcell.Row, 3).value = wsForm.Range("N18:S18").value ' LAST NAME
    tWS.Cells(foundcell.Row, 4).value = wsForm.Range("N20:S20").value ' GENDER
    tWS.Cells(foundcell.Row, 5).value = wsForm.Range("N22:S22").value ' DOB
    tWS.Cells(foundcell.Row, 6).value = wsForm.Range("N24:S24").value ' PHONE
    tWS.Cells(foundcell.Row, 7).value = wsForm.Range("N26:S26").value ' ADDRESS
    tWS.Cells(foundcell.Row, 8).value = wsForm.Range("N28:S28").value ' BLOOD GROUP
    tWS.Cells(foundcell.Row, 9).value = wsForm.Range("N30:S30").value ' GENOTYPE
    
    'BACK TO ENTRY PAGE
    wsForm.Activate
    
    ' SAVE THE WRK BOOK
    tWB.Save
    
    ' close the workbook
    tWB.Close
    
    MsgBox "Patient information updated!", vbInformation
 wsForm.Range("N16:S16,N18:S18,N20:S20,N22:S22,N24:S24,N26:S26,N28:S28,N30:S30").ClearContents
End Sub

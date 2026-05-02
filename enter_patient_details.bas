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

Sub insertdata()
    Dim ws As Worksheet
    Dim nextRow As Long
    Dim checkRange As Range
    Dim conn As Object
    Dim rs As Object
    Dim sqlcheck As String
    Dim sqlinsert As String
    Dim pat_id As String
    Dim dob As Date
    Dim age As Integer
    Dim lastID As String
    Dim numpart As Long
    Dim rawID As Variant
     
    Set ws = ThisWorkbook.Sheets("patients")
    
    Set conn = CreateObject("ADODB.connection")
    
    conn.Open "DSN=MySQLConnection;UID=root;PWD=Moyinoluwa@1;"
    
    If ws.Range("F7").value = "" Or _
   ws.Range("H7").value = "" Or _
   ws.Range("F11").value = "" Or _
   ws.Range("H11").value = "" Or _
   ws.Range("F15").value = "" Or _
   ws.Range("H15").value = "" Or _
   ws.Range("F19").value = "" Or _
   ws.Range("H19").value = "" Then
   
    MsgBox "Please fill all required fields!", vbExclamation
    Exit Sub
End If


    'check for duplicates
        sqlcheck = "SELECT COUNT(*) as total from patient_data where " & _
        "first_name='" & ws.Range("F7").value & "' AND " & _
        "last_name='" & ws.Range("H7").value & "' AND " & _
        "gender='" & ws.Range("F11").value & "' AND " & _
        "DOB='" & Format(CDate(ws.Range("H11").value), "yyyy-mm-dd") & "' AND " & _
        "phone_number='" & ws.Range("F15").value & "' AND " & _
        "address='" & ws.Range("H15").value & "' AND " & _
        "blood_group='" & ws.Range("F19").value & "' AND " & _
        "genotype='" & ws.Range("H19").value & "'"
        
        Set rs = conn.Execute(sqlcheck)
        If rs.Fields("total").value > 0 Then
            MsgBox "Patient Already Registered"
            conn.Close
            Exit Sub
        End If
        
        ' calculating age
        dob = ws.Range("H11").value
        age = DateDiff("yyyy", dob, Date)
        
        ' checking if the birthmonth has arrived
        If Date < DateSerial(Year(Date), Month(dob), Day(dob)) Then
        age = age - 1
        End If
4
        ' generating patient id
       Set rs = conn.Execute("SELECT MAX(patient_id) AS lastID FROM patient_data")
    rawID = rs("lastID").value

If IsNull(rawID) Or rawID = "" Then
    pat_id = "P0001"
Else
    pat_id = "Qua/26/" & Format(CLng(rawID) + 1, "0000")
End If
    
        ' insert data
       sqlinsert = "INSERT INTO patient_data (first_name, last_name, gender, DOB, phone_number, address, blood_group, genotype, registered_date, age, pat_id) VALUES ('" & _
        SafeSQL(ws.Range("F7").value) & "','" & _
        SafeSQL(ws.Range("H7").value) & "','" & _
        SafeSQL(ws.Range("F11").value) & "','" & _
        Format(CDate(ws.Range("H11").value), "yyyy-mm-dd") & "','" & _
        SafeSQL(ws.Range("F15").value) & "','" & _
        SafeSQL(ws.Range("H15").value) & "','" & _
        SafeSQL(ws.Range("F19").value) & "','" & _
        SafeSQL(ws.Range("H19").value) & "','" & _
        Format(Now(), "yyyy-mm-dd HH:mm:ss") & "'," & _
        age & ",'" & pat_id & "')"
            
    Debug.Print sqlinsert
    conn.Execute sqlinsert
    
        ' Clear form after submit
    ws.Range("F7,H7,F11,H11,F15,H15,F19,H19").ClearContents
    
    conn.Close
    Set conn = Nothing

    MsgBox "patient recorded successfully Patient ID = " + pat_id, vbInformation
    ws.Range("f7").Activate

End Sub


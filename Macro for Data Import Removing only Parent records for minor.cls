VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub TransformData()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim i As Integer
    Dim j As Integer
    Dim dob As Date
    Dim age As Integer
    
    Set wsSource = ThisWorkbook.Sheets("Source Material")
    Set wsTarget = ThisWorkbook.Sheets("Import Template")
    
    j = 2 ' Start from the second row in the target sheet
    
    For i = 2 To wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
        'check the condtion for number of rows
        ' Parent row
        firstOccurrence = wsSource.Cells(i, 22).Value
        childRowNum = wsSource.Cells(i, 23).Value
        ' Logic for creating parent and child rows
        ' Get the birth date from column 13
        dob = wsSource.Cells(i, 14).Value
        age = Year(Date) - Year(dob)
            
        If Date < DateSerial(Year(Date), Month(dob), Day(dob)) Then
        age = age - 1 ' Adjust if the person hasn't had their birthday yet this year
        End If
        
            ' Set IsMinor based on age
        Dim isMinor As Integer
        isMinor = IIf(age < 18, 1, 0) ' 1 if minor, 0 if adult
        
        If firstOccurrence = 1 And childRowNum = 1 And isMinor = 0 Then
            ' Create parent row
            wsTarget.Cells(j, 1).Value = wsSource.Cells(i, 24).Value ' ParentSignedWaiverId
            wsTarget.Cells(j, 3).Value = wsSource.Cells(i, 9).Value ' FirstName
            wsTarget.Cells(j, 4).Value = wsSource.Cells(i, 11).Value ' LastName
            wsTarget.Cells(j, 5).Value = Format(wsSource.Cells(i, 14).Value, "MM/DD/YYYY") ' DateOfBirth
            wsTarget.Cells(j, 6).Value = wsSource.Cells(i, 15).Value ' Email
            wsTarget.Cells(j, 7).Value = wsSource.Cells(i, 12).Value ' ContactNumber
            
            
            
            wsTarget.Cells(j, 8).Value = isMinor ' IsMinor for parent
            wsTarget.Cells(j, 9).Value = Date ' CreatedDate (using BookingId as placeholder)
            wsTarget.Cells(j, 9).NumberFormat = "MM/DD/YYYY" ' CreatedDate Format
            wsTarget.Cells(j, 10).Value = DateAdd("yyyy", 1, wsTarget.Cells(j, 9).Value)
            wsTarget.Cells(j, 10).NumberFormat = "MM/DD/YYYY" ' ExpiryDate
            j = j + 1
            
            ' Create child row
            wsTarget.Cells(j, 1).Value = wsSource.Cells(i, 24).Value ' ParentSignedWaiverId
            wsTarget.Cells(j, 3).Value = Split(wsSource.Cells(i, 1).Value, " ")(0) ' FirstName
            
            Dim nameParts As Variant

            nameParts = Split(wsSource.Cells(i, 1).Value, " ") ' Split by space
            
            wsTarget.Cells(j, 4).Value = nameParts(UBound(nameParts)) ' LastName
            wsTarget.Cells(j, 5).Value = Format(wsSource.Cells(i, 2).Value, "MM/DD/YYYY") ' DateOfBirth (from first column)
            
            ' Get the birth date from column 13
            dob = wsSource.Cells(i, 2).Value
            age = Year(Date) - Year(dob)
            
            If Date < DateSerial(Year(Date), Month(dob), Day(dob)) Then
            age = age - 1 ' Adjust if the person hasn't had their birthday yet this year
            End If
        
            ' Set IsMinor based on age
            
            isMinor = IIf(age < 18, 1, 0) ' 1 if minor, 0 if adult
            
            
            wsTarget.Cells(j, 8).Value = isMinor ' IsMinor for child
            wsTarget.Cells(j, 9).Value = Date ' CreatedDate (using BookingId as placeholder)
            wsTarget.Cells(j, 9).NumberFormat = "MM/DD/YYYY" ' CreatedDate Format
            wsTarget.Cells(j, 10).Value = DateAdd("yyyy", 1, wsTarget.Cells(j, 9).Value)
            wsTarget.Cells(j, 10).NumberFormat = "MM/DD/YYYY" ' ExpiryDate
            j = j + 1
        
        ElseIf firstOccurrence = 1 And childRowNum = 1 And isMinor = 1 Then
            ' Create child row
            wsTarget.Cells(j, 1).Value = wsSource.Cells(i, 24).Value ' ParentSignedWaiverId
            wsTarget.Cells(j, 3).Value = Split(wsSource.Cells(i, 1).Value, " ")(0) ' FirstName
            
            

            nameParts = Split(wsSource.Cells(i, 1).Value, " ") ' Split by space
            
            wsTarget.Cells(j, 4).Value = nameParts(UBound(nameParts)) ' LastName
            wsTarget.Cells(j, 5).Value = Format(wsSource.Cells(i, 2).Value, "MM/DD/YYYY") ' DateOfBirth (from first column)
            
            ' Get the birth date from column 13
            dob = wsSource.Cells(i, 2).Value
            age = Year(Date) - Year(dob)
            
            If Date < DateSerial(Year(Date), Month(dob), Day(dob)) Then
            age = age - 1 ' Adjust if the person hasn't had their birthday yet this year
            End If
        
            ' Set IsMinor based on age
            
            isMinor = IIf(age < 18, 1, 0) ' 1 if minor, 0 if adult
            
            
            wsTarget.Cells(j, 8).Value = isMinor ' IsMinor for child
            wsTarget.Cells(j, 9).Value = Date ' CreatedDate (using BookingId as placeholder)
            wsTarget.Cells(j, 9).NumberFormat = "MM/DD/YYYY" ' CreatedDate Format
            wsTarget.Cells(j, 10).Value = DateAdd("yyyy", 1, wsTarget.Cells(j, 9).Value)
            wsTarget.Cells(j, 10).NumberFormat = "MM/DD/YYYY" ' ExpiryDate
            j = j + 1
        
        
        ElseIf firstOccurrence = 0 And childRowNum = 2 Then
            ' Create only child row (no parent row)
            wsTarget.Cells(j, 1).Value = wsSource.Cells(i, 24).Value ' ParentSignedWaiverId
            wsTarget.Cells(j, 3).Value = Split(wsSource.Cells(i, 1).Value, " ")(0) ' FirstName
            

            nameParts = Split(wsSource.Cells(i, 1).Value, " ") ' Split by space
            
            wsTarget.Cells(j, 4).Value = nameParts(UBound(nameParts)) ' LastName
            wsTarget.Cells(j, 5).Value = Format(wsSource.Cells(i, 2).Value, "MM/DD/YYYY") ' DateOfBirth (from first column)
            
            ' Get the birth date from column 13
            dob = wsSource.Cells(i, 2).Value
            age = Year(Date) - Year(dob)
            
            If Date < DateSerial(Year(Date), Month(dob), Day(dob)) Then
            age = age - 1 ' Adjust if the person hasn't had their birthday yet this year
            End If
        
            ' Set IsMinor based on age
            
            isMinor = IIf(age < 18, 1, 0) ' 1 if minor, 0 if adult
            
            
            wsTarget.Cells(j, 8).Value = isMinor ' IsMinor for child
            wsTarget.Cells(j, 9).Value = Date ' CreatedDate (using BookingId as placeholder)
            wsTarget.Cells(j, 9).NumberFormat = "MM/DD/YYYY" ' CreatedDate Format
            wsTarget.Cells(j, 10).Value = DateAdd("yyyy", 1, wsTarget.Cells(j, 9).Value)
            wsTarget.Cells(j, 10).NumberFormat = "MM/DD/YYYY" ' ExpiryDate
            j = j + 1
        End If
    Next i
End Sub

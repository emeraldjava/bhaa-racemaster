Attribute VB_Name = "Module3"
Sub ExtractMemberDetails()
'
' This macro runs from the BHAA Extract Worksheet
' It will copy all the required fields for Registration to the Membership worksheet

Dim RowNo, Mem_RowNo, ColNo, Mem_ColNo, LastRow, Reg_RowNo
Dim BHAA_ID, Lastname, FirstName, Gender, Std, DoB, Category, CompanyName, CompanyNo, Age
Dim Dates(85), CategoryM(85), CategoryL(85), Ages(85), Cntr, CategoryCheck

Application.ScreenUpdating = False

' Read Dates Worksheet to get the age category for each possible age for the date of the race
' Save in Category table for men CategoryM and Ladies CategoryL

Worksheets("Dates").Activate
For RowNo = 1 To 85
    Dates(RowNo) = Cells(RowNo, 3).Value
    CategoryM(RowNo) = Cells(RowNo, 5).Value
    CategoryL(RowNo) = Cells(RowNo, 6).Value
    Ages(RowNo) = Cells(RowNo, 7).Value
Next RowNo

Mem_RowNo = 2
Reg_RowNo = 2

' Open Membership Worksheet and Clear Contents
Sheets("Membership").Activate
Range("A3:J9999").Select
Selection.ClearContents
Cells(3, 1).Select

' Open Registration Worksheet and Clear Contents
Sheets("Registration").Activate
Range("A3:M9999").Select
Selection.ClearContents
Cells(3, 1).Select

' Open Pre-Registered Worksheet and Clear Contents
Sheets("Pre-Registered").Activate
Range("A3:M9999").Select
Selection.ClearContents
Cells(3, 1).Select

' Open BHAA Extract Worksheet
Worksheets("BHAA Extract").Activate

' Determine the last row no
LastRow = WorksheetFunction.CountA(Range("A:A")) + 1

' Read each row in turn until we reach last row
' Save details in equivalent variable
    
For RowNo = 2 To LastRow
    BHAA_ID = Cells(RowNo, 1).Value
    Lastname = Cells(RowNo, 4).Value
    FirstName = Cells(RowNo, 3).Value
    Gender = Cells(RowNo, 7).Value
    If Gender = "F" Then
        Gender = "W"
    End If
    Std = Cells(RowNo, 10).Value
    DoB = Cells(RowNo, 11).Value
    CompanyName = Trim(Cells(RowNo, 9).Value)
    CompanyNo = Cells(RowNo, 8).Value
    Pre_Registered = Cells(RowNo, 12).Value
    
    Worksheets("Dates").Activate
    Cntr = 85

' Compare DoB with the dates saved in the Category tables CategoryM and CategoryL
' Save the calculated Category
    Category = ""
    Do While Cntr > 1
        If DoB < Dates(Cntr) Then
            Age = Ages(Cntr)
            If Gender = CategoryM(1) Then
                Category = CategoryM(Cntr)
            End If
            If Gender = CategoryL(1) Then
                Category = CategoryL(Cntr)
            End If
            Cntr = 1
        End If
        Cntr = Cntr - 1
            
    Loop
    
' Open the Membership worksheet and copy data saved from the BHAA Extract worksheet

    Worksheets("Membership").Activate
    Mem_RowNo = Mem_RowNo + 1
    Cells(Mem_RowNo, 1).Value = BHAA_ID
    Cells(Mem_RowNo, 2).Value = Lastname
    Cells(Mem_RowNo, 3).Value = FirstName
    Cells(Mem_RowNo, 4).Value = Gender
    Cells(Mem_RowNo, 5).Value = Std
    Cells(Mem_RowNo, 6).Value = DoB
    Cells(Mem_RowNo, 7).Value = Category
    Cells(Mem_RowNo, 8).Value = CompanyName
    Cells(Mem_RowNo, 9).Value = CompanyNo
   
' Open the Registration worksheet and copy data for those with Pre-Registered Flag = "Y"
    If Pre_Registered = "Y" Then
        Worksheets("Pre-Registered").Activate
        Reg_RowNo = Reg_RowNo + 1
        Cells(Reg_RowNo, 2).Value = BHAA_ID
        Cells(Reg_RowNo, 3).Value = Lastname
        Cells(Reg_RowNo, 4).Value = FirstName
        Cells(Reg_RowNo, 5).Value = Gender
        Cells(Reg_RowNo, 6).Value = Std
        Cells(Reg_RowNo, 7).Value = DoB
        Cells(Reg_RowNo, 8).Value = Category
        Cells(Reg_RowNo, 9).Value = CompanyName
        Cells(Reg_RowNo, 10).Value = CompanyNo
        Cells(Reg_RowNo, 11).Value = "Y"
    End If

    Worksheets("BHAA Extract").Activate
    
Next RowNo

Application.ScreenUpdating = True

End Sub

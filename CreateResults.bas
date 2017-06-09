Attribute VB_Name = "CreateResults"
Sub CreateResults()
'
' This macro runs to Create the Combined Worksheet which is input to Race Master results calculation
' It will merge the Finishing Order worksheet with the Registration worksheet

Dim RowNo, ColNo, LastRow, LastRegRow, Reg_RowNo, Finish_RowNo
Dim BHAA_ID, Lastname, FirstName, Gender, Std, DoB, Category, CompanyName, CompanyNo
Dim RaceNo As Long, Reg_RaceNo As Long, FinishTime, Found_Flag

Application.ScreenUpdating = False

' Set start Row no for Combined Worksheet to start at 2
RowNo = 2

' Open Combined Worksheet and Clear Contents
Sheets("Combined").Activate
Range("A2:Q9999").Select
Selection.ClearContents

' Open Finishing Order Worksheet and Determine the last row no
Worksheets("Finishing Order").Activate
LastRow = WorksheetFunction.CountA(Range("A:A"))

' Open Registration Worksheet and Determine the last row no
Worksheets("Registration").Activate
LastRegRow = WorksheetFunction.CountA(Range("A:A")) + 1
Range("M3:M9999").Select
Selection.ClearContents

' Read each row in turn on Finishing Order Worksheet until we reach last row
' For each Race Number found, scroll through Registration worksheet and pick up details for runner with that Race Number
For Finish_RowNo = 2 To LastRow
    Worksheets("Finishing Order").Activate
    RaceNo = Cells(Finish_RowNo, 1).Value
    FinishTime = Cells(Finish_RowNo, 2).Value
    
    Worksheets("Registration").Activate
    Found_Flag = "False"
    For Reg_RowNo = 3 To LastRegRow
        Reg_RaceNo = Cells(Reg_RowNo, 1).Value
        If Reg_RaceNo = RaceNo Then
            BHAA_ID = Cells(Reg_RowNo, 2).Value
            Lastname = Cells(Reg_RowNo, 3).Value
            FirstName = Cells(Reg_RowNo, 4).Value
            Gender = Cells(Reg_RowNo, 5).Value
            Std = Cells(Reg_RowNo, 6).Value
            DoB = Cells(Reg_RowNo, 7).Value
            Category = Cells(Reg_RowNo, 8).Value
            CompanyName = Trim(Cells(Reg_RowNo, 9).Value)
            CompanyNo = Cells(Reg_RowNo, 10).Value
            Found_Flag = "True"
            Cells(Reg_RowNo, 13).Value = "Y"
            Exit For
        End If
    Next Reg_RowNo
    
' Open the Combined worksheet and copy data for the current Race Number
    Worksheets("Combined").Activate
    If Found_Flag = "True" Then
        Cells(RowNo, 1).Value = Finish_RowNo - 1
        Cells(RowNo, 2).Value = RaceNo
        Cells(RowNo, 3).Value = BHAA_ID
        Cells(RowNo, 4).Value = FinishTime
        Cells(RowNo, 5).Value = Lastname
        Cells(RowNo, 6).Value = FirstName
        Cells(RowNo, 7).Value = Gender
        Cells(RowNo, 8).Value = Std
        Cells(RowNo, 9).Value = DoB
        Cells(RowNo, 10).Value = Category
        Cells(RowNo, 11).Value = CompanyName
        Cells(RowNo, 12).Value = CompanyNo
        RowNo = RowNo + 1
    Else
        Cells(RowNo, 1).Value = Finish_RowNo - 1
        Cells(RowNo, 2).Value = RaceNo
        Cells(RowNo, 4).Value = FinishTime
        RowNo = RowNo + 1
    End If
Next Finish_RowNo

Worksheets("Combined").Activate
    
Cells(2, 1).Select

Application.ScreenUpdating = True

End Sub


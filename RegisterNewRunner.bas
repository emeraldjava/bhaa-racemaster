Attribute VB_Name = "RegisterNewRunner"
Sub RegisterNewRunner()


Dim Reg_RowNo As Long

First_Name = First_Name.Value
Sur_Name = Sur_Name.Value
Option_Button1 = OptionButton1.Value

Date_of_Birth = Date_of_Birth.Value
Race_Number = Race_Number.Value

'Make Sheet1 active
Worksheets("Registration").Activate

'Determine Reg_RowNo
Reg_RowNo = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information
Cells(Reg_RowNo, 3).Value = First_Name.Value
Cells(Reg_RowNo, 4).Value = Sur_Name.Value

If OptionButton1.Value = True Then
    Cells(Reg_RowNo, 5).Value = "M"
Else
    Cells(Reg_RowNo, 5).Value = "L"
End If

Cells(Reg_RowNo, 7).Value = Date_of_Birth.Value
Cells(Reg_RowNo, 1).Value = Race_Number.Value

'Close Userform
' Unload Me

Worksheets("Membership").Activate


End Sub

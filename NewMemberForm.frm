VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} NewMemberForm 
   Caption         =   "NewMemberForm"
   ClientHeight    =   6400
   ClientLeft      =   50
   ClientTop       =   370
   ClientWidth     =   9830
   OleObjectBlob   =   "NewMemberForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "NewMemberForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
Dim Reg_RowNo As Long
Dim RaceNo As Long
Dim Dates(85), CategoryM(85), CategoryL(85), Ages(85), Cntr, CategoryCheck, Gender
Dim DoB As Date, DoB1
Dim Category_Date As Date
' Read Dates Worksheet to get the age category for each possible age for the date of the race
' Save in Category table for men CategoryM and Ladies CategoryL

Worksheets("Dates").Activate
For RowNo = 1 To 85
    Dates(RowNo) = Cells(RowNo, 3).Value
    CategoryM(RowNo) = Cells(RowNo, 5).Value
    CategoryL(RowNo) = Cells(RowNo, 6).Value
    Ages(RowNo) = Cells(RowNo, 7).Value
Next RowNo


' Check if Race Number already allocated to Pre-Registered runner   *******************************************

Worksheets("Pre-Registered").Activate

'Determine Last RowNo on Pre-Registered List
Reg_RowNo = WorksheetFunction.CountA(Range("C:C")) + 1

RaceNo = Race_Number.Value
For RowNo = 3 To Reg_RowNo
    If Cells(RowNo, 1).Value = RaceNo Then
        MsgBox "Race Number already allocated on Pre-Registered Line " & RowNo
        Exit Sub
    End If
Next RowNo


' Check if Race Number already allocated on Registration List       *******************************************

Worksheets("Registration").Activate

'Determine Last RowNo on Registration List
Reg_RowNo = WorksheetFunction.CountA(Range("C:C")) + 1

RaceNo = Race_Number.Value
For RowNo = 3 To Reg_RowNo
    If Cells(RowNo, 1).Value = RaceNo Then
        MsgBox "Race Number already allocated on Line " & RowNo
        Exit Sub
    End If
Next RowNo

'Check Valid date entered

Date_of_Birth.Value = Format(Date_of_Birth.Value, "dd/mm/yyyy")

If IsDate(Trim(Date_of_Birth.Value)) = False Then
    MsgBox "Invalid Date of Birth"
    Exit Sub
End If

If OptionButton1.Value = True Then
    Gender = "M"
Else
    Gender = "W"
End If

' Compare DoB with the dates saved in the Category tables CategoryM and CategoryL
' Save the calculated Category
Cntr = 85
Category = ""
DoB = Date_of_Birth.Value
Do While Cntr > 10
    Category_Date = Left(Dates(Cntr), 10)
    If DoB < Category_Date Then
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

' If no Category allocated
If Category = "" Then
    MsgBox "Check Date of Birth entered"
    Exit Sub
End If
DoB1 = "dd/mm/yyyy"
dd = Left(Date_of_Birth.Value, 2)
mm = Mid(Date_of_Birth.Value, 4, 2)
yyyy = Right(Date_of_Birth.Value, 5)
If dd < 13 And mm < 13 Then
    DoB1 = mm & "/" & dd & yyyy
Else
    DoB1 = Date_of_Birth.Value
End If
    

'Transfer information to next line
Reg_RowNo = Reg_RowNo + 1
Cells(Reg_RowNo, 1).Value = Race_Number.Value
Cells(Reg_RowNo, 3).Value = Last_Name.Value
Cells(Reg_RowNo, 4).Value = First_Name.Value
Cells(Reg_RowNo, 5).Value = Gender
Cells(Reg_RowNo, 7).Value = DoB1
Cells(Reg_RowNo, 8).Value = Category
Cells(Reg_RowNo, 9).Value = Company_Name.Value
Cells(Reg_RowNo, 12).Value = Entry_Fee_Paid.Value

First_Name.Value = ""
Last_Name.Value = ""
OptionButton1.Value = False
OptionButton2.Value = False
Date_of_Birth.Value = ""
Race_Number.Value = ""

' Save WorkBook
ActiveWorkbook.Save

'Close Userform
Unload Me

Cells(Reg_RowNo, 1).Select


End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub NewMember_Click()

End Sub

Private Sub OptionButton1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

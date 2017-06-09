VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RegisterForm 
   Caption         =   "RegisterForm"
   ClientHeight    =   5304
   ClientLeft      =   50
   ClientTop       =   370
   ClientWidth     =   7200
   OleObjectBlob   =   "RegisterForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RegisterForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
Dim Mem_RowNo As Long
Dim lCount As Long
Dim BHAA_Number As Long
Dim BHAA_ID_No As Long
Dim Match_Found


'Make Sheet1 active
Worksheets("Membership").Activate

'Determine Last_RowNo
Last_RowNo = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information to Listbox

If BHAA_ID.Value <> "" Then
    If IsNumeric(BHAA_ID.Value) Then
    Else
        MsgBox "Invalid BHAA ID number"
        Exit Sub
    End If
End If

If BHAA_ID.Value <> "" Then
    For Mem_RowNo = 3 To Last_RowNo
        BHAA_Number = Cells(Mem_RowNo, 1).Value
        BHAA_ID_No = BHAA_ID.Value
        If Cells(Mem_RowNo, 1).Value = BHAA_ID_No Then
            ListBox1.AddItem Cells(Mem_RowNo, 1).Value & " " & Cells(Mem_RowNo, 2).Value & " " & Cells(Mem_RowNo, 3).Value & " " & Cells(Mem_RowNo, 6).Value
            Exit For
        End If
    Next Mem_RowNo
Else
    If Lastname.Value <> "" Then
        Match_Found = "False"
        For Mem_RowNo = 3 To Last_RowNo
            Second_Name = Cells(Mem_RowNo, 2).Value
            Form_Name = Lastname.Value
            If Second_Name = Lastname.Value Then
                ListBox1.AddItem Cells(Mem_RowNo, 1).Value & " " & Cells(Mem_RowNo, 2).Value & " " & Cells(Mem_RowNo, 3).Value & " " & Cells(Mem_RowNo, 6).Value
                Match_Found = "True"
            End If
        Next Mem_RowNo
        If Match_Found = "False" Then
            MsgBox "No Name match found - Name is case sensitive"
            Exit Sub
        End If
    End If
        
End If

Entry_Fee_Paid.Value = 10

'Close Userform
' Unload Me

Worksheets("Membership").Activate

End Sub

Private Sub CommandButton2_Click()
Dim Mem_RowNo As Long
Dim lCount As Long
Dim BHAA_Number
Dim BHAA_ID_No As Long
Dim L_Name, F_Name, Gender, Standard, DoB, Category, Company, CompanyNo, EntryFee
Dim RaceNo As Long
Dim Reg_RowNo, Last_RowNo, Cntr, Prev_RaceNo

'Make Sheet1 active
Worksheets("Membership").Activate

'Determine Last_RowNo
Last_RowNo = WorksheetFunction.CountA(Range("A:A")) + 1

'Transfer information from ListBox
For Cntr = 0 To 20
    If ListBox1.Selected(Cntr) = True Then
        If IsNumeric(Left(ListBox1.List(Cntr), 5)) Then
            BHAA_Number = Left(ListBox1.List(Cntr), 5)
        Else
            If IsNumeric(Left(ListBox1.List(Cntr), 4)) Then
                BHAA_Number = Left(ListBox1.List(Cntr), 4)
            Else
                If IsNumeric(Left(ListBox1.List(Cntr), 3)) Then
                    BHAA_Number = Left(ListBox1.List(Cntr), 3)
                Else
                    BHAA_Number = Left(ListBox1.List(Cntr), 2)
                End If
            End If
        End If
    End If
Next Cntr

If Race_Number.Value = "" Then
    MsgBox "Please enter Race Number to continue"
    Exit Sub
End If

If BHAA_Number <> "" Then
    For Mem_RowNo = 3 To Last_RowNo
        BHAA_ID_No = BHAA_Number
        If Cells(Mem_RowNo, 1).Value = BHAA_ID_No Then
            L_Name = Cells(Mem_RowNo, 2).Value
            F_Name = Cells(Mem_RowNo, 3).Value
            Gender = Cells(Mem_RowNo, 4).Value
            Standard = Cells(Mem_RowNo, 5).Value
            DoB = Cells(Mem_RowNo, 6).Value
            Category = Cells(Mem_RowNo, 7).Value
            Company = Cells(Mem_RowNo, 8).Value
            CompanyNo = Cells(Mem_RowNo, 9).Value
            RaceNo = Race_Number.Value
            EntryFee = Entry_Fee_Paid.Value
            
            Worksheets("Pre-Registered").Activate

            'Determine Last RowNo on Registration List
            Reg_RowNo = WorksheetFunction.CountA(Range("C:C")) + 1
            
            ' Check if Race number allocated already and if BHAA ID already Registered for this race
            For RowNo = 3 To Reg_RowNo
                If Cells(RowNo, 1).Value = RaceNo Then
                    MsgBox "Race Number already allocated to Pre-Registered Line " & RowNo
                    Exit Sub
                End If
                If Cells(RowNo, 2).Value = BHAA_ID_No Then
                    MsgBox "BHAA ID for this Member already Pre-Registered on Line " & RowNo
                    Exit Sub
                End If

            Next RowNo
            
            Worksheets("Registration").Activate

            'Determine Last RowNo on Registration List
            Reg_RowNo = WorksheetFunction.CountA(Range("C:C")) + 1
            
            ' Check if Race number allocated already and if BHAA ID already Registered for this race
            For RowNo = 3 To Reg_RowNo
                If Cells(RowNo, 1).Value = RaceNo Then
                    MsgBox "Race Number already allocated on Line " & RowNo
                    Exit Sub
                End If
                If Cells(RowNo, 2).Value = BHAA_ID_No Then
                    MsgBox "BHAA ID for this Member already Registered on Line " & RowNo
                    Exit Sub
                End If

            Next RowNo
            
            'Transfer information to Registration Sheet
            Reg_RowNo = Reg_RowNo + 1
            Cells(Reg_RowNo, 1).Value = RaceNo
            Cells(Reg_RowNo, 2).Value = BHAA_ID_No
            Cells(Reg_RowNo, 3).Value = L_Name
            Cells(Reg_RowNo, 4).Value = F_Name
            Cells(Reg_RowNo, 5).Value = Gender
            Cells(Reg_RowNo, 6).Value = Standard
            Cells(Reg_RowNo, 7).Value = DoB
            Cells(Reg_RowNo, 8).Value = Category
            Cells(Reg_RowNo, 9).Value = Company
            Cells(Reg_RowNo, 10).Value = CompanyNo
            Cells(Reg_RowNo, 12).Value = EntryFee

            Exit For
        End If
    Next Mem_RowNo
        
End If

Race_Number.Value = ""

' Save WorkBook
ActiveWorkbook.Save

'Close Userform
Unload Me

Cells(Reg_RowNo, 1).Select

End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Initialize()

End Sub

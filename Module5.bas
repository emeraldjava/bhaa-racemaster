Attribute VB_Name = "Module5"
Sub Calc()

'
' Calculation Macro
' Macro recorded 29/06/2004 by Frank Wade
'
' Keyboard Shortcut: Ctrl+p
'
' This macro calculates the individual and team winners for Ladies and Men
' All calculations are based on the Combined List

' The "Combined" worksheet contains the master listing of all runners and places.
' The "Mens" worksheet contains the "Men's Teams" with Names in team points order.
' The "Mens List" worksheet contains all men from Combined list.
'       The "Mens List" worksheet also contains individual women who did not make a
'       place on one of the Women's teams. These women are eligible for a place on a Mixed team competing in the Mens section.
' The "Mens Teams" worksheet contains the standards, numbers and places of all Mens teams,
'       sorted into Grades and points order. There must be 3 runners on each team
' The "Ladies" worksheet contains the individual ladies winners and teams in points order.
' The "Ladies List" worksheet contains all ladies from Combined list with a membership number.
' The "Dates" worksheet contains the date of the race and the age category for all ages based on Date of Birth

'============================================================================================================================

Dim RowNo, ColNo, RowNo1, RowNo2, PointsTot, LadyFlag
Dim TeamNo, TeamName, Place, RaceNo, Name, RaceTime, Std, Gender, LadyPlace, MensPlace
Dim LTeamNo, TPlace, TeamPlace, Col1, Grade, StdTotal
Dim LastRow, LastLadyRow, LastMensRow
Dim LVet(10), MVet(10), J, JJ, LVetPlace(10), MVetPlace(10)
Dim ClassNo, Class1, Class2, Class3, Class4, HoldGrade, NewClass
Dim LadiesOnly_Flag, LastMensListRow, LastMensTeamsRow
Dim DoB, Dates(85), CategoryM(85), CategoryL(85), Ages(85), Cntr, CategoryCheck, Category
Dim RacePlace
Dim Member_Place(3), Member_RaceNo(3), Member_Name(3), Member_Std(3)

'=============================================================================================================================
' Step 0 : Check category is correct by checking DoB with "Dates" Worksheet
' Step 1 : Read through "Combined" Worksheet and separate Mens and Ladies into working sheets "Ladies","Ladies List","Mens","Mens List".
' Step 2 : Read through the "Ladies List" and extract those on ladies' teams
' Step 3 : Read Through "Ladies List" and extract ladies who are not part of a team and add to "Mens List"
' Step 4 : Read through the "Mens List" and extract the teams to the  "Mens teams" worksheet
' Step 5 : Write Prize Winning Mens teams to "Mens" worksheet
' Step 6 : VETS Categories : Read Through "Combined List" and extract vets to "Mens Vets" and "ladies Vets"
' Step 7 : Copy contents of 'Mens Teams' to 'New Mens Teams' sheet and Allocate a Class of 1 to 4 based on Grades A to D
' Step 8 : PRINT SECTION : Copy results to "Print Mens Results" and "Print Ladies Results"

'==============================================================================================================================

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

' Open Combined Worksheet
' Select Column E and scroll down to last row to get value for LastRow
Worksheets("Combined").Activate
Range("B2").Select
Selection.End(xlDown).Select
LastRow = ActiveCell.Row

' Read each row in turn until we reach last row
' Save date of birth in variable DoB
' Set all Column's font colour to Black
    
For RowNo = 2 To LastRow
    DoB = Cells(RowNo, 9).Value
    If DoB = "" Then
    Else
        Cntr = 85
        For ColNo = 1 To 16
            Cells(RowNo, ColNo).Select
            Selection.Font.ColorIndex = 0
        Next ColNo
        Cells(RowNo, 16).Value = ""

' Trim trailing blanks from Team name in Combined list
        TeamName = Trim(Cells(RowNo, 11).Value)
        Cells(RowNo, 11).Value = TeamName
   
' Compare DoB with the dates saved in the Category tables CategoryM and CategoryL
' Write the calculated Category to column 15 titled CategoryCheck and also to column 10 titled category

        Do While Cntr > 1
            If DoB < Dates(Cntr) Then
                Cells(RowNo, 14).Value = Ages(Cntr)
                If Cells(RowNo, 7).Value = CategoryM(1) Then
                    Cells(RowNo, 15).Value = CategoryM(Cntr)
                End If
                If Cells(RowNo, 7).Value = CategoryL(1) Then
                    Cells(RowNo, 15).Value = CategoryL(Cntr)
                End If
                Cntr = 1
            End If
            Cntr = Cntr - 1
        Loop
        
    End If ' End of Check for DoB = ""
    
' If the Category in the Category column does not match the Category in CategoryCheck
' Highlight the value in Category in Red and write a warning message in column 16

    If Cells(RowNo, 15).Value <> Cells(RowNo, 10).Value Then
        Cells(RowNo, 10).Select
        Selection.Font.ColorIndex = 3
        Cells(RowNo, 15).Select
        Selection.Font.ColorIndex = 3
        CategoryCheck = "Category changed from "
        Category = Cells(RowNo, 10).Value
        If Category = "" Or Category = " " Then
            Category = "Blank"
        End If
        Cells(RowNo, 16).Value = CategoryCheck & Category & " to " & Cells(RowNo, 15).Value
        Cells(RowNo, 16).Select
        Selection.Font.ColorIndex = 3
        Cells(RowNo, 10).Value = Cells(RowNo, 15).Value
    End If
    
Next RowNo


'===========================================================================================

' Step 1 : Read through the "Combined List" and separate into men and Ladies
'       Step 1A :  Check if each runner has a valid race number
'       Step 1B :  Write all ladies to "Ladies List"
'       Step 1C :  Write first three individual ladies to "Ladies" sheet
'       Step 1D :  Write all men with a Standard to "Mens List"
'       Step 1E :  Write first three individual Men to "Mens" sheet

LadiesOnly_Flag = False

' Select Column A and scroll down to last row to get value for LastRow
Worksheets("Combined").Activate
Range("b2").Select
Selection.End(xlDown).Select
LastRow = ActiveCell.Row

Cells(1, 27).Value = "b2"                       ' Save LastRow value in "Combined" sheet Column AB
Cells(1, 28).Value = LastRow

RowNo1 = 3      ' Use RowNo1 for Ladies list starting point
LadyPlace = 0
MensPlace = 0
RowNo2 = 3      ' Use RowNo2 as Row number for "Mens list" worksheet

'   Auto-populate Place no in Column A
For RowNo = 2 To LastRow
    Place = RowNo - 1
    Cells(RowNo, 1).Value = Place
Next RowNo

'   Repeat for each row on the Combined" Worksheet
For RowNo = 2 To LastRow
    
    Worksheets("Combined").Activate
    TeamNo = Trim(Cells(RowNo, 12).Value)
    Cells(RowNo, 12).Value = TeamNo
    Gender = Cells(RowNo, 7).Value

'   Step 1A :   Check if each runner has a race number
    If IsNumeric(Cells(RowNo, 2).Value) Then
        TeamName = Cells(RowNo, 11).Value
        Place = Cells(RowNo, 1).Value
        RaceNo = Cells(RowNo, 2).Value
        RaceTime = Cells(RowNo, 4).Value
        Name = Trim(Cells(RowNo, 6).Value) & " " & Trim(Cells(RowNo, 5).Value)
        Std = Cells(RowNo, 8).Value
        
'   Step 1B :  Write all ladies with race numbers to "Ladies List"
        If Gender = "L" Or Gender = "F" Or Gender = "W" Then
            Worksheets("Ladies List").Activate
            RowNo1 = RowNo1 + 1
            Cells(RowNo1, 1).Value = RowNo1 - 3
            Cells(RowNo1, 2).Value = Place
            Cells(RowNo1, 3).Value = RaceNo
            Cells(RowNo1, 4).Value = Name
            Cells(RowNo1, 5).Value = Std
            Cells(RowNo1, 6).Value = TeamName
            Cells(RowNo1, 7).Value = TeamNo
            Cells(RowNo1, 8).Value = "*"
        
'   Step 1C :  Write first three individual ladies to "Ladies" sheet
            If LadyPlace < 3 Then
                LadyPlace = LadyPlace + 1
                Worksheets("Ladies").Activate
                Cells(LadyPlace + 4, 1).Value = LadyPlace
                Cells(LadyPlace + 4, 2).Value = Name
                Cells(LadyPlace + 4, 3).Value = TeamNo
                Cells(LadyPlace + 4, 4).Value = RaceTime
                Cells(LadyPlace + 4, 7).Value = Place
                Cells(LadyPlace + 4, 8).Value = RaceNo
                Cells(LadyPlace + 4, 9).Value = TeamName
                Cells(LadyPlace + 4, 10).Value = Std
                Worksheets("Combined").Activate
                Cells(RowNo, 17).Value = LadyPlace
            End If
        End If
        
'   Step 1D :  Write all men with a Standard to "Mens List"
'       -   Add women with no team to this list later
        If Gender = "M" And IsNumeric(Std) And Std > 0 Then
            Worksheets("Mens List").Activate
            RowNo2 = RowNo2 + 1
            Cells(RowNo2, 1).Value = RowNo2 - 3
            Cells(RowNo2, 2).Value = Place
            Cells(RowNo2, 3).Value = RaceNo
            Cells(RowNo2, 4).Value = Name
            Cells(RowNo2, 5).Value = Std
            Cells(RowNo2, 6).Value = TeamName
            Cells(RowNo2, 7).Value = TeamNo
        End If
            
'   Step 1E :  Write first three individual Men to "Mens" sheet
        If Gender = "M" And MensPlace < 3 Then
            Worksheets("Mens").Activate
            MensPlace = MensPlace + 1
            Cells(MensPlace + 4, 1).Value = MensPlace
            Cells(MensPlace + 4, 2).Value = Name
            Cells(MensPlace + 4, 3).Value = TeamNo
            Cells(MensPlace + 4, 4).Value = RaceTime
            Cells(MensPlace + 4, 7).Value = Place
            Cells(MensPlace + 4, 8).Value = RaceNo
            Cells(MensPlace + 4, 9).Value = TeamName
            Cells(MensPlace + 4, 10).Value = Std
            Worksheets("Combined").Activate
            Cells(RowNo, 17).Value = MensPlace
        End If
        
    End If

Next RowNo

 
'===========================================================================================

' Step 2 : Read through the "Ladies List" and extract and process ladies' teams
'       Step 2A :  Sort "Ladies List" by team no and race place
'       Step 2B :  Read the sorted Ladies List and move team by team to "Ladies Teams"
'       Step 2C :  Sort by "Ladies Teams" team no and race place
'       Step 2D :  Write Prize Winning ladies teams to Ladies worksheet
'       Step 2E :  Get Place and Race number from "Ladies teams" for Runner 1, 2 and 3

' Skip this section if no Ladies in Race
' Steps 2, 3 and 4 relate to Ladies Race
' If there are no ladies, skip this section.

If RowNo1 = 3 And LadyPlace = 0 Then
    LadyFlag = False
Else
    LadyFlag = True
    Worksheets("Ladies List").Activate
         
    ' Step 2A  :  Sort "Ladies List" by team no and race place
    Range("a4:g999").Select
    Range("g999").Activate
    Selection.Sort Key1:=Range("G4"), Order1:=xlDescending, Key2:=Range("b4") _
        , Order2:=xlAscending, Header _
        :=xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("a1:a1").Select
    
    ' Step 2B :  Read the sorted Ladies List and move team by team to "Ladies Teams"
    '
    ' Team list on Ladies Teams starts at Row 5
    LTeamNo = 5
    
    ' Select Column B and scroll down to last row to get value for LastRow
    Worksheets("Ladies List").Activate
    Range("g4").Select
    Selection.End(xlDown).Select
    LastRow = ActiveCell.Row
    Cells(1, 27).Value = "g4"                   ' ***********  Change 20110131 - two lines added to save LastRow value in "Ladies List" sheet
    Cells(1, 28).Value = LastRow
    
    ' Repeat for each row -
    
    For RowNo = 4 To LastRow
        
        Worksheets("Ladies List").Activate
    
        TeamNo = Cells(RowNo, 7).Value
        If TeamNo = "" Or TeamNo = 1 Then
            Exit For
        End If
               
        If Cells(RowNo, 7).Value = Cells(RowNo + 1, 7).Value Then
            Team_Runners = 2
            If Cells(RowNo, 7).Value = Cells(RowNo + 2, 7).Value Then
                Team_Runners = 3
                TeamName = Cells(RowNo, 6).Value
                For Cntr = 0 To 2
                    Member_Place(Cntr) = Cells(RowNo + Cntr, 2).Value
                    Member_RaceNo(Cntr) = Cells(RowNo + Cntr, 3).Value
                    Member_Name(Cntr) = Cells(RowNo + Cntr, 4).Value
                    Member_Std(Cntr) = Cells(RowNo + Cntr, 5).Value
                    Cells(RowNo + Cntr, 8).Value = ""                         ' Blank out * in column 8 for Ladies with a team
                Next Cntr
                RowNo = RowNo + 2
                
                Worksheets("Ladies Teams").Activate
                Cells(LTeamNo, 2).Value = TeamNo
                ColNo = 0
                
                For Cntr = 0 To 2
                    Cells(LTeamNo, ColNo + 3).Value = Member_RaceNo(Cntr)
                    Cells(LTeamNo, ColNo + 4).Value = Member_Place(Cntr)
                    Cells(LTeamNo, ColNo + 5).Value = Member_Name(Cntr)
                    Cells(LTeamNo, ColNo + 6).Value = Member_Std(Cntr)
                    ColNo = ColNo + 4
                Next Cntr
                
                Cells(LTeamNo, 15).Value = Member_Place(0) + Member_Place(1) + Member_Place(2)
                Cells(LTeamNo, 16).Value = TeamName
                
                LTeamNo = LTeamNo + 1
            End If
        End If
        
    Next RowNo
            
    Worksheets("Ladies Teams").Activate
    
    ' Step 2C :  Sort by "Ladies Teams" team no and race place
    
    Range("B4:P99").Select
    Range("P99").Activate
    Selection.Sort Key1:=Range("O4"), Order1:=xlAscending, Key2:=Range("L4") _
        , Order2:=xlAscending, Header _
        :=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("a1:a1").Select

    ' Select Column B and scroll down to last row to get value for LastRow of "Ladies Teams"
    Worksheets("Ladies Teams").Activate
    
    If Cells(5, 2).Value = "" Then                      ' Check for full Ladies team
        LastRow = 4
        Cells(1, 27).Value = "row 5 col 2 is blank"    ' ***********  Change 20110131 - two lines added to save LastRow value in "Ladies Teams" sheet
        Cells(1, 28).Value = LastRow
    Else
        Range("b4").Select
        Selection.End(xlDown).Select
        LastRow = ActiveCell.Row

        Cells(1, 27).Value = "b4"               ' ***********  Change 20110131 - two lines added to save LastRow value in "Ladies Teams" sheet
        Cells(1, 28).Value = LastRow
    End If
    
                
    ' Step 2D :  Write Prize Winning ladies teams to Ladies worksheet
    '           Get Place Runner details from "Ladies teams" for Runner 1, 2 and 3
    
    LadyPlace = 13
              
    For LTeamNo = 5 To LastRow
        Worksheets("Ladies Teams").Activate
        LadyPlace = LadyPlace + 1
        If IsNumeric(Cells(LTeamNo, 2).Value) Then
            TeamPlace = Cells(LTeamNo, 1).Value
            TeamNo = Cells(LTeamNo, 2).Value
            ColNo = 0
            For Cntr = 0 To 2
                Member_RaceNo(Cntr) = Cells(LTeamNo, ColNo + 3).Value
                Member_Place(Cntr) = Cells(LTeamNo, ColNo + 4).Value
                Member_Name(Cntr) = Cells(LTeamNo, ColNo + 5).Value
                Member_Std(Cntr) = Cells(LTeamNo, ColNo + 6).Value
                ColNo = ColNo + 4
            Next Cntr
            PointsTot = Cells(LTeamNo, 15).Value
            TeamName = Cells(LTeamNo, 16).Value
            
            Worksheets("Ladies").Activate
            Cells(LadyPlace, 2).Value = TeamName
            Cells(LadyPlace, 3).Value = TeamNo
            Cells(LadyPlace, 5).Value = PointsTot
            For Cntr = 0 To 2
                Cells(LadyPlace, 7).Value = Member_Place(Cntr)
                Cells(LadyPlace, 8).Value = Member_RaceNo(Cntr)
                Cells(LadyPlace, 9).Value = Member_Name(Cntr)
                Cells(LadyPlace, 10).Value = Member_Std(Cntr)
                LadyPlace = LadyPlace + 1
            Next Cntr
            
        End If              ' End If of IsNumeric
    Next LTeamNo
    
    ' End of Step2E
    
'===========================================================================================

' Step 3 : Read Through "Ladies List" and extract ladies who are not part of a team
'
' These ladies are eligible to make up a man's team so add them to "Mens List"
    
    ' Select Column A and scroll down to last row to get value for LastRow
    Worksheets("Ladies List").Activate
    Range("a3").Select                                  ' Change from "a4" to "a3" 7th March 2011
    Selection.End(xlDown).Select
    LastLadyRow = ActiveCell.Row
    Cells(2, 27).Value = "Step 3 a4"
    Cells(2, 28).Value = LastLadyRow
    
    For RowNo = 4 To LastLadyRow
        Worksheets("Ladies List").Activate
        
        If TeamNo = "" Or TeamNo = 1 Then               ' Exit loop if we have reached day runners or Ladies with no team
            Exit For
        End If
        Std = Cells(RowNo, 5).Value
        If Cells(RowNo, 8).Value = "*" And IsNumeric(Std) And Std > 0 Then
            Place = Cells(RowNo, 2).Value
            RaceNo = Cells(RowNo, 3).Value
            TeamName = Cells(RowNo, 6).Value
            TeamNo = Cells(RowNo, 7).Value
            Name = Cells(RowNo, 4).Value
            
            ' Write unattached ladies with a numeric Standard value to Mens List
            Worksheets("Mens List").Activate
            RowNo2 = RowNo2 + 1                         ' RowNo2 has Row counter for Mens List
            MensPlace = RowNo2 - 3
            Cells(RowNo2, 1).Value = MensPlace
            Cells(RowNo2, 2).Value = Place
            Cells(RowNo2, 3).Value = RaceNo
            Cells(RowNo2, 4).Value = Name
            Cells(RowNo2, 5).Value = Std
            Cells(RowNo2, 6).Value = TeamName
            Cells(RowNo2, 7).Value = TeamNo
    
        End If
        
    Next RowNo

End If
'============================================================================================
'==  End of Ladies List Processing
'============================================================================================

' Step 4 : Read through the "Mens List" and extract the teams to the  "Mens teams" worksheet

Worksheets("Mens List").Activate
    
' Sort "Mens List" by team no and race place
    Range("b3:g999").Select
    Range("g999").Activate
    Selection.Sort Key1:=Range("G3"), Order1:=xlDescending, Key2:=Range("b3") _
        , Order2:=xlAscending, Header _
        :=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("a1:a1").Select
    
'===================================================================================================
' Read through the sorted Mens List and move team by team to Mens Teams

Worksheets("Mens List").Activate
LTeamNo = 5
LastMensListRow = RowNo2
Cells(1, 27).Value = "RowNo2"
Cells(1, 28).Value = LastMensListRow

For RowNo = 4 To LastMensListRow
    
    Worksheets("Mens List").Activate

    TeamNo = Trim(Cells(RowNo, 7).Value)
    If TeamNo = "" Or TeamNo = 0 Or TeamNo = 1 Then
        Exit For
    End If

    If Cells(RowNo, 7).Value = Cells(RowNo + 1, 7).Value Then
        Team_Runners = 2
        If Cells(RowNo, 7).Value = Cells(RowNo + 2, 7).Value Then
            Team_Runners = 3
            TeamName = Cells(RowNo, 6).Value
            For Cntr = 0 To 2
                Member_Place(Cntr) = Cells(RowNo + Cntr, 2).Value
                Member_RaceNo(Cntr) = Cells(RowNo + Cntr, 3).Value
                Member_Name(Cntr) = Cells(RowNo + Cntr, 4).Value
                Member_Std(Cntr) = Cells(RowNo + Cntr, 5).Value
            Next Cntr
            RowNo = RowNo + 2
            
            Worksheets("Mens Teams").Activate
            Cells(LTeamNo, 2).Value = TeamNo
            ColNo = 0
            
            For Cntr = 0 To 2
                Cells(LTeamNo, ColNo + 3).Value = Member_RaceNo(Cntr)
                Cells(LTeamNo, ColNo + 4).Value = Member_Place(Cntr)
                Cells(LTeamNo, ColNo + 5).Value = Member_Name(Cntr)
                Cells(LTeamNo, ColNo + 6).Value = Member_Std(Cntr)
                ColNo = ColNo + 4
            Next Cntr
            
            Cells(LTeamNo, 15).Value = Member_Place(0) + Member_Place(1) + Member_Place(2)
            StdTotal = Member_Std(0) + Member_Std(1) + Member_Std(2)
            Cells(LTeamNo, 16).Value = StdTotal
            Cells(LTeamNo, 17).Value = TeamName
    ' Assign Grades to teams based on Standard Total
            Select Case StdTotal
                Case 1 To 30
                    Grade = "A"
                Case 31 To 36
                    Grade = "B"
                Case 37 To 42
                    Grade = "C"
                Case Else
                    Grade = "D"
            End Select
            Cells(LTeamNo, 18).Value = Grade
            
            LTeamNo = LTeamNo + 1
        End If
    End If
    
Next RowNo
       
'============================================================================================================================

' Step 5 : Write Prize Winning Mens teams to "Mens" worksheet
'============================================================================================================================

LastMensTeamsRow = LTeamNo - 1

Worksheets("Mens Teams").Activate
  
' Sort "Mens Teams" by Grade, Total Pts and Place 3
    Range("B4:R99").Select
    Range("R99").Activate
    Selection.Sort Key1:=Range("R4"), Order1:=xlAscending, Key2:=Range("O4") _
        , Order2:=xlAscending, Key3:=Range("L4"), Order1:=xlAscending, Header _
        :=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("a1:a1").Select
    
' Use MensPlace variable to control Row numbers for "Mens" worksheet
MensPlace = 13
HoldGrade = 0

' Check if any Mens teams exist - this change made to cater for Ladies only race
If Cells(5, 2).Value = "" Or Cells(5, 2).Value = " " Then
    LadiesOnly_Flag = True
Else
    For LTeamNo = 5 To LastMensTeamsRow
        Worksheets("Mens Teams").Activate
        If IsNumeric(Cells(LTeamNo, 2).Value) Then
            TeamNo = Cells(LTeamNo, 2).Value
            ColNo = 0
            For Cntr = 0 To 2
                Member_RaceNo(Cntr) = Cells(LTeamNo, ColNo + 3).Value
                Member_Place(Cntr) = Cells(LTeamNo, ColNo + 4).Value
                Member_Name(Cntr) = Cells(LTeamNo, ColNo + 5).Value
                Member_Std(Cntr) = Cells(LTeamNo, ColNo + 6).Value
                ColNo = ColNo + 4
            Next Cntr
            PointsTot = Cells(LTeamNo, 15).Value
            TeamName = Cells(LTeamNo, 17).Value
            StdTotal = Member_Std(0) + Member_Std(1) + Member_Std(2)
            Grade = Cells(LTeamNo, 18).Value
           
            Worksheets("Mens").Activate
            If HoldGrade = Grade Then
                TeamPlace = TeamPlace + 1
            Else
                MensPlace = MensPlace + 1
                Cells(MensPlace, 1).Value = "Grade " & Grade
                HoldGrade = Grade
                TeamPlace = 1
            End If
            
            MensPlace = MensPlace + 1
            Cells(MensPlace, 1).Value = TeamPlace
            Cells(MensPlace, 2).Value = TeamName
            Cells(MensPlace, 3).Value = TeamNo
            Cells(MensPlace, 4).Value = StdTotal
            Cells(MensPlace, 5).Value = PointsTot
            Cells(MensPlace, 6).Value = Grade
            For Cntr = 0 To 2
                Cells(MensPlace, 7).Value = Member_Place(Cntr)
                Cells(MensPlace, 8).Value = Member_RaceNo(Cntr)
                Cells(MensPlace, 9).Value = Member_Name(Cntr)
                Cells(MensPlace, 10).Value = Member_Std(Cntr)
                MensPlace = MensPlace + 1
            Next Cntr
            
        End If              ' End If of IsNumeric
    Next LTeamNo


End If

'=========================================================================================================================================
'
' Step 6 : VETS Categories : Read Through "Combined List" and extract vets to "Mens Vets" and "Ladies Vets"
'=========================================================================================================================================

Worksheets("Combined").Activate
    
' Sort "Combined List" by Category and Race Place
    Range("A1:Q9999").Select
    Range("Q9999").Activate
    Selection.Sort Key1:=Range("J1"), Order1:=xlAscending, Key2:=Range("A1") _
        , Order2:=xlAscending, Header _
        :=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
    Range("a1:a1").Select
    
' Set start row numbers for vets on Vets Lists in Increments of 10 starting at 4 e.g. 4, 14, 24 etc
VetRow = 4
For Cntr = 1 To 9
    LVet(Cntr) = VetRow
    MVet(Cntr) = VetRow
    VetRow = VetRow + 10
Next Cntr

' Initialise Vet Place array to zero
For J = 1 To 10
    LVetPlace(J) = 0
    MVetPlace(J) = 0
Next J


' Select Column A and scroll down to last row to get value for LastRow
Range("a2").Select
Selection.End(xlDown).Select
LastRow = ActiveCell.Row

Cells(2, 27).Value = "a2 in Vets step 6"                       ' Save LastRow value in "Combined" sheet
Cells(2, 28).Value = LastRow

RowNo2 = 2

For RowNo = 2 To LastRow
    Worksheets("Combined").Activate
    If Cells(RowNo, 17).Value = "" Then                         ' Check if Individual Place in Column Q to denote 1st 3 in Individual
        Place = Cells(RowNo, 1).Value
        RaceNo = Cells(RowNo, 2).Value
        RaceTime = Cells(RowNo, 4).Value
        TeamName = Cells(RowNo, 11).Value
        TeamNo = Cells(RowNo, 12).Value
        Name = Trim(Cells(RowNo, 6).Value) & " " & Trim(Cells(RowNo, 5).Value)
        Std = Cells(RowNo, 8).Value
        J = 0
        JJ = "S"
        
    ' Write vets to Mens Vets List
        If IsNumeric(Cells(RowNo, 10).Value) Then       ' Numeric Category for Men. Alpha category for Women
            J = Cells(RowNo, 10).Value
            If J > 0 And J < 10 Then
                MVet(J) = MVet(J) + 1
                MVetPlace(J) = MVetPlace(J) + 1
                If MVetPlace(J) < 5 Then
                    Worksheets("Mens Vets").Activate
                    Cells(MVet(J), 2).Value = Place
                    Cells(MVet(J), 3).Value = RaceNo
                    Cells(MVet(J), 4).Value = RaceTime
                    Cells(MVet(J), 5).Value = Name
                    Cells(MVet(J), 6).Value = TeamName
                    Cells(MVet(J), 7).Value = TeamNo
                    Cells(MVet(J), 8).Value = Std
                End If
            End If
        Else
            JJ = Cells(RowNo, 10).Value
            Select Case JJ                      ' Convert Ladies Categories from Alpha to Numeric equivalent
                Case "A"
                    J = 1
                Case "B"
                    J = 2
                Case "C"
                    J = 3
                Case "D"
                    J = 4
                Case "E"
                    J = 5
                Case "F"
                    J = 6
                Case "G"
                    J = 7
                Case "H"
                    J = 8
            End Select
            
    ' Write vets to Ladies Vets List
            If J > 0 And J < 9 Then
                LVet(J) = LVet(J) + 1
                LVetPlace(J) = LVetPlace(J) + 1
                If LVetPlace(J) < 5 Then
                    Worksheets("Ladies Vets").Activate
                    Cells(LVet(J), 2).Value = Place
                    Cells(LVet(J), 3).Value = RaceNo
                    Cells(LVet(J), 4).Value = RaceTime
                    Cells(LVet(J), 5).Value = Name
                    Cells(LVet(J), 6).Value = TeamName
                    Cells(LVet(J), 7).Value = TeamNo
                    Cells(LVet(J), 8).Value = Std
                End If
            End If
    
        End If
    End If
    
Next RowNo


' Re-Sort "Combined List" by race place

Worksheets("Combined").Activate

Range("a1:Q9999").Select
Range("Q9999").Activate
Selection.Sort Key1:=Range("A1"), Order1:=xlAscending, Header _
    :=xlYes, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
Range("a1:a1").Select

'=================================================================================================================================

' Step 7 : Copy contents of 'Mens Teams' to 'New Mens Teams' sheet and Allocate a Class of 1 to 4 based on Grades A to D
'          Allows the Race Administrator to regrade teams into Classes and Balance numbers of Teams in each Class.
'=================================================================================================================================

Sheets("Mens Teams").Select
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("New Mens Teams").Select
Cells.Select
ActiveSheet.Paste
Cells(1, 10).Value = "Enter Class splits in Column N on 'New Mens Teams' sheet"
Cells(3, 19).Value = "Class"


Worksheets("New Mens Teams").Activate
  
'==================================================================================
  
' "Mens Teams" is already sorted by Grade, Total Pts and Place 3

' Skip if a Ladies only race

If LadiesOnly_Flag = False Then

' Select Column B and scroll down to last row to get value for LastRow
    Range("B5").Select
    Selection.End(xlDown).Select
    LastRow = ActiveCell.Row
    
    Cells(1, 27).Value = "B5 where ladiesOnly Flag is false"      ' ***********  Change 20110131 - two lines added to save LastRow value in "New Mens teams" sheet
    Cells(1, 28).Value = LastRow
    
    ClassNo = LastRow - 4

    Class1 = ClassNo \ 4
    Class2 = ClassNo \ 4
    Class3 = ClassNo \ 4
    Class4 = ClassNo - (Class1 + Class2 + Class3)
    
    For RowNo = 5 To LastRow
        If RowNo < Class1 + 5 Then
            Cells(RowNo, 19).Value = 1
        Else
            If RowNo < (Class1 + Class2 + 5) Then
                Cells(RowNo, 19).Value = 2
                Else
                    If RowNo < (Class1 + Class2 + Class3 + 5) Then
                        Cells(RowNo, 19).Value = 3
                    Else
                        Cells(RowNo, 19).Value = 4
                    End If
            End If
        End If
    Next RowNo
    
' Loop through rows from the last to front to adjust classes
' Try this three times
    For Cntr = 1 To 3
        Do
            RowNo = RowNo - 1
            If Cells(RowNo, 19).Value <> Cells(RowNo - 1, 19).Value Then
                If Cells(RowNo, 17).Value = Cells(RowNo - 1, 17).Value Then
                    Cells(RowNo, 19).Value = Cells(RowNo - 1, 19).Value
                End If
            End If
        Loop While RowNo > 8
        
        RowNo = LastRow
    Next Cntr


End If      ' new endif for Ladies Only race

'===================================================================================================
'**************************************************************************************************
'
' Step 8:   PRINT SECTION : Copy results to "Print Mens Results" and "Print Ladies Results"
'
'===================================================================================================

' +++++++++++   Read Through "Mens" worksheet and extract first 3 men for Individual List  ++++++++++
RowNo1 = 3
For RowNo = 5 To 7
    Worksheets("Mens").Activate
    Place = Cells(RowNo, 1).Value
    Name = Cells(RowNo, 2).Value
    TeamNo = Cells(RowNo, 3).Value
    RaceTime = Cells(RowNo, 4).Value
    RacePlace = Cells(RowNo, 7).Value
    RaceNo = Cells(RowNo, 8).Value
    TeamName = Cells(RowNo, 9).Value
    Std = Cells(RowNo, 10).Value
        
' Write 1st 3 INDIVIDUAL MENS PLACINGS on "Print Mens Result" worksheet
    Worksheets("Print Mens Result").Activate
    RowNo1 = RowNo1 + 1
    Cells(RowNo1, 1).Value = Place
    Cells(RowNo1, 2).Value = Name
    Cells(RowNo1, 3).Value = RaceTime
    Cells(RowNo1, 4).Value = ""
    Cells(RowNo1, 5).Value = RacePlace
    Cells(RowNo1, 6).Value = RaceNo
    Cells(RowNo1, 7).Value = TeamName
    Cells(RowNo1, 8).Value = Std
Next RowNo

'===================================================================================================================================
'   Get Vets Results from Mens Vets for transfer to "Print Mens Results"
'    Print MENS VETS PLACINGS

RowNo1 = 9
NewGrade = ""

For RowNo = 1 To 90
    Worksheets("Mens Vets").Activate
    If Cells(RowNo, 1).Value = "Mens Vets Placings" Then
        NewGrade = Cells(RowNo, 5).Value
    End If
    If Cells(RowNo, 1).Value = 1 Then
        RowNo1 = RowNo1 + 1
        For Cntr = 1 To 4
            TeamPlace = Cells(RowNo, 1).Value
            Place = Cells(RowNo, 2).Value
            RaceNo = Cells(RowNo, 3).Value
            RaceTime = Cells(RowNo, 4).Value
            Name = Cells(RowNo, 5).Value
            TeamName = Cells(RowNo, 6).Value
            Std = Cells(RowNo, 8).Value
            RowNo = RowNo + 1
            
            Worksheets("Print Mens Result").Activate
            If NewGrade = "" Then
            Else
                Cells(RowNo1, 1).Value = NewGrade
                NewGrade = ""
                RowNo1 = RowNo1 + 2
            End If
            Cells(RowNo1, 1).Value = TeamPlace
            Cells(RowNo1, 2).Value = Name
            Cells(RowNo1, 3).Value = RaceTime
            Cells(RowNo1, 4).Value = ""
            Cells(RowNo1, 5).Value = Place
            Cells(RowNo1, 6).Value = RaceNo
            Cells(RowNo1, 7).Value = TeamName
            Cells(RowNo1, 8).Value = Std
            RowNo1 = RowNo1 + 1
            
            Worksheets("Mens Vets").Activate
        Next Cntr
    End If

Next RowNo

'========================================================================================================================

' +++++++++++   Read Through "Mens" worksheet and extract first 3 teams in each Class  ++++++++++

' IMPORTANT NOTE :  This section MUST BE COPIED TO "New Mens Team Placings" macro
' When copied REPLACE the line Worksheets("Mens").Activate with Worksheets("New Mens Teams Placings").Activate
' Any changes to team class splits could change team placings so must be reflected in "Print Mens Result" worksheet

'  Print MENS TEAMS PLACINGS


NewGrade = ""

For RowNo = 14 To 300
    Worksheets("Mens").Activate
    
    Select Case Cells(RowNo, 1).Value
        Case "Grade A"
            RowNo1 = 76
            NewGrade = Cells(RowNo, 1).Value
        Case "Grade B"
            RowNo1 = 90
            NewGrade = Cells(RowNo, 1).Value
        Case "Grade C"
            RowNo1 = 104
            NewGrade = Cells(RowNo, 1).Value
        Case "Grade D"
            RowNo1 = 118
            NewGrade = Cells(RowNo, 1).Value
    End Select
        
' Write Headings and Team results in Class A, B, C and D on "Print Mens Result" worksheet

    If Cells(RowNo, 1).Value = 1 Then
        For Cntr = 1 To 11
            TeamPlace = Cells(RowNo, 1).Value
            If TeamPlace = "" Then
                If Cells(RowNo, 2).Value = Cells(RowNo - 1, 2).Value Then
                    TeamName = ""
                Else
                    TeamName = Cells(RowNo, 2).Value
                End If
                StdTotal = ""
                PointsTot = ""
            Else
                TeamName = Trim(Cells(RowNo, 2).Value)
                StdTotal = Cells(RowNo, 4).Value
                PointsTot = Cells(RowNo, 5).Value
            End If
            Place = Cells(RowNo, 7).Value
            RaceNo = Cells(RowNo, 8).Value
            Name = Cells(RowNo, 9).Value
            Std = Cells(RowNo, 10).Value
            RowNo = RowNo + 1
        
            Worksheets("Print Mens Result").Activate
            If NewGrade = "" Then
            Else
                Cells(RowNo1 - 2, 1).Value = NewGrade
                NewGrade = ""
            End If
            Cells(RowNo1, 1).Value = TeamPlace
            Cells(RowNo1, 2).Value = TeamName
            Cells(RowNo1, 3).Value = StdTotal
            Cells(RowNo1, 4).Value = PointsTot
            Cells(RowNo1, 5).Value = Place
            Cells(RowNo1, 6).Value = RaceNo
            Cells(RowNo1, 7).Value = Name
            Cells(RowNo1, 8).Value = Std
            RowNo1 = RowNo1 + 1
            
            Worksheets("Mens").Activate
        Next Cntr
        
    End If
    
    If RowNo1 = 129 Then
        Exit For
    End If

Next RowNo

'===========================   END of SECTION TO BE COPIED TO "NEW MENS TEAMS PLACINGS"   ==========================================


'===================================================================================================
'
' Print LADIES RESULTS

Worksheets("Ladies").Activate

' +++++++++++   Read Through "Ladies" and extract first 3 Ladies for Individual List  ++++++++++

NewGrade = "Individual"

RowNo1 = 2
For RowNo = 5 To 7
    Worksheets("Ladies").Activate
    
    Place = Cells(RowNo, 1).Value
    Name = Trim(Cells(RowNo, 2).Value)
    RaceTime = Cells(RowNo, 4).Value
    RacePlace = Cells(RowNo, 7).Value
    RaceNo = Cells(RowNo, 8).Value
    TeamName = Cells(RowNo, 9).Value
    Std = Cells(RowNo, 10).Value
        
' Write Headings and Individual results on "Print Ladies Result" worksheet
    Worksheets("Print Ladies Result").Activate
           
    If NewGrade = "" Then
    Else
        Cells(RowNo1, 1).Value = NewGrade
        NewGrade = ""
        RowNo1 = RowNo1 + 1
    End If
      
    RowNo1 = RowNo1 + 1
    Cells(RowNo1, 1).Value = Place
    Cells(RowNo1, 2).Value = Name
    Cells(RowNo1, 3).Value = RaceTime
    Cells(RowNo1, 4).Value = ""
    Cells(RowNo1, 5).Value = RacePlace
    Cells(RowNo1, 6).Value = RaceNo
    Cells(RowNo1, 7).Value = TeamName
    Cells(RowNo1, 8).Value = Std
          
    If Place = 4 Then
        Cells(9, 1).Value = "Ladies Vets Placings"
        Exit For
    End If
Next RowNo


'========================================================================================================================
'   Print LADIES VETS

'   Get Vets Results from Ladies Vets for transfer to "Print Ladies Results"

RowNo1 = 9
NewGrade = ""

For RowNo = 1 To 100
    Worksheets("Ladies Vets").Activate
    If Cells(RowNo, 1).Value = "Ladies Vets Placings" Then
        NewGrade = Cells(RowNo, 5).Value
    End If
    If Cells(RowNo, 1).Value = 1 Then
        RowNo1 = RowNo1 + 1
        For Cntr = 1 To 4
            TeamPlace = Cells(RowNo, 1).Value
            Place = Cells(RowNo, 2).Value
            RaceNo = Cells(RowNo, 3).Value
            RaceTime = Cells(RowNo, 4).Value
            Name = Cells(RowNo, 5).Value
            TeamName = Cells(RowNo, 6).Value
            Std = Cells(RowNo, 8).Value
            RowNo = RowNo + 1
        
            Worksheets("Print Ladies Result").Activate
            If NewGrade = "" Then
            Else
                Cells(RowNo1, 1).Value = NewGrade
                NewGrade = ""
                RowNo1 = RowNo1 + 2
            End If
            Cells(RowNo1, 1).Value = TeamPlace
            Cells(RowNo1, 2).Value = Name
            Cells(RowNo1, 3).Value = RaceTime
            Cells(RowNo1, 4).Value = ""
            Cells(RowNo1, 5).Value = Place
            Cells(RowNo1, 6).Value = RaceNo
            Cells(RowNo1, 7).Value = TeamName
            Cells(RowNo1, 8).Value = Std
            RowNo1 = RowNo1 + 1
            
            Worksheets("Ladies Vets").Activate
        Next Cntr
    End If

Next RowNo

'========================================================================================================================
' Print LADIES TEAMS PLACINGS
'    Read Through "Ladies" worksheet and extract first 3 teams

NewGrade = "    "                         ' Making NewGrade = spaces to trigger printing of headings - change if team categories introduced

RowNo1 = 69
For RowNo = 10 To 30
    Worksheets("Ladies").Activate

    If Cells(RowNo, 1).Value = 1 Then
        For Cntr = 1 To 11
            TeamPlace = Cells(RowNo, 1).Value
            If TeamPlace = "" Then
                If Cells(RowNo, 2).Value = Cells(RowNo - 1, 2).Value Then
                    TeamName = ""
                Else
                    TeamName = Cells(RowNo, 2).Value
                End If
                StdTotal = ""
                PointsTot = ""
            Else
                TeamName = Trim(Cells(RowNo, 2).Value)
                PointsTot = Cells(RowNo, 5).Value
            End If
            Place = Cells(RowNo, 7).Value
            RaceNo = Cells(RowNo, 8).Value
            Name = Cells(RowNo, 9).Value
            Std = Cells(RowNo, 10).Value
            RowNo = RowNo + 1
        
            Worksheets("Print Ladies Result").Activate
            If NewGrade = "" Then
            Else
                Cells(RowNo1 - 3, 1).Value = "Ladies Teams Placings"
                Cells(RowNo1 - 2, 1).Value = NewGrade
                NewGrade = ""
            End If
            Cells(RowNo1, 1).Value = TeamPlace
            Cells(RowNo1, 2).Value = TeamName
            Cells(RowNo1, 3).Value = StdTotal
            Cells(RowNo1, 4).Value = PointsTot
            Cells(RowNo1, 5).Value = Place
            Cells(RowNo1, 6).Value = RaceNo
            Cells(RowNo1, 7).Value = Name
            Cells(RowNo1, 8).Value = Std
            RowNo1 = RowNo1 + 1
            
            Worksheets("Ladies").Activate
        Next Cntr
        
    End If

Next RowNo

'===========================   END of SECTION to Print Ladies Result   ==========================================


Worksheets("Mens").Activate
Range("a2").Select

Worksheets("Mens List").Activate
Range("a2").Select

Worksheets("Mens Teams").Activate
Range("a2").Select

Worksheets("Mens Vets").Activate
Range("a2").Select

Worksheets("Ladies").Activate
Range("a2").Select

Worksheets("Ladies List").Activate
Range("a2").Select

Worksheets("Ladies Vets").Activate
Range("a2").Select

Worksheets("Ladies Teams").Activate
Range("a2").Select

Worksheets("Combined").Activate
Range("a2").Select

Worksheets("Print Ladies Result").Activate
Range("a2").Select

Worksheets("Print Mens Result").Activate
Range("a2").Select

'Worksheets("New Mens Teams").Activate
'Range("a2").Select

Application.ScreenUpdating = True

End Sub

Sub Classes()
'
' Classes Macro
' Macro recorded 01/05/2007 by Frank Wade
'
'
' The "New Mens Teams" worksheet contains the standards, numbers and places of all Mens teams,
'       sorted into Classes in Standard Points Order
' Race Administrator can reBalance teams in the Classes by changing Class No in Column S
' This feature no longer in use in BHAA

Dim RowNo, ColNo, RowNo1, RowNo2, PointsTot, Okflag, LadyFlag
Dim TeamNo, TeamName, Place, RaceNo, Name, RaceTime, Std, Gender, LadyPlace, MensPlace
Dim LTeamNo, TeamPlace, Grade, StdTotal
Dim LastRow, LastLadyRow, LastMensRow
Dim AlphaClass(4), ClassNo, Class1, Class2, Class3, Class4, HoldClass, ClassPlace
Dim Member_Place(3), Member_RaceNo(3), Member_Name(3), Member_Std(3)
Dim NewClass, Cntr


Application.ScreenUpdating = False

Sheets("Print Mens Result").Select
Cells.Select
Application.CutCopyMode = False
Selection.Copy
Sheets("New Mens Result").Select
Cells.Select
ActiveSheet.Paste

Sheets("New Mens Teams Placings").Select
Range("A4:J300").Select
Selection.ClearContents

HoldClass = 0

AlphaClass(1) = "A"
AlphaClass(2) = "B"
AlphaClass(3) = "C"
AlphaClass(4) = "D"

Worksheets("New Mens Teams").Activate
  
' Sort "Mens Teams" by Class, Total Pts and Place 3

    Range("B5:S240").Select
    
'=============================================================================================
    
' SORT FOR OLD PC
    Selection.Sort Key1:=Range("S5"), Order1:=xlAscending, Key2:=Range("O5") _
        , Order2:=xlAscending, Key3:=Range("L5"), Order3:=xlAscending, Header:= _
        xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
 
'=============================================================================================
    
    Range("a1:a1").Select
       
' Write Prize Winning Mens teams to "New Mens Teams Placings" worksheet
'
' Use MensPlace variable to control Rows for "New Mens Teams Placings" worksheet
MensPlace = 4
HoldClass = 0
TeamPlace = 0

' Select Column A and scroll down to last row to get value for LastRow
Worksheets("New Mens Teams").Activate
Range("b5").Select
Selection.End(xlDown).Select
LastRow = ActiveCell.Row

'       Get details from "New Mens Teams" for each Runner
For RowNo1 = 5 To LastRow
    Worksheets("New Mens Teams").Activate
    If IsNumeric(Cells(RowNo1, 2).Value) Then
        TeamNo = Cells(RowNo1, 2).Value
        PointsTot = Cells(RowNo1, 15).Value
        StdTotal = Cells(RowNo1, 16).Value
        TeamName = Cells(RowNo1, 17).Value
        ClassNo = Cells(RowNo1, 19).Value
        ColNo = 3
        For Cntr = 1 To 3
            Member_RaceNo(Cntr) = Cells(RowNo1, ColNo).Value
            Member_Place(Cntr) = Cells(RowNo1, ColNo + 1).Value
            Member_Name(Cntr) = Cells(RowNo1, ColNo + 2).Value
            Member_Std(Cntr) = Cells(RowNo1, ColNo + 3).Value
            ColNo = ColNo + 4
        Next Cntr
        If HoldClass = ClassNo Then
            TeamPlace = TeamPlace + 1
        Else
            MensPlace = MensPlace + 1
            HoldClass = ClassNo
            TeamPlace = 1
        End If
        
        
'       Get Name and Standard data from Mens List using lookup with Race Number

        Worksheets("Mens List").Activate
        For RowNo = 4 To LastMensRow
            If RaceNo = Cells(RowNo, 3).Value Then
                TeamName = Cells(RowNo, 6).Value
                Name = Cells(RowNo, 4).Value
                Std = Cells(RowNo, 5).Value
                                
                ' Load team placing on line in "Mens List"
                Cells(RowNo, 8).Value = TeamPlace
                
                RowNo = LastMensRow + 1
            End If
            
        Next RowNo
        
        Worksheets("New Mens Teams Placings").Activate
        If TeamPlace = 1 Then
            Cells(MensPlace, 1).Value = "Class " & HoldClass
        End If
        MensPlace = MensPlace + 1
        Cells(MensPlace, 1).Value = TeamPlace
        Cells(MensPlace, 2).Value = TeamName
        Cells(MensPlace, 3).Value = TeamNo
        Cells(MensPlace, 4).Value = StdTotal
        Cells(MensPlace, 5).Value = PointsTot
        Cells(MensPlace, 6).Value = HoldClass
        For Cntr = 1 To 3
            Cells(MensPlace, 7).Value = Member_Place(Cntr)
            Cells(MensPlace, 8).Value = Member_RaceNo(Cntr)
            Cells(MensPlace, 9).Value = Member_Name(Cntr)
            Cells(MensPlace, 10).Value = Member_Std(Cntr)
            MensPlace = MensPlace + 1
        Next Cntr

    End If
        
Next RowNo1


'========================================================================================================================
'========================================================================================================================

'========================================================================================================================
'  Print NEW MENS TEAMS PLACINGS
'========================================================================================================================

NewClass = ""

For RowNo = 5 To 300
    Worksheets("New Mens Teams Placings").Activate
    
    Select Case Cells(RowNo, 1).Value
        Case "Class 1"
            RowNo1 = 76
            NewClass = Cells(RowNo, 1).Value
        Case "Class 2"
            RowNo1 = 90
            NewClass = Cells(RowNo, 1).Value
        Case "Class 3"
            RowNo1 = 104
            NewClass = Cells(RowNo, 1).Value
        Case "Class 4"
            RowNo1 = 118
            NewClass = Cells(RowNo, 1).Value
    End Select
        
' Write Team results in Class 1, 2, 3 and 4 on "New Mens Result" worksheet

    If Cells(RowNo, 1).Value = 1 Then
        For Cntr = 1 To 11
            TeamPlace = Cells(RowNo, 1).Value
            If TeamPlace = "" Then
                If Cells(RowNo, 2).Value = Cells(RowNo - 1, 2).Value Then
                    TeamName = ""
                Else
                    TeamName = Cells(RowNo, 2).Value
                End If
                StdTotal = ""
                PointsTot = ""
            Else
                TeamName = Trim(Cells(RowNo, 2).Value)
                StdTotal = Cells(RowNo, 4).Value
                PointsTot = Cells(RowNo, 5).Value
            End If
            Place = Cells(RowNo, 7).Value
            RaceNo = Cells(RowNo, 8).Value
            Name = Cells(RowNo, 9).Value
            Std = Cells(RowNo, 10).Value
            RowNo = RowNo + 1
        
            Worksheets("New Mens Result").Activate
            If NewClass = "" Then
            Else
                Cells(RowNo1 - 2, 1).Value = NewClass
                NewClass = ""
            End If
            Cells(RowNo1, 1).Value = TeamPlace
            Cells(RowNo1, 2).Value = TeamName
            Cells(RowNo1, 3).Value = StdTotal
            Cells(RowNo1, 4).Value = PointsTot
            Cells(RowNo1, 5).Value = Place
            Cells(RowNo1, 6).Value = RaceNo
            Cells(RowNo1, 7).Value = Name
            Cells(RowNo1, 8).Value = Std
            RowNo1 = RowNo1 + 1
            
            Worksheets("New Mens Teams Placings").Activate
        Next Cntr
        
    End If
    
    If RowNo1 = 129 Then
        Exit For
    End If

Next RowNo

'===========================   END of SECTION TO BE COPIED TO "NEW MENS TEAMS PLACINGS"   ==========================================

'===================================================================================================

Application.ScreenUpdating = True

Worksheets("Combined").Activate
Range("a2").Select



End Sub
Sub Clear()
'
' Clear Macro
' Macro recorded 03/07/2004 by Administrator
'
' Keyboard Shortcut: Ctrl+k
'
    Sheets("Mens").Select
    Range("A4:J8").Select
    Selection.ClearContents
    Range("A14:J9999").Select
    Selection.ClearContents
    Range("A4").Select
    
    Sheets("Mens List").Select
    Range("B4:H9999").Select
    Selection.ClearContents
    Range("F37").Select
    
    Sheets("Mens Teams").Select
    Range("B5:R9999").Select
    Selection.ClearContents
    Range("I17").Select
    
    Sheets("Ladies").Select
    Range("B5:J8").Select
    Selection.ClearContents
    
    Range("B14:J9999").Select
    Selection.ClearContents
    Range("I32").Select
    
    Sheets("Ladies List").Select
    Range("A4:I9999").Select
    Selection.ClearContents
    Range("F37").Select
    
    Sheets("Ladies Teams").Select
    Range("B5:P9998").Select
    Selection.ClearContents
    Range("F15").Select
    
    Sheets("Ladies Vets").Select
    Range("B5:H8").Select
    Selection.ClearContents
    Range("B15:H18").Select
    Selection.ClearContents
    Range("B25:H28").Select
    Selection.ClearContents
    Range("B35:H38").Select
    Selection.ClearContents
    Range("B45:H48").Select
    Selection.ClearContents
    Range("B55:H58").Select
    Selection.ClearContents
    Range("B65:H68").Select
    Selection.ClearContents
    Range("B75:H78").Select
    Selection.ClearContents
    Range("A1").Select
    
    Sheets("Mens Vets").Select
    Range("B5:H8").Select
    Selection.ClearContents
    Range("B15:H18").Select
    Selection.ClearContents
    Range("B25:H28").Select
    Selection.ClearContents
    Range("B35:H38").Select
    Selection.ClearContents
    Range("B45:H48").Select
    Selection.ClearContents
    Range("B55:H58").Select
    Selection.ClearContents
    Range("B65:H68").Select
    Selection.ClearContents
    Range("B75:H78").Select
    Selection.ClearContents
    Range("B85:H88").Select
    Selection.ClearContents
    Range("A1").Select
     
    Sheets("New Mens Teams").Select
    Range("B5:S300").Select
    Selection.ClearContents

    Sheets("New Mens Teams Placings").Select
    Range("A4:J300").Select
    Selection.ClearContents
    
    Sheets("Print Mens Result").Select
    Range("B4:H7").Select
    Selection.ClearContents
    Range("B12:H15").Select
    Selection.ClearContents
    Range("B19:H22").Select
    Selection.ClearContents
    Range("B26:H29").Select
    Selection.ClearContents
    Range("B33:H36").Select
    Selection.ClearContents
    Range("B40:H43").Select
    Selection.ClearContents
    Range("B47:H50").Select
    Selection.ClearContents
    Range("B54:H57").Select
    Selection.ClearContents
    Range("B61:H64").Select
    Selection.ClearContents
    Range("B68:H71").Select
    Selection.ClearContents
    Range("B76:H86").Select
    Selection.ClearContents
    Range("B90:H100").Select
    Selection.ClearContents
    Range("B104:H114").Select
    Selection.ClearContents
    Range("B118:H128").Select
    Selection.ClearContents
    Range("A2").Select
    
    Sheets("Print Ladies Result").Select
    Range("B4:H7").Select
    Selection.ClearContents
    Range("B12:H15").Select
    Selection.ClearContents
    Range("B19:H22").Select
    Selection.ClearContents
    Range("B26:H29").Select
    Selection.ClearContents
    Range("B33:H36").Select
    Selection.ClearContents
    Range("B40:H43").Select
    Selection.ClearContents
    Range("B47:H50").Select
    Selection.ClearContents
    Range("B54:H57").Select
    Selection.ClearContents
    Range("B61:H64").Select
    Selection.ClearContents
    Range("B69:H79").Select
    Selection.ClearContents
    Range("A2").Select
        
    Sheets("New Mens Result").Select
    Range("B4:H7").Select
    Selection.ClearContents
    Range("B12:H15").Select
    Selection.ClearContents
    Range("B19:H22").Select
    Selection.ClearContents
    Range("B26:H29").Select
    Selection.ClearContents
    Range("B33:H36").Select
    Selection.ClearContents
    Range("B40:H43").Select
    Selection.ClearContents
    Range("B47:H50").Select
    Selection.ClearContents
    Range("B54:H57").Select
    Selection.ClearContents
    Range("B61:H64").Select
    Selection.ClearContents
    Range("B68:H71").Select
    Selection.ClearContents

    Range("B76:H86").Select
    Selection.ClearContents
    Range("B90:H100").Select
    Selection.ClearContents
    Range("B104:H114").Select
    Selection.ClearContents
    Range("B118:H128").Select
    Selection.ClearContents
    Range("A2").Select
    Sheets("Combined").Select
    Range("N2:Q500").Select
    Selection.ClearContents

    Range("A2").Select

End Sub

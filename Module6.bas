Attribute VB_Name = "Module6"
Sub HomePage()

     'Declare all variables

    Dim ws As Worksheet, curws As Worksheet, shtName As String

    Dim nRow As Long, i As Long, N As Long, x As Long, tmpCount As Long

    Dim cLeft, cTop, cHeight, cWidth, cb As Shape, strMsg As String

    Dim cCnt As Long, cAddy As String, cShade As Long
    
    Dim SheetDescription


     '--------------------------------------------------------

    cShade = 2 '<<== SET BACKGROUND COLOR DESIRED HERE ( 2 is White )

     '--------------------------------------------------------

     'Turn off events and screen flickering.

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False

    nRow = 3: x = 0

     'Check if sheet exists already; direct where to go if not.


    Sheets("HomePage").Activate


     'Set chart sheet varible counter

    tmpCount = ActiveWorkbook.Charts.Count

    If tmpCount > 0 Then tmpCount = 1


'    ActiveSheet.Name = "HomePage"
 
    Sheets("HomePage").Select
    Range("B3:D21").Select
    Selection.ClearContents

    cShade = 37 '<<== SET BACKGROUND COLOR DESIRED HERE (37 is Blue)

'    ActiveSheet.Name = "HomePage"

    With Sheets("HomePage")

        Range("B2:D2").Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 12611584
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        .Range("B2").Value = "Worksheet Name"
        .Range("B2").Font.Bold = True
        .Range("B2").Font.Italic = False
        .Range("B2").Font.Name = "Arial"
        .Range("B2").Font.Size = "14"
        .Range("D2").Value = "Description"
        .Range("D2").Font.Bold = True
        .Range("D2").Font.Name = "Arial"
        .Range("D2").Font.Size = "14"
        .Range("A4").Select

    End With
     'Set variables for loop/iterations

    N = ActiveWorkbook.Sheets.Count + tmpCount

    If x = 1 Then N = N - 1

    For i = 2 To N

' Open each sheet and save description from cell B1 in SheetDescription
        Sheets(i).Activate
        SheetDescription = Cells(1, 2).Value
        If Cells(1, 2).Value = "" Then
            SheetDescription = Cells(1, 1).Value
        End If
        
        With Sheets("HomePage")

                shtName = Sheets(i).Name

                 'Add a hyperlink to A1 of each sheet.

                .Range("B" & nRow).Hyperlinks.Add _
                Anchor:=.Range("B" & nRow), Address:="#'" & _
                shtName & "'!A1", TextToDisplay:=shtName

                .Range("B" & nRow).HorizontalAlignment = xlLeft
                .Range("B" & nRow).Font.Bold = False
                .Range("B" & nRow).Font.Name = "Arial"
                .Range("B" & nRow).Font.Size = "12"

                .Range("D" & nRow).Value = SheetDescription
                .Range("D" & nRow).HorizontalAlignment = xlLeft
                .Range("D" & nRow).Font.Bold = False
                .Range("D" & nRow).Font.Name = "Arial"
                .Range("D" & nRow).Font.Size = "12"
                

            nRow = nRow + 1

        End With

continueLoop:

    Next i

     'Perform some last minute formatting.

    With Sheets("HomePage")

        .Range("D:D").EntireColumn.AutoFit
        

    End With

     'Turn events back on.

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

    Sheets("HomePage").Select
    Range("A1").Select
    
    MsgBox "Complete!" & strMsg, vbInformation, "Complete!"
    

End Sub

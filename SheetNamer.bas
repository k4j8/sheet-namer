Attribute VB_Name = "SheetNamer"
' created by Kyle Johnston on 2014-07-15
' last update: 2015-06-29

Public exists As Boolean

Sub Main()
' This is the only program in the module the user should have direct access to.
' All other subs are activated by a sub or a button. (This is the main.)

' if Sheet_Namer is selected: updates sheet names
' if 1 cell (not Sheet_Namer) is selected: creates Sheet_Namer
' if multiple cells (not Sheet_Namer) are selected: assigns cell values to sheet names


'' BACKUP

' check if Sheet_Namer exists
Dim SheetNamerExists As Boolean
SheetNamerExists = False
For Each ws In Sheets
    If ws.Name = "Sheet_Namer" Then
        SheetNamerExists = True
    End If
Next

' if Sheet_Namer exists and this is not a restore, create a backup
' if Sheet_Namer does not exist, the backup created would be invalid
If ActiveSheet.Name <> "Sheet_Namer_Backup" And SheetNamerExists Then
    Set UserSheet = ActiveSheet
    Call BackupSheetNames
    UserSheet.Activate
End If


'' FILLED CELLS

' if Sheet Name is active, FilledCells is used to select range of sheet names
' otherwise, FilledCells determines whether to create Sheet_Namer or update names directly
Dim FilledCells As Integer
FilledCells = 0

' update FilledCells
If ActiveSheet.Name = "Sheet_Namer" Then
    ' if using Sheet_Namer, select range of filled cells
    ActiveSheet.Range("A1").value = "Sheet_Namer" ' cell A1 must be filled for next line to run
    Dim CheckCell As String ' cell to be checked if empty
    CheckCell = ActiveSheet.Range("A1").value
    While CheckCell <> ""
        FilledCells = FilledCells + 1 ' increase number of non-empty cells
        CheckCell = ActiveSheet.Range("A" & FilledCells + 1) ' try next cell
    Wend
    Worksheets("Sheet_Namer").Range("A1:A" & FilledCells).Select ' select the range of filled cells
Else
    ' if not using Sheet_Namer, determine if single cell or multiple cells are selected
    For Each cell In Selection
        FilledCells = FilledCells + 1 ' increase number of non-empty cells
        If cell = "" And FilledCells > 1 Then
            MsgBox ("Cells must contain values.")
            End
        End If
    Next
End If


'' CREATE Sheet_Namer OR UPDATE NAMES

Dim wsCount As Integer ' populates Sheet_Namer or renames sheets from selection
wsCount = 0

If FilledCells = 1 And ActiveSheet.Name <> "Sheet_Namer" Then
    
    Call CreateSheetNamer
    
Else
    
    Call UpdateSheetNamer
    
End If

End Sub

Sub CreateSheetNamer()

Application.ScreenUpdating = False

' delete Sheet_Namer
For Each ws In Worksheets
    If ws.Name = "Sheet_Namer" Then
        Application.DisplayAlerts = False
        ws.Delete
        Application.DisplayAlerts = True
    End If
Next

' create Sheet_Namer
Sheets.Add Before:=Sheets(1)
ActiveSheet.Name = "Sheet_Namer"
ActiveSheet.Range("A1:A1").Font.Bold = True
ActiveSheet.Range("C1:C1").Font.Bold = True
ActiveSheet.Range("C2:D2").Font.Bold = True
ActiveSheet.Range("C1").value = "Instructions"
ActiveSheet.Range("C2").value = "cells selected"
ActiveSheet.Range("D2").value = "what the macro does"
ActiveSheet.Range("C3").value = "if Sheet_Namer is active..."
ActiveSheet.Range("D3").value = "updates sheet names"
ActiveSheet.Range("C4").value = "if 1 cell (not in Sheet_Namer) is selected..."
ActiveSheet.Range("D4").value = "creates/resets Sheet_Namer"
ActiveSheet.Range("C5").value = "if multiple cells (not in Sheet_Namer) are selected..."
ActiveSheet.Range("D5").value = "assigns cell values to sheet names"
For Each ws In Worksheets
    ' populate Sheet_Namer with existing names
    If ws.Name <> "Sheet_Namer_Backup" Then
        wsCount = wsCount + 1
        Worksheets("Sheet_Namer").Range("A" & wsCount) = ws.Name
    End If
Next
ActiveSheet.Range("A1:A" & wsCount).NumberFormat = "@" ' set sheet names to text type
With ActiveSheet.Range("A1:A" & wsCount).Interior
    ' set background color of sheet names area
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent2
    .TintAndShade = 0.799981688894314
    .PatternTintAndShade = 0
End With
With ActiveSheet.[D14]
    ' create Extra Features title
    .value = "Extra Features"
    .HorizontalAlignment = xlCenter
    .Font.Bold = True
End With
With ActiveSheet.Range("D9:D14").Interior
    ' set background color of extra features area
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent1
    .TintAndShade = 0.599993896298105
    .PatternTintAndShade = 0
End With
With ActiveWorkbook.Sheets("Sheet_Namer").Tab
    ' set background color of Sheet_Namer tab
    .ThemeColor = xlThemeColorAccent6
    .TintAndShade = 0.399975585192419
End With
With Range("C21:C24")
    ' merge cells for instructions
    .Merge
    .HorizontalAlignment = xlGeneral
    .VerticalAlignment = xlTop
    .WrapText = True
    .Orientation = 0
    .AddIndent = False
    .IndentLevel = 0
    .ShrinkToFit = False
    .ReadingOrder = xlContext
    .MergeCells = True
End With
Columns("A").AutoFit
Columns("C").AutoFit
Columns("D").AutoFit

ActiveSheet.Range("C7").value = "You may delete this sheet at any time."

' create hyperlinks
For Each xCell In Worksheets("Sheet_Namer").Range("A1:A" & wsCount) ' range of filled cells
    If xCell.value <> "Sheet_Namer" Then
        ActiveSheet.Hyperlinks.Add Anchor:=Range("B" & xCell.Row), Address:="", SubAddress:= _
            "'" & xCell.value & "'!A1", TextToDisplay:="Activate"
    End If
Next


'' SHEET NAMER BUTTONS

' button size and location
Dim ButtonSize As Range
Set ButtonSize = ActiveSheet.Range("C9:C11")


' create button to run SheetNamer

ActiveSheet.Buttons.Add(ButtonSize.Left, ButtonSize.Top, 90, 30).Select
Selection.OnAction = "SheetNamer.Main"
Selection.Characters.text = "Update Sheet Names"

ActiveSheet.Buttons.Add(ButtonSize.Left + 90 + 5, ButtonSize.Top, 20, 20).Select
Selection.OnAction = "UpdateInstructions"
Selection.Characters.text = "?"


' create button to reset SheetNamer

ActiveSheet.Buttons.Add(ButtonSize.Left + 130, ButtonSize.Top, 90, 30).Select
Selection.OnAction = "CreateSheetNamer"
Selection.Characters.text = "Reset Sheet_Namer"

ActiveSheet.Buttons.Add(ButtonSize.Left + 130 + 90 + 5, ButtonSize.Top, 20, 20).Select
Selection.OnAction = "ResetInstructions"
Selection.Characters.text = "?"


' create button to restore backup
ActiveSheet.Buttons.Add(ButtonSize.Left, ButtonSize.Top + 40, 90, 30).Select
Selection.OnAction = "ShowHideBackup"
Selection.Characters.text = "Show/Hide Backup Sheet Names"

' create button to delete backup
ActiveSheet.Buttons.Add(ButtonSize.Left + 130, ButtonSize.Top + 40, 90, 30).Select
Selection.OnAction = "DeleteBackup"
Selection.Characters.text = "Delete Backup Sheet Names"

' create button to delete sheet
ActiveSheet.Buttons.Add(ButtonSize.Left, ButtonSize.Top + 100, 90 * 2 + 10, 30).Select
Selection.OnAction = "DeleteSheet"
Selection.Characters.text = "Delete this Sheet"


'' EXTRA FEATURE BUTTONS

' create button to select cell A1 in all sheets

ActiveSheet.Buttons.Add(ButtonSize.Left + 258, ButtonSize.Top, 120, 30).Select
Selection.OnAction = "SelectCellA1"
Selection.Characters.text = "Select Cell A1 on All Sheets"

ActiveSheet.Buttons.Add(ButtonSize.Left + 258 + 120 + 5, ButtonSize.Top, 20, 20).Select
Selection.OnAction = "SelectCellA1Instructions"
Selection.Characters.text = "?"


' create button to reset the right-click menu

ActiveSheet.Buttons.Add(ButtonSize.Left + 258, ButtonSize.Top + 40, 120, 30).Select
Selection.OnAction = "ResetRightClickMenuItems"
Selection.Characters.text = "Reset Right Click Menu"

ActiveSheet.Buttons.Add(ButtonSize.Left + 258 + 120 + 5, ButtonSize.Top + 40, 20, 20).Select
Selection.OnAction = "ResetRightClickMenuItemsInstructions"
Selection.Characters.text = "?"


ActiveWindow.DisplayGridlines = False
Worksheets("Sheet_Namer").[a1].Select
Application.ScreenUpdating = True

End Sub

Sub UpdateSheetNamer()
' multiple cells selected or Sheet_Namer is active, so rename sheets

Application.ScreenUpdating = False

' report an error if an apostrophe is used
For Each c1 In Selection
    If InStr(c1.value, "'") <> 0 Then
        MsgBox ("Apostrophes are not allowed in names.")
        End
    End If
Next

' report an error if two names are the same in the list
Dim same As Boolean ' variable same prevents error when comparing cell to itself
For Each c1 In Selection
    same = False
    For Each c2 In Selection
        ' compare with other cells in list
        If StrComp(c1, c2) = 0 Then
            If same Then
                MsgBox ("Duplicate names are not allowed.")
                End
            Else
                same = True
            End If
        End If
    Next
Next

' change name of sheet if name matches
Dim DuplicateNum As Integer ' used if duplicate occurs
DuplicateNum = 1
For Each c1 In Selection
    For Each ws In Worksheets
        ' compare with sheet names
        If StrComp(c1, ws.Name) = 0 Then
            ' if names are the same, change sheet name
            ws.Name = ws.Name & " duplicate " & DuplicateNum
            DuplicateNum = DuplicateNum + 1
        End If
    Next
Next


' apply cell value to sheet names in order

Call CheckBackup
Dim wsAllCount As Integer
If Not exists Then
    wsAllCount = Worksheets.Count
Else
    wsAllCount = Worksheets.Count - 1
End If

For Each cell In Selection
    If wsCount <= wsAllCount Then
        wsCount = wsCount + 1
        If Worksheets(wsCount).Name = "Sheet_Namer_Backup" Then
            wsCount = wsCount + 1
        End If
        If wsCount <= Worksheets.Count Then
            Worksheets(wsCount).Name = cell
        End If
    End If
Next


'' UPDATE Sheet_Namer (IF USING)

If ActiveSheet.Name = "Sheet_Namer" Then
    
    ' add hyperlinks
    For Each xCell In Worksheets("Sheet_Namer").Range("A1:A" & wsAllCount) ' range of filled cells
        If xCell.value <> "Sheet_Namer" Then
            ActiveSheet.Hyperlinks.Add Anchor:=Range("B" & xCell.Row), Address:="", SubAddress:= _
                "'" & xCell.value & "'!A1", TextToDisplay:="Activate"
        End If
    Next
    
    ' change background color of cells
    ActiveSheet.Range("A1:A" & wsAllCount).NumberFormat = "@"
    With ActiveSheet.Range("A1:A" & wsAllCount).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent2
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    Call UpdateStatus("Sheet names updated from list.")
    
    Columns("A").AutoFit
    ActiveSheet.Range("A1").Select
    Application.ScreenUpdating = True
    
End If

End Sub

Sub CheckBackup()
' checks if Sheet_Namer_Backup exists

exists = False
For Each ws In Worksheets
    If ws.Name = "Sheet_Namer_Backup" Then
        exists = True
    End If
Next

End Sub

Sub BackupSheetNames()
' creates Sheet_Namer_Backup to restore previous sheet names


'' CREATE BACKUP SHEET

Call CheckBackup

If Not exists Then
    ' if Sheet_Namer_Backup doesn't exist
    Sheets.Add After:=Sheets(Worksheets("Sheet_Namer").Index)
    ActiveSheet.Name = "Sheet_Namer_Backup"
    Columns("A:E").ColumnWidth = 20
    With ActiveSheet.Tab
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399975585192419
    End With
    
    ' create button to restore names
    Dim ButtonSize As Range
    Dim ColumnNum As Integer
    ColumnNum = 65
    For Each xCell In Worksheets("Sheet_Namer_Backup").Range("A1:E1")
        Set ButtonSize = xCell
        ActiveSheet.Buttons.Add(ButtonSize.Left + 10, ButtonSize.Top, ButtonSize.Width - 20, ButtonSize.Height).Select
        Selection.OnAction = "RestoreNames" & Chr(ColumnNum)
        ColumnNum = ColumnNum + 1
        Selection.Characters.text = "Restore Names"
    Next
    ActiveSheet.Visible = xlVeryHidden
End If


'' PREPARE BACKUP

' copy cells over one column
Worksheets("Sheet_Namer_Backup").Range("A2:D1000").Copy Worksheets("Sheet_Namer_Backup").Range("B2:E1000")


'' RUN BACKUP

Dim SaveCell As Integer ' cell to save name to
SaveCell = 1 ' start saving at A2
For Each ws In Worksheets
    If ws.Name <> "Sheet_Namer_Backup" Then
        SaveCell = SaveCell + 1
        Worksheets("Sheet_Namer_Backup").Range("A" & SaveCell).value = ws.Name
    End If
Next

End Sub

Sub RestoreNamesA()
' restore column A in Sheet_Namer_Backup

Call RestoreNames("A")

End Sub

Sub RestoreNamesB()
' restore column B in Sheet_Namer_Backup

Call RestoreNames("B")

End Sub

Sub RestoreNamesC()
' restore column C in Sheet_Namer_Backup

Call RestoreNames("C")

End Sub

Sub RestoreNamesD()
' restore column D in Sheet_Namer_Backup

Call RestoreNames("D")

End Sub

Sub RestoreNamesE()
' restore column E in Sheet_Namer_Backup

Call RestoreNames("E")

End Sub

Sub RestoreNames(RestoreColumn)
' triggered by any one of restore column buttons

Dim FilledCells As Integer
FilledCells = 0

Dim CheckCell As String
CheckCell = ActiveSheet.Range(RestoreColumn & "2").value ' cell to be checked if empty

While CheckCell <> ""
    FilledCells = FilledCells + 1 ' increase number of non-empty cells
    CheckCell = ActiveSheet.Range(RestoreColumn & "" & FilledCells + 2) ' try next cell
Wend
ActiveSheet.Range(RestoreColumn & "2:" & RestoreColumn & FilledCells + 1).Select ' select the range of filled cells

Call SheetNamer.Main
Call CreateSheetNamer
Call UpdateStatus("Sheet names from column " & RestoreColumn & " restored.")

End Sub

Sub DeleteSheet()

Application.DisplayAlerts = False
ActiveSheet.Delete
Application.DisplayAlerts = True

End Sub

Sub ShowHideBackup()
' shows Sheet_Namer_Backup if hidden, hides Sheet_Namer_Backup if visible

Call CheckBackup

If exists Then
    If Worksheets("Sheet_Namer_Backup").Visible <> True Then
            Worksheets("Sheet_Namer_Backup").Visible = True
            Worksheets("Sheet_Namer_Backup").Activate
            Worksheets("Sheet_Namer_Backup").Range("A2").Select
            Call UpdateStatus("Backup shown.")
    Else
        Worksheets("Sheet_Namer_Backup").Visible = xlVeryHidden
        Call UpdateStatus("Backup hidden.")
    End If
Else
    Call UpdateStatus("Backup does not exist. Click Update Sheet Names to create a new backup.")
End If

End Sub

Sub DeleteBackup()

Application.ScreenUpdating = False

Call CheckBackup

If exists Then
    Application.DisplayAlerts = False
    Worksheets("Sheet_Namer_Backup").Visible = True
    Worksheets("Sheet_Namer_Backup").Delete
    Application.DisplayAlerts = True
    Call UpdateStatus("Deleted backup sheet names.")
Else
    Call UpdateStatus("Backup does not exist. Click Update Sheet Names to create a new backup.")
End If

Application.ScreenUpdating = True

End Sub

Sub SelectCellA1()
' Scrolls to top right corner of every sheet and selects cell A1.

Application.ScreenUpdating = False
Application.EnableEvents = False

Dim sht As Worksheet, csheet As Worksheet
Application.ScreenUpdating = False
Set csheet = ActiveSheet
For Each sht In ActiveWorkbook.Worksheets
  If sht.Visible Then
    sht.Activate
    Range("A1").Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
  End If
Next sht
csheet.Activate

Application.EnableEvents = True
Application.ScreenUpdating = True

Call UpdateStatus("A1 cells selected on all sheets.")

End Sub

Sub ResetRightClickMenuItems()
' resets the right-click menu that appears when right-clicking on a cell

Application.CommandBars("Cell").Reset
Call UpdateStatus("Right-click menu has been reset.")

End Sub

Sub UpdateStatus(status)
' provides user feedback for actions

Worksheets("Sheet_Namer").[C21:C24].ClearContents

Worksheets("Sheet_Namer").[C19].value = status
With Worksheets("Sheet_Namer").[C19:D19].Interior
    .ThemeColor = xlThemeColorAccent6
    .TintAndShade = 0.399975585192419
End With

End Sub

Sub Instructions(text)
' provides instructions for buttons

Worksheets("Sheet_Namer").[C21:C24].value = text
Worksheets("Sheet_Namer").[C21:C24].WrapText = True

End Sub

Sub UpdateInstructions()

Call Instructions("Sets all of the names in the box on this sheet as the names of the sheets in this workbook in order from top to bottom. Same as running this macro while Sheet_Namer is active.")

End Sub

Sub ResetInstructions()

Call Instructions("Resets the Sheet_Namer sheet, filling the box with current sheet names. Same as running this macro on a single cell on a worksheet other than Sheet_Namer.")

End Sub
Sub SelectCellA1Instructions()

Call Instructions("Sets the visible window of each sheet to the upper-left corner.")

End Sub

Sub ResetRightClickMenuItemsInstructions()

Call Instructions("When right-clicking a cell, some menus may contain duplicate options such as 'Insert date.' Click this button to reset the menu.")

End Sub

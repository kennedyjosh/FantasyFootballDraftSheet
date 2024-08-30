Sub ToggleDraftedStatus()

    Dim ws As Worksheet
    Dim playerName As String
    Dim button As Shape
    Dim targetSheetName As String
    Dim targetSheet As Worksheet
    Dim newButtonText As String
    Dim newCellText As String

    ' Define 1-based index of the column containing position
    Dim positionColumn As Integer
    positionColumn = 4

    ' Get the button that was clicked
    Set button = ActiveSheet.Shapes(Application.Caller)

    ' Get the player's name (assuming name is in column B of the button's row)
    playerName = ActiveSheet.Cells(button.TopLeftCell.Row, 2).Value2

    ' Set newButtonText and newCellText
    If button.OLEFormat.Object.Caption = "Draft" Then
        newButtonText = "Drafted"
        newCellText = "d"
    Else
        newButtonText = "Draft"
        newCellText = ""
    End If

    ' Update the ActiveSheet
    button.OLEFormat.Object.Caption = newButtonText
    ActiveSheet.Cells(button.TopLeftCell.Row, 1).Value = newCellText

    ' Determine the target sheets to update
    If ActiveSheet.Name = "Overall" Then
        ' ActiveSheet is Overall; Update the positional sheet
        targetSheetName = ActiveSheet.Cells(button.TopLeftCell.Row, positionColumn).Value2
        Set targetSheet = ThisWorkbook.Sheets(targetSheetName)
        UpdateButtonCaption targetSheet, playerName, newButtonText, newCellText
    Else
        ' ActiveSheet is the positional sheet; Update the Overall sheet and maybe the Flex sheet
        Set targetSheet = ThisWorkbook.Sheets("Overall")
        UpdateButtonCaption targetSheet, playerName, newButtonText, newCellText
    End If
End Sub

Sub UpdateButtonCaption(ws As Worksheet, playerName As String, btnText As String, cellText As String)
    Dim button As Shape
    Dim nameCell As Range
    Dim cell As Range

    ' Find the cell with the player's name
    Set nameCell = ws.Columns("B").Find(playerName, LookIn:=xlValues, LookAt:=xlWhole)

    ' Check if the player's name was found
    If Not nameCell Is Nothing Then
        ' Directly reference the button by its name
        ' Need to subtract 1 to account for no button in the first (header) row
        Set button = ws.Shapes.Item(nameCell.Row - 1)

        ' Update the button caption if found
        If Not button Is Nothing Then
            button.OLEFormat.Object.Caption = btnText
        End If

        ' Get underlying cell
        Set cell = ws.Range("A" & nameCell.Row)

        ' Update text in cell (underneath button) to trigger formatting
        If Not cell Is Nothing Then
            cell.Value = cellText
        End If
    End If
End Sub


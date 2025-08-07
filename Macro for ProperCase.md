```vb
Sub DoProperCase()
'
' DoProperCase Macro
'
' Keyboard Shortcut: Ctrl+Shift+P
'
    On Error Resume Next

    For Each cell In Selection.Cells
        cell.Value = Application.WorksheetFunction.Proper(cell.Value)
    Next
End Sub

```

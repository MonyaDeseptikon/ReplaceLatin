Attribute VB_Name = "Округление"
Sub Округление()
Dim X As Variant, D As Range
Set D = Selection
For Each X In D.Cells
X.Value = Round(X, 2) ' округление до 2х цифр после запятой
Next

End Sub

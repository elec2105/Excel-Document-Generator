Attribute VB_Name = "Module3"
'Option Explicit

Sub Макрос1()
Dim Shp As Object
    For Each Shp In ActiveSheet.Shapes
        If Shp.FormControlType = xlCheckBox Then
            If Shp.AlternativeText <> ActiveSheet.Shapes("Check Box 1").AlternativeText Then
                ActiveSheet.Shapes("Check Box 1").Select
                If Selection.Value = xlOff Then
                    Shp.Select
                    Selection.Value = xlOff
                Else
                    Shp.Select
                    Selection.Value = xlOn
                End If
            End If
        End If
    Next
[a1].Select
End Sub
  



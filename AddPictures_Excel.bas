Attribute VB_Name = "AddPictures"
Option Explicit


Function AddPicturesExcel(ByVal pathfolder_with_pictires_files As String, ByVal range_with_pictures_names As Range)

Dim picture_size As Integer
Dim el As Range
Dim sha As Shape

picture_size = 90

On Error Resume Next
For Each el In range_with_pictures_names

    If ShapeIsThere(ActiveSheet, CStr(el)) = False Then
        ActiveSheet.Shapes.AddPicture(pathfolder_with_pictires_files & el & ".jpg", False, True, _
                el.Offset(0, 1).Left + 100, el.Offset(0, 1).Top + 3, picture_size, picture_size).Name = CStr(el)
    End If

Next

For Each sha In ActiveSheet.Shapes
    Debug.Print sha.Name
Next sha

End Function

Function ShapeIsThere(ByVal Where As Worksheet, ByVal ShapeName As String) As Boolean
  On Error Resume Next
  ShapeIsThere = Not Where.Shapes(ShapeName) Is Nothing
End Function


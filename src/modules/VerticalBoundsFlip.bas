Attribute VB_Name = "VerticalBoundsFlip"
Option Explicit

' Flips two shapes vertically while keeping their combined top and bottom bounds fixed.
' The upper shape is moved to the bottom and the lower shape is moved to the top.
Public Sub Flip_Vertical_Within_Bounds()
    Dim selectedShapes As ShapeRange

    If ActiveWindow.Selection.HasChildShapeRange Then
        If ActiveWindow.Selection.ChildShapeRange.Count <> 2 Then
            MsgBox "You must have two objects selected.", vbCritical
            Exit Sub
        End If
        Set selectedShapes = ActiveWindow.Selection.ChildShapeRange
    Else
        If ActiveWindow.Selection.ShapeRange.Count <> 2 Then
            MsgBox "You must have two objects selected.", vbCritical
            Exit Sub
        End If
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
    End If

    FlipShapesVerticallyWithinBounds selectedShapes
End Sub

Private Sub FlipShapesVerticallyWithinBounds(ByVal selectedShapes As ShapeRange)
    Dim upperShape As Shape
    Dim lowerShape As Shape
    Dim combinedTop As Single
    Dim combinedBottom As Single

    If selectedShapes(1).Top <= selectedShapes(2).Top Then
        Set upperShape = selectedShapes(1)
        Set lowerShape = selectedShapes(2)
    Else
        Set upperShape = selectedShapes(2)
        Set lowerShape = selectedShapes(1)
    End If

    combinedTop = upperShape.Top
    combinedBottom = upperShape.Top + upperShape.Height
    If lowerShape.Top < combinedTop Then combinedTop = lowerShape.Top
    If lowerShape.Top + lowerShape.Height > combinedBottom Then
        combinedBottom = lowerShape.Top + lowerShape.Height
    End If

    ' Set the lower shape first so its top reaches the original upper bound.
    lowerShape.Top = combinedTop
    ' Then set the upper shape so its bottom reaches the original lower bound.
    upperShape.Top = combinedBottom - upperShape.Height
End Sub

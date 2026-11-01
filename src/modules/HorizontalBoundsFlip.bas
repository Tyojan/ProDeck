Attribute VB_Name = "HorizontalBoundsFlip"
Option Explicit

' Flips two shapes horizontally while keeping their combined left and right bounds fixed.
' The left shape is moved to the right and the right shape is moved to the left.
Public Sub Flip_Horizontal_Within_Bounds()
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

    FlipShapesHorizontallyWithinBounds selectedShapes
End Sub

Private Sub FlipShapesHorizontallyWithinBounds(ByVal selectedShapes As ShapeRange)
    Dim leftShape As Shape
    Dim rightShape As Shape
    Dim combinedLeft As Single
    Dim combinedRight As Single

    If selectedShapes(1).Left <= selectedShapes(2).Left Then
        Set leftShape = selectedShapes(1)
        Set rightShape = selectedShapes(2)
    Else
        Set leftShape = selectedShapes(2)
        Set rightShape = selectedShapes(1)
    End If

    combinedLeft = leftShape.Left
    combinedRight = leftShape.Left + leftShape.Width
    If rightShape.Left < combinedLeft Then combinedLeft = rightShape.Left
    If rightShape.Left + rightShape.Width > combinedRight Then
        combinedRight = rightShape.Left + rightShape.Width
    End If

    ' Set the right shape first so its left edge reaches the original left bound.
    rightShape.Left = combinedLeft
    ' Then set the left shape so its right edge reaches the original right bound.
    leftShape.Left = combinedRight - leftShape.Width
End Sub

Attribute VB_Name = "ShapeAdjustmentPainter"
Option Explicit

' Stores the yellow adjustment-handle values copied from one AutoShape.
' The values remain available until PowerPoint closes or another shape is copied.
Private SavedAdjustments() As Single
Private SavedAutoShapeType As MsoAutoShapeType
Private SavedAdjustmentCount As Long

Public Sub CopyShapeAdjustments()
    Dim ReferenceShape As Shape
    Dim AdjustmentIndex As Long

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one reference AutoShape first.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    If ActiveWindow.Selection.ShapeRange.Count <> 1 Then
        MsgBox "Select exactly one reference AutoShape.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    Set ReferenceShape = ActiveWindow.Selection.ShapeRange(1)

    If ReferenceShape.Adjustments.Count = 0 Then
        MsgBox "The selected shape has no yellow adjustment handles to copy.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    SavedAutoShapeType = ReferenceShape.AutoShapeType
    SavedAdjustmentCount = ReferenceShape.Adjustments.Count
    ReDim SavedAdjustments(1 To SavedAdjustmentCount)

    For AdjustmentIndex = 1 To SavedAdjustmentCount
        SavedAdjustments(AdjustmentIndex) = ReferenceShape.Adjustments(AdjustmentIndex)
    Next AdjustmentIndex
End Sub

Public Sub PasteShapeAdjustments()
    Dim TargetShape As Shape
    Dim AdjustmentIndex As Long
    Dim AppliedCount As Long
    Dim SkippedCount As Long

    If SavedAdjustmentCount = 0 Then
        MsgBox "Copy adjustment values from a reference shape first.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one or more target AutoShapes.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    For Each TargetShape In ActiveWindow.Selection.ShapeRange
        If TargetShape.AutoShapeType = SavedAutoShapeType _
           And TargetShape.Adjustments.Count = SavedAdjustmentCount Then
            For AdjustmentIndex = 1 To SavedAdjustmentCount
                TargetShape.Adjustments(AdjustmentIndex) = SavedAdjustments(AdjustmentIndex)
            Next AdjustmentIndex
            AppliedCount = AppliedCount + 1
        Else
            SkippedCount = SkippedCount + 1
        End If
    Next TargetShape

    If SkippedCount > 0 Then
        MsgBox "Applied to " & AppliedCount & " shape(s). Skipped " & SkippedCount & _
               " shape(s) because their AutoShape type differs from the reference.", _
               vbInformation, "ProDeck"
    End If
End Sub

Public Sub Rotate_Right_30(control As IRibbonControl)
    RotateSelectedShapes 30
End Sub

Public Sub Rotate_Right_45(control As IRibbonControl)
    RotateSelectedShapes 45
End Sub

Public Sub Rotate_Right_60(control As IRibbonControl)
    RotateSelectedShapes 60
End Sub

Public Sub Rotate_Left_30(control As IRibbonControl)
    RotateSelectedShapes -30
End Sub

Public Sub Rotate_Left_45(control As IRibbonControl)
    RotateSelectedShapes -45
End Sub

Public Sub Rotate_Left_60(control As IRibbonControl)
    RotateSelectedShapes -60
End Sub

Private Sub RotateSelectedShapes(ByVal degrees As Single)
    Dim selectedShapes As ShapeRange
    Dim shapeIndex As Long

    If ActiveWindow.Selection.Type <> ppSelectionShapes Then
        MsgBox "Select one or more shapes first.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    If ActiveWindow.Selection.HasChildShapeRange Then
        Set selectedShapes = ActiveWindow.Selection.ChildShapeRange
    Else
        Set selectedShapes = ActiveWindow.Selection.ShapeRange
    End If

    For shapeIndex = 1 To selectedShapes.Count
        selectedShapes(shapeIndex).Rotation = selectedShapes(shapeIndex).Rotation + degrees
    Next shapeIndex
End Sub

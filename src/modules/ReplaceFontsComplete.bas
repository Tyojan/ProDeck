Attribute VB_Name = "ReplaceFontsComplete"
Option Explicit

Private gSourceFont As String
Private gFontMenu As CommandBar

' Replaces fonts through the PowerPoint object model.  Unlike the built-in
' replacement dialog, this also visits notes, masters, layouts, tables,
' grouped shapes, master text styles, and bullet fonts.
Public Sub ReplaceFontsComplete(Optional control As Variant)
    Dim pres As Presentation

    On Error Resume Next
    Set pres = ActivePresentation
    On Error GoTo 0

    If pres Is Nothing Then
        MsgBox "Open a presentation before using Complete Font Replace.", vbExclamation, "ProDeck"
        Exit Sub
    End If

    gSourceFont = vbNullString
    ShowSourceFontMenu pres
End Sub

Public Sub SelectCompleteFontSource()
    Dim control As CommandBarControl
    Dim targetFont As String

    On Error Resume Next
    Set control = Application.CommandBars.ActionControl
    gSourceFont = CStr(control.Tag)
    On Error GoTo 0
    If Len(gSourceFont) = 0 Then Exit Sub

    targetFont = Trim$(InputBox( _
        "Replace '" & gSourceFont & "' with this font name:", _
        "ProDeck: All Fonts"))
    If Len(targetFont) = 0 Then Exit Sub

    ApplyCompleteFontReplaceCore gSourceFont, targetFont
End Sub

Private Sub ApplyCompleteFontReplaceCore(ByVal sourceFont As String, ByVal targetFont As String)
    Dim replaceAll As Boolean
    Dim changed As Long
    Dim pres As Presentation

    On Error Resume Next
    Set pres = ActivePresentation
    On Error GoTo 0
    If pres Is Nothing Then Exit Sub

    replaceAll = (sourceFont = "*")
    If replaceAll Then
        If MsgBox( _
            "Replace every font reference with '" & targetFont & "'?" & vbCrLf & vbCrLf & _
            "This makes inherited fonts explicit and also changes bullet fonts.", _
            vbOKCancel + vbExclamation, "ProDeck: Complete Font Replace") <> vbOK Then Exit Sub
    End If

    ' PowerPoint can resolve some theme and embedded-font references that are
    ' not exposed as an individual shape font.  The explicit traversal below
    ' then covers all editable objects and removes remaining inheritance.
    If Not replaceAll Then
        On Error Resume Next
        pres.Fonts.Replace sourceFont, targetFont
        On Error GoTo 0
    End If

    ReplacePresentationFonts pres, sourceFont, targetFont, replaceAll, changed

    MsgBox "Complete font replacement finished." & vbCrLf & _
           "Updated font references: " & CStr(changed), _
           vbInformation, "ProDeck"
End Sub

Private Sub ShowSourceFontMenu(ByVal pres As Presentation)
    Dim button As CommandBarButton
    Dim fontNames As Collection
    Dim fontName As Variant

    On Error Resume Next
    Application.CommandBars("ProDeck source font list").Delete
    On Error GoTo 0

    ' A popup menu is supported by the Ribbon-era CommandBars API, whereas
    ' a docked or floating legacy toolbar is not reliably shown in PowerPoint.
    Set gFontMenu = Application.CommandBars.Add("ProDeck source font list", msoBarPopup, False, True)

    Set fontNames = GetPresentationFontNames(pres)
    For Each fontName In fontNames
        Set button = gFontMenu.Controls.Add(msoControlButton)
        If CStr(fontName) = "*" Then
            button.Caption = "* (All Fonts)"
            button.Tag = "*"
        Else
            button.Caption = CStr(fontName)
            button.Tag = CStr(fontName)
        End If
        button.OnAction = "SelectCompleteFontSource"
    Next fontName

    gFontMenu.ShowPopup
End Sub

Private Function GetPresentationFontNames(ByVal pres As Presentation) As Collection
    Dim names As New Collection
    Dim fontItem As Object

    AddUniqueFontName names, "*" ' Explicitly chooses full replacement.

    On Error Resume Next
    For Each fontItem In pres.Fonts
        AddUniqueFontName names, CStr(fontItem.Name)
    Next fontItem
    On Error GoTo 0

    ' Pres.Fonts is useful but can omit inherited or otherwise unreported
    ' entries.  Inspect the same scope as the complete replacement, so the
    ' source-font list makes masters, notes, bullets, tables, and groups
    ' discoverable before the user chooses a replacement.
    CollectPresentationFontNames pres, names

    Set GetPresentationFontNames = names
End Function

Private Sub AddUniqueFontName(ByVal names As Collection, ByVal fontName As String)
    Dim key As String

    fontName = Trim$(fontName)
    If Len(fontName) = 0 Then Exit Sub
    key = LCase$(fontName)

    On Error Resume Next
    names.Add fontName, key
    On Error GoTo 0
End Sub

Private Sub CollectPresentationFontNames(ByVal pres As Presentation, ByVal names As Collection)
    Dim slideObj As Slide
    Dim designObj As Design
    Dim layoutObj As CustomLayout

    For Each slideObj In pres.Slides
        CollectShapeCollectionFontNames slideObj.Shapes, names
        On Error Resume Next
        CollectShapeCollectionFontNames slideObj.NotesPage.Shapes, names
        On Error GoTo 0
    Next slideObj

    For Each designObj In pres.Designs
        CollectMasterFontNames designObj.SlideMaster, names
        For Each layoutObj In designObj.SlideMaster.CustomLayouts
            CollectShapeCollectionFontNames layoutObj.Shapes, names
        Next layoutObj
    Next designObj

    On Error Resume Next
    CollectMasterFontNames pres.NotesMaster, names
    CollectMasterFontNames pres.HandoutMaster, names
    On Error GoTo 0
End Sub

Private Sub CollectMasterFontNames(ByVal masterObj As Object, ByVal names As Collection)
    Dim styleIndex As Long
    Dim levelIndex As Long
    Dim styleLevel As Object

    If masterObj Is Nothing Then Exit Sub
    CollectShapeCollectionFontNames masterObj.Shapes, names

    For styleIndex = ppTitleStyle To ppDefaultStyle
        For levelIndex = 1 To 5
            Set styleLevel = Nothing
            On Error Resume Next
            Set styleLevel = masterObj.TextStyles(styleIndex).Levels(levelIndex)
            On Error GoTo 0
            If Not styleLevel Is Nothing Then
                CollectFontObjectNames styleLevel.Font, names
                On Error Resume Next
                CollectFontObjectNames styleLevel.ParagraphFormat.Bullet.Font, names
                On Error GoTo 0
            End If
        Next levelIndex
    Next styleIndex
End Sub

Private Sub CollectShapeCollectionFontNames(ByVal shapes As Object, ByVal names As Collection)
    Dim shp As Shape
    If shapes Is Nothing Then Exit Sub
    For Each shp In shapes
        CollectShapeFontNames shp, names
    Next shp
End Sub

Private Sub CollectShapeFontNames(ByVal shp As Shape, ByVal names As Collection)
    Dim itemIndex As Long
    Dim rowIndex As Long
    Dim columnIndex As Long

    If shp.Type = msoGroup Then
        For itemIndex = 1 To shp.GroupItems.Count
            CollectShapeFontNames shp.GroupItems(itemIndex), names
        Next itemIndex
    End If

    If ShapeHasTextFrame(shp) Then CollectTextRangeFontNames shp.TextFrame.TextRange, names

    On Error Resume Next
    AddUniqueFontName names, CStr(shp.TextEffect.FontName)
    If shp.HasTable Then
        For rowIndex = 1 To shp.Table.Rows.Count
            For columnIndex = 1 To shp.Table.Columns.Count
                CollectShapeFontNames shp.Table.Cell(rowIndex, columnIndex).Shape, names
            Next columnIndex
        Next rowIndex
    End If
    On Error GoTo 0
End Sub

Private Sub CollectTextRangeFontNames(ByVal textRange As TextRange, ByVal names As Collection)
    Dim paragraphIndex As Long
    Dim paragraphCount As Long
    Dim paragraphRange As TextRange

    CollectFontObjectNames textRange.Font, names
    On Error Resume Next
    paragraphCount = textRange.Paragraphs.Count
    On Error GoTo 0
    For paragraphIndex = 1 To paragraphCount
        Set paragraphRange = Nothing
        On Error Resume Next
        Set paragraphRange = textRange.Paragraphs(paragraphIndex, 1)
        On Error GoTo 0
        If Not paragraphRange Is Nothing Then
            On Error Resume Next
            CollectFontObjectNames paragraphRange.ParagraphFormat.Bullet.Font, names
            On Error GoTo 0
        End If
    Next paragraphIndex
End Sub

Private Sub CollectFontObjectNames(ByVal fontObj As Object, ByVal names As Collection)
    If fontObj Is Nothing Then Exit Sub
    CollectFontPropertyName fontObj, "Name", names
    CollectFontPropertyName fontObj, "NameAscii", names
    CollectFontPropertyName fontObj, "NameFarEast", names
    CollectFontPropertyName fontObj, "NameComplexScript", names
    CollectFontPropertyName fontObj, "NameOther", names
End Sub

Private Sub CollectFontPropertyName(ByVal fontObj As Object, ByVal propertyName As String, ByVal names As Collection)
    Dim fontName As String
    On Error Resume Next
    fontName = CStr(CallByName(fontObj, propertyName, VbGet))
    If Err.Number = 0 Then AddUniqueFontName names, fontName
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub ReplacePresentationFonts(ByVal pres As Presentation, ByVal sourceFont As String, _
                                     ByVal targetFont As String, ByVal replaceAll As Boolean, _
                                     ByRef changed As Long)
    Dim slideObj As Slide
    Dim designObj As Design
    Dim layoutObj As CustomLayout
    Dim masterCount As Long

    For Each slideObj In pres.Slides
        ReplaceShapesInCollection slideObj.Shapes, sourceFont, targetFont, replaceAll, changed
        On Error Resume Next
        ReplaceShapesInCollection slideObj.NotesPage.Shapes, sourceFont, targetFont, replaceAll, changed
        On Error GoTo 0
    Next slideObj

    For Each designObj In pres.Designs
        masterCount = masterCount + 1
        ReplaceMasterFonts designObj.SlideMaster, sourceFont, targetFont, replaceAll, changed
        For Each layoutObj In designObj.SlideMaster.CustomLayouts
            ReplaceShapesInCollection layoutObj.Shapes, sourceFont, targetFont, replaceAll, changed
        Next layoutObj
    Next designObj

    ' Older files occasionally do not surface their only master in Designs.
    If masterCount = 0 Then
        On Error Resume Next
        ReplaceMasterFonts pres.SlideMaster, sourceFont, targetFont, replaceAll, changed
        On Error GoTo 0
    End If

    On Error Resume Next
    ReplaceMasterFonts pres.NotesMaster, sourceFont, targetFont, replaceAll, changed
    ReplaceMasterFonts pres.HandoutMaster, sourceFont, targetFont, replaceAll, changed
    On Error GoTo 0
End Sub

Private Sub ReplaceMasterFonts(ByVal masterObj As Object, ByVal sourceFont As String, _
                               ByVal targetFont As String, ByVal replaceAll As Boolean, _
                               ByRef changed As Long)
    Dim styleIndex As Long
    Dim levelIndex As Long
    Dim styleLevel As Object

    If masterObj Is Nothing Then Exit Sub

    ReplaceShapesInCollection masterObj.Shapes, sourceFont, targetFont, replaceAll, changed

    ' Title, body, and other text styles each have five paragraph levels.
    For styleIndex = ppTitleStyle To ppDefaultStyle
        For levelIndex = 1 To 5
            Set styleLevel = Nothing
            On Error Resume Next
            Set styleLevel = masterObj.TextStyles(styleIndex).Levels(levelIndex)
            On Error GoTo 0
            If Not styleLevel Is Nothing Then
                ReplaceFontObject styleLevel.Font, sourceFont, targetFont, replaceAll, changed
                On Error Resume Next
                ReplaceFontObject styleLevel.ParagraphFormat.Bullet.Font, sourceFont, targetFont, replaceAll, changed
                On Error GoTo 0
            End If
        Next levelIndex
    Next styleIndex
End Sub

Private Sub ReplaceShapesInCollection(ByVal shapes As Object, ByVal sourceFont As String, _
                                      ByVal targetFont As String, ByVal replaceAll As Boolean, _
                                      ByRef changed As Long)
    Dim shp As Shape

    If shapes Is Nothing Then Exit Sub
    For Each shp In shapes
        ReplaceShapeFonts shp, sourceFont, targetFont, replaceAll, changed
    Next shp
End Sub

Private Sub ReplaceShapeFonts(ByVal shp As Shape, ByVal sourceFont As String, _
                              ByVal targetFont As String, ByVal replaceAll As Boolean, _
                              ByRef changed As Long)
    Dim itemIndex As Long
    Dim rowIndex As Long
    Dim columnIndex As Long

    If shp.Type = msoGroup Then
        For itemIndex = 1 To shp.GroupItems.Count
            ReplaceShapeFonts shp.GroupItems(itemIndex), sourceFont, targetFont, replaceAll, changed
        Next itemIndex
    End If

    If ShapeHasTextFrame(shp) Then
        ReplaceTextRangeFonts shp.TextFrame.TextRange, sourceFont, targetFont, replaceAll, changed
    End If

    ' TextEffect is used by legacy WordArt shapes.
    On Error Resume Next
    If shp.TextEffect.Text <> "" Then
        If FontMatches(shp.TextEffect.FontName, sourceFont, replaceAll) Then
            shp.TextEffect.FontName = targetFont
            changed = changed + 1
        End If
    End If
    On Error GoTo 0

    On Error Resume Next
    If shp.HasTable Then
        For rowIndex = 1 To shp.Table.Rows.Count
            For columnIndex = 1 To shp.Table.Columns.Count
                ReplaceShapeFonts shp.Table.Cell(rowIndex, columnIndex).Shape, sourceFont, targetFont, replaceAll, changed
            Next columnIndex
        Next rowIndex
    End If
    On Error GoTo 0
End Sub

Private Sub ReplaceTextRangeFonts(ByVal textRange As TextRange, ByVal sourceFont As String, _
                                  ByVal targetFont As String, ByVal replaceAll As Boolean, _
                                  ByRef changed As Long)
    Dim paragraphIndex As Long
    Dim paragraphCount As Long
    Dim paragraphRange As TextRange

    ReplaceFontObject textRange.Font, sourceFont, targetFont, replaceAll, changed

    On Error Resume Next
    paragraphCount = textRange.Paragraphs.Count
    On Error GoTo 0

    For paragraphIndex = 1 To paragraphCount
        Set paragraphRange = Nothing
        On Error Resume Next
        Set paragraphRange = textRange.Paragraphs(paragraphIndex, 1)
        On Error GoTo 0
        If Not paragraphRange Is Nothing Then
            ReplaceFontObject paragraphRange.ParagraphFormat.Bullet.Font, sourceFont, targetFont, replaceAll, changed
        End If
    Next paragraphIndex
End Sub

Private Sub ReplaceFontObject(ByVal fontObj As Object, ByVal sourceFont As String, _
                              ByVal targetFont As String, ByVal replaceAll As Boolean, _
                              ByRef changed As Long)
    If fontObj Is Nothing Then Exit Sub

    ReplaceFontProperty fontObj, "Name", sourceFont, targetFont, replaceAll, changed
    ReplaceFontProperty fontObj, "NameAscii", sourceFont, targetFont, replaceAll, changed
    ReplaceFontProperty fontObj, "NameFarEast", sourceFont, targetFont, replaceAll, changed
    ReplaceFontProperty fontObj, "NameComplexScript", sourceFont, targetFont, replaceAll, changed
    ReplaceFontProperty fontObj, "NameOther", sourceFont, targetFont, replaceAll, changed
End Sub

Private Sub ReplaceFontProperty(ByVal fontObj As Object, ByVal propertyName As String, _
                                ByVal sourceFont As String, ByVal targetFont As String, _
                                ByVal replaceAll As Boolean, ByRef changed As Long)
    Dim currentFont As String

    On Error Resume Next
    currentFont = CStr(CallByName(fontObj, propertyName, VbGet))
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    If FontMatches(currentFont, sourceFont, replaceAll) Then
        On Error Resume Next
        CallByName fontObj, propertyName, VbLet, targetFont
        If Err.Number = 0 Then changed = changed + 1
        Err.Clear
        On Error GoTo 0
    End If
End Sub

Private Function FontMatches(ByVal currentFont As String, ByVal sourceFont As String, _
                             ByVal replaceAll As Boolean) As Boolean
    If replaceAll Then
        FontMatches = (Len(Trim$(currentFont)) > 0)
    Else
        FontMatches = (StrComp(Trim$(currentFont), Trim$(sourceFont), vbTextCompare) = 0)
    End If
End Function

Private Function ShapeHasTextFrame(ByVal shp As Shape) As Boolean
    On Error Resume Next
    ShapeHasTextFrame = (shp.HasTextFrame = msoTrue)
    On Error GoTo 0
End Function

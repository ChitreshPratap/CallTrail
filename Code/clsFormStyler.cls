VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormStyler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsFormStyler
' Positions and styles controls that ALREADY EXIST on a page.

' You place the controls in the VBE designer - anywhere, any size, it
' doesn't matter - and this class arranges them into a consistent grid
' and applies uniform fonts, colours and borders at run time.

' Controls are looked up BY NAME from the page, so your Arrange routine
' reads as a straight top-to-bottom description of the screen.
'
' NOTE: this is the class to use when the controls already exist.
' clsFormLayout (separate file) CREATES controls from nothing - use
' that one only if you want a page with an empty designer Page.
' Don't use both on the same page.
'
' All measurements are in POINTS, the UserForm's native unit.
' =====================================================================

Private m_container As Object       ' Page, Frame or UserForm
Private m_pageTitle As String

' Layout state
Private m_left As Single
Private m_top As Single
Private m_startTop As Single
Private m_btnLeft As Single
Private m_inButtonRow As Boolean

' --- Tunables: set these after Init, before arranging ---
Public LabelWidth As Single
Public fieldWidth As Single
Public RowHeight As Single
Public RowGap As Single
Public SectionGap As Single
Public FontName As String
Public FontSize As Single
Public HeaderColor As Long
Public NoteColor As Long
Public EditableBackColor As Long
Public ReadOnlyBackColor As Long


Private m_scale As Single

' Width in points of REF_TEXT in Segoe UI 9pt, measured on a standard
' 96 DPI Windows machine with the font installed. Calibrate compares the
' local measurement against this.
'
' To re-baseline on your own machine: run Calibrate, then read
' MeasuredReferenceWidth and put that number here.

Private Const REF_TEXT As String = "Provider Name: ABCDEFGHIJ 0123456789"
Private Const REF_WIDTH As Single = 171

Private m_measuredRef As Single


' Tunables are scaled once, here. Explicit width/height arguments passed
' to individual methods are scaled at the point of use via S().
' Scales the tunables AND the starting cursor, which Init set before the
' scale was known. Guarded so a second call can't compound the scaling.
Private m_scaleApplied As Boolean


Private Sub Class_Initialize()
    
    LabelWidth = 105
    fieldWidth = 230
    RowHeight = 18
    RowGap = 5
    SectionGap = 14
    FontName = "Segoe UI"
    FontSize = 9
    HeaderColor = RGB(31, 78, 120)
    NoteColor = RGB(90, 90, 90)
    EditableBackColor = &H80000005      ' window white
    ReadOnlyBackColor = &H8000000F      ' button-face grey
    m_scale = 1
    
End Sub

' =====================================================================
' CROSS-MACHINE CONSISTENCY
'
' The usual reason a form looks right on your PC and wrong on someone
' else's is FONT SUBSTITUTION. If the font you specified isn't installed,
' Windows quietly swaps in another one with different metrics. Same point
' size, different actual width - so captions clip and fields misalign,
' even though every coordinate in your code is identical.
'
' Calibrate measures how wide a known string ACTUALLY renders on this
' machine, compares it to how wide it rendered on yours, and scales the
' layout by the difference. It corrects for substitution and for any
' font-rendering difference between Office versions in one number.
'
' What this does NOT fix, honestly:
'   - Display scaling (125%/150%). Current Office builds scale UserForms
'     themselves; older ones don't. See the Zoom note in the guide.
'   - Screen resolution. If your form is 780pt wide and their laptop is
'     1366x768, no scaling helps - design for the smallest screen.
' =====================================================================


Public Property Get ScaleFactor() As Single
    ScaleFactor = m_scale
End Property

' ---------------------------------------------------------------------
' Apply a known scale - e.g. the one the shell already measured.
'
' Call this AFTER Init and AFTER setting the tunables (LabelWidth,
' FieldWidth, RowHeight, RowGap), because it scales whatever values they
' currently hold. Setting a tunable afterwards would put an unscaled
' number back in, which is exactly the bug this ordering prevents.
'
'   ui.Init m_pg, "My Tab"
'   ui.LabelWidth = 110
'   ui.FieldWidth = 250
'   ui.UseScale m_ctx.UIScale      <- last, before any Row/Header call
' ---------------------------------------------------------------------
Public Sub UseScale(ByVal factor As Single)
    If factor <= 0 Then Exit Sub
    m_scale = factor
    ApplyScaleToTunables
End Sub

Public Property Get MeasuredReferenceWidth() As Single
    MeasuredReferenceWidth = m_measuredRef
End Property

' ---------------------------------------------------------------------
' Measures the local font and sets the scale. Call ONCE, right after
' Init, before arranging anything.
'
' Creates a hidden AutoSize label, measures it, deletes it. Clamped to
' 0.8-1.4 because a wilder number means something unexpected happened
' and blindly trusting it would produce a worse layout than doing
' nothing.
' ---------------------------------------------------------------------
Public Sub Calibrate()
    Dim lbl As MSForms.Label
    Dim tempName As String
    Dim ratio As Single

    tempName = "lblCalibrate_" & Format$(Now, "hhnnss")

    On Error GoTo GiveUp

    Set lbl = m_container.Controls.Add("Forms.Label.1", tempName, False)
    With lbl
        .Font.Name = FontName
        .Font.Size = FontSize
        .AutoSize = True            ' width now reflects real rendered metrics
        .caption = REF_TEXT
    End With

    m_measuredRef = lbl.width
    m_container.Controls.Remove tempName

    If m_measuredRef <= 0 Then GoTo GiveUp

    ratio = m_measuredRef / REF_WIDTH
    If ratio < 0.8 Then ratio = 0.8
    If ratio > 1.4 Then ratio = 1.4

    UseScale ratio
    Exit Sub

GiveUp:
    ' Calibration is a nice-to-have. If anything goes wrong, fall back to
    ' unscaled layout rather than failing to draw the form at all.
    On Error Resume Next
    m_container.Controls.Remove tempName
    On Error GoTo 0
    m_scale = 1
End Sub


Private Sub ApplyScaleToTunables()
    If m_scaleApplied Then Exit Sub
    m_scaleApplied = True

    LabelWidth = LabelWidth * m_scale
    fieldWidth = fieldWidth * m_scale
    RowHeight = RowHeight * m_scale
    RowGap = RowGap * m_scale
    SectionGap = SectionGap * m_scale

    m_left = m_left * m_scale
    m_top = m_top * m_scale
    m_startTop = m_startTop * m_scale
End Sub

' Scale a single measurement
Private Function S(ByVal v As Single) As Single
    S = v * m_scale
End Function

Public Sub Init(container As Object, ByVal pageTitle As String, _
                 Optional ByVal startLeft As Single = 14, _
                 Optional ByVal startTop As Single = 12)
    Set m_container = container
    m_pageTitle = pageTitle
    m_left = startLeft          ' scaled later by UseScale/Calibrate
    m_top = startTop
    m_startTop = startTop
End Sub

' =====================================================================
' Position control
' =====================================================================
Public Property Get CurrentTop() As Single
    CurrentTop = m_top
End Property
Public Property Let CurrentTop(ByVal v As Single)
    m_top = v
End Property

Public Property Get currentLeft() As Single
    currentLeft = m_left
End Property

' Start a second column - e.g. a detail panel to the right of a list
Public Sub NewColumn(ByVal newLeft As Single, Optional ByVal resetTop As Boolean = True)
    m_left = S(newLeft)
    If resetTop Then m_top = m_startTop
End Sub

Public Sub Gap(Optional ByVal points As Single = -1)
    If points < 0 Then points = SectionGap Else points = S(points)
    m_top = m_top + points
End Sub

' =====================================================================
' Arranging
' =====================================================================

' ---------------------------------------------------------------------
' A bold section heading spanning the label+field block.
' ---------------------------------------------------------------------
Public Sub Header(ByVal labelName As String, Optional ByVal caption As String = "")
    Dim lbl As MSForms.Label
    Set lbl = Resolve(labelName, False)

    With lbl
        .Left = m_left
        .Top = m_top
        .width = LabelWidth + fieldWidth
        .height = RowHeight
        If caption <> "" Then .caption = caption
        .Font.Name = FontName
        .Font.Size = FontSize + 1
        .Font.Bold = True
        .foreColor = HeaderColor
    End With

    m_top = m_top + RowHeight + RowGap
End Sub

' ---------------------------------------------------------------------
' One row: caption label on the left, input control on the right.
'
' captionName may be "" if the field has no caption label.
' Handles TextBox, ComboBox, ListBox, CheckBox and Label fields, and
' styles each appropriately for its type.
' ---------------------------------------------------------------------
Public Sub Row(ByVal captionName As String, ByVal fieldName As String, _
                Optional ByVal caption As String = "", _
                Optional ByVal fieldWidth As Single = 0, _
                Optional ByVal fieldHeight As Single = 0, _
                Optional ByVal readOnlyLook As Boolean = False)
    Dim fld As MSForms.control
    Dim w As Single, h As Single

    Set fld = Resolve(fieldName, False)
    w = IIf(fieldWidth > 0, S(fieldWidth), Me.fieldWidth)
    h = IIf(fieldHeight > 0, S(fieldHeight), RowHeight)

    PlaceCaption captionName, caption, h

    With fld
        .Left = m_left + LabelWidth
        .Top = m_top
        .width = w
        .height = h
    End With

    StyleField fld, (fieldHeight > RowHeight), readOnlyLook

    m_top = m_top + h + RowGap
End Sub

' ---------------------------------------------------------------------
' A read-only value shown as a Label rather than a box - visually
' obvious that there's nothing to type into.
' ---------------------------------------------------------------------
Public Sub ValueRow(ByVal captionName As String, ByVal valueLabelName As String, _
                     Optional ByVal caption As String = "", _
                     Optional ByVal fieldWidth As Single = 0, _
                     Optional ByVal fieldHeight As Single = 0)
    Dim lbl As MSForms.Label
    Dim w As Single, h As Single

    Set lbl = Resolve(valueLabelName, False)
    w = IIf(fieldWidth > 0, S(fieldWidth), Me.fieldWidth)
    h = IIf(fieldHeight > 0, S(fieldHeight), RowHeight)

    PlaceCaption captionName, caption, h

    With lbl
        .Left = m_left + LabelWidth
        .Top = m_top
        .width = w
        .height = h
        .Font.Name = FontName
        .Font.Size = FontSize
        .Font.Bold = True
        .WordWrap = (h > RowHeight)
    End With

    m_top = m_top + h + RowGap
End Sub

' ---------------------------------------------------------------------
' A control spanning the full block with no caption beside it -
' typically a ListBox.
' ---------------------------------------------------------------------
Public Sub FullWidth(ByVal ctrlName As String, ByVal width As Single, ByVal height As Single)
    Dim ctl As MSForms.control
    Dim w As Single, h As Single
    Set ctl = Resolve(ctrlName, False)
    w = S(width): h = S(height)

    With ctl
        .Left = m_left
        .Top = m_top
        .width = w
        .height = h
    End With

    StyleField ctl, False, False
    m_top = m_top + h + RowGap
End Sub

' ---------------------------------------------------------------------
' A note / hint label spanning the block.
' ---------------------------------------------------------------------
Public Sub Note(ByVal labelName As String, Optional ByVal caption As String = "", _
                 Optional ByVal height As Single = 0, Optional ByVal width As Single = 0)
    Dim lbl As MSForms.Label
    Dim h As Single, w As Single

    Set lbl = Resolve(labelName, True)
    If lbl Is Nothing Then Exit Sub     ' notes are allowed to be absent

    h = IIf(height > 0, S(height), RowHeight)
    w = IIf(width > 0, S(width), LabelWidth + fieldWidth)

    With lbl
        .Left = m_left
        .Top = m_top
        .width = w
        .height = h
        If caption <> "" Then .caption = caption
        .Font.Name = FontName
        .Font.Size = FontSize
        .foreColor = NoteColor
        .WordWrap = (h > RowHeight)
    End With

    m_top = m_top + h + RowGap
End Sub

' ---------------------------------------------------------------------
' Buttons laid out left-to-right on one row.
'   BeginButtonRow
'   Button "cmdSave", "Save", 90
'   Button "cmdClear", "Clear", 70
'   EndButtonRow
' ---------------------------------------------------------------------
Public Sub BeginButtonRow(Optional ByVal indent As Single = -1)
    If indent < 0 Then indent = LabelWidth Else indent = S(indent)
    m_btnLeft = m_left + indent
    m_inButtonRow = True
End Sub

Public Sub Button(ByVal ctrlName As String, Optional ByVal caption As String = "", _
                   Optional ByVal width As Single = 80, Optional ByVal gapAfter As Single = 6)
    Dim btn As MSForms.CommandButton
    Set btn = Resolve(ctrlName, False)

    If Not m_inButtonRow Then BeginButtonRow

    With btn
        .Left = m_btnLeft
        .Top = m_top
        .width = S(width)
        .height = RowHeight + S(4)
        If caption <> "" Then .caption = caption
        .Font.Name = FontName
        .Font.Size = FontSize
    End With

    m_btnLeft = m_btnLeft + S(width) + S(gapAfter)
End Sub

Public Sub EndButtonRow()
    m_top = m_top + RowHeight + S(4) + RowGap
    m_inButtonRow = False
End Sub

' ---------------------------------------------------------------------
' Style a control without moving it - for anything you've positioned by
' hand but still want to match the app's look.
' ---------------------------------------------------------------------
Public Sub StyleOnly(ByVal ctrlName As String, Optional ByVal readOnlyLook As Boolean = False)
    
    Dim ctl As MSForms.control
    Set ctl = Resolve(ctrlName, True)
    If ctl Is Nothing Then Exit Sub
    StyleField ctl, False, readOnlyLook
End Sub

' =====================================================================
' SHELL LAYOUT
' These work differently from the row-stacking methods above. A shell
' isn't a stack of rows - it's one big control filling the window with a
' status bar pinned to the bottom. So these size and anchor rather than
' advancing a cursor, and they ignore CurrentTop entirely.
'
' Call these on the UserForm itself (pass Me as the container), not on
' a Page.
' =====================================================================

' ---------------------------------------------------------------------
' Sizes the form and centres it on screen.
'
' UserForm.Width/Height include the window chrome, so the usable area
' (InsideWidth/InsideHeight) is smaller. Everything below positions
' against the INSIDE measurements, which is why the form is sized first.
' ---------------------------------------------------------------------
Public Sub SizeForm(frm As Object, ByVal totalWidth As Single, ByVal totalHeight As Single, _
                     Optional ByVal caption As String = "")
    On Error Resume Next
    With frm
        .width = S(totalWidth)
        .height = S(totalHeight)
        If caption <> "" Then .caption = caption
        .StartUpPosition = 0        ' manual, so we can centre it ourselves
        .Left = (Application.width - .width) / 2
        .Top = (Application.height - .height) / 2
        .Font.Name = FontName
        .Font.Size = FontSize
    End With
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
' Makes a control fill the container, leaving margins and reserving
' space at the bottom for a status bar / button row.
' ---------------------------------------------------------------------
Public Sub Fill(ByVal ctrlName As String, _
                 Optional ByVal margin As Single = 6, _
                 Optional ByVal reserveBottom As Single = 30, _
                 Optional ByVal reserveTop As Single = 0)
    Dim ctl As MSForms.control
    Set ctl = Resolve(ctrlName, False)

    With ctl
        .Left = S(margin)
        .Top = S(margin) + S(reserveTop)
        .width = ContainerWidth() - (S(margin) * 2)
        .height = ContainerHeight() - (S(margin) * 2) - S(reserveBottom) - S(reserveTop)
    End With
End Sub

' ---------------------------------------------------------------------
' Pins a control to the bottom-left - the status bar.
' ---------------------------------------------------------------------
Public Sub AnchorBottomLeft(ByVal ctrlName As String, ByVal width As Single, _
                             Optional ByVal height As Single = 0, _
                             Optional ByVal margin As Single = 8)
    Dim ctl As MSForms.control
    Dim h As Single

    Set ctl = Resolve(ctrlName, False)
    h = IIf(height > 0, S(height), RowHeight)

    With ctl
        .Left = S(margin)
        .width = S(width)
        .height = h
        .Top = ContainerHeight() - h - S(margin)
    End With
End Sub

' ---------------------------------------------------------------------
' Pins a control to the bottom-right - the Close button.
' ---------------------------------------------------------------------
Public Sub AnchorBottomRight(ByVal ctrlName As String, ByVal width As Single, _
                              Optional ByVal height As Single = 0, _
                              Optional ByVal margin As Single = 8)
    Dim ctl As MSForms.control
    Dim h As Single

    Set ctl = Resolve(ctrlName, False)
    h = IIf(height > 0, S(height), RowHeight + S(4))

    With ctl
        .width = S(width)
        .height = h
        .Left = ContainerWidth() - S(width) - S(margin)
        .Top = ContainerHeight() - h - S(margin)
    End With
End Sub

' ---------------------------------------------------------------------
' Status bar styling - subdued, so it reads as chrome rather than as
' content competing with the tab the user is actually looking at.
' ---------------------------------------------------------------------
Public Sub StyleStatusBar(ByVal ctrlName As String)
    Dim lbl As MSForms.Label
    Set lbl = Resolve(ctrlName, True)
    If lbl Is Nothing Then Exit Sub

    With lbl
        .Font.Name = FontName
        .Font.Size = FontSize
        .foreColor = NoteColor
        .TextAlign = fmTextAlignLeft
        .caption = ""
    End With
End Sub

' ---------------------------------------------------------------------
' MultiPage styling. Tab captions are set by the page controllers, so
' this only handles font and tab appearance.
' ---------------------------------------------------------------------
Public Sub StyleMultiPage(ByVal ctrlName As String, _
                           Optional ByVal tabFixedWidth As Single = 0)
    Dim mp As MSForms.MultiPage
    Set mp = Resolve(ctrlName, False)

    On Error Resume Next
    With mp
        .Font.Name = FontName
        .Font.Size = FontSize
        .Style = fmTabStyleTabs
        If tabFixedWidth > 0 Then
            ' Equal-width tabs look deliberate; auto-width tabs jump
            ' around as captions change length.
            .tabFixedWidth = tabFixedWidth
        End If
    End With
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------
' A title banner across the top of the form. The label must already
' exist; skipped silently if it doesn't.
' ---------------------------------------------------------------------
Public Sub TitleBar(ByVal ctrlName As String, ByVal caption As String, _
                     Optional ByVal height As Single = 22, _
                     Optional ByVal margin As Single = 6)
    Dim lbl As MSForms.Label
    Set lbl = Resolve(ctrlName, True)
    If lbl Is Nothing Then Exit Sub

    With lbl
        .Left = S(margin)
        .Top = S(margin)
        .width = ContainerWidth() - (S(margin) * 2)
        .height = S(height)
        .caption = caption
        .Font.Name = FontName
        .Font.Size = FontSize + 3
        .Font.Bold = True
        .foreColor = HeaderColor
        .TextAlign = fmTextAlignLeft
    End With
End Sub

' ---------------------------------------------------------------------
' How much room this container actually gives you.
'
' Use these instead of hardcoded numbers in a page's ArrangeUI, so a
' change to the shell's FORM_WIDTH flows through to the tabs rather than
' leaving a ListBox stranded at its old size:
'
'   ui.FullWidth "lstViewClaims", ui.AvailableWidth * 0.45, 220
'
' ---------------------------------------------------------------------
Public Property Get AvailableWidth(Optional ByVal margin As Single = 14) As Single
    AvailableWidth = ContainerWidth() - (S(margin) * 2)
End Property

Public Property Get AvailableHeight(Optional ByVal margin As Single = 14) As Single
    AvailableHeight = ContainerHeight() - (S(margin) * 2)
End Property

' Remaining vertical room below the current row - useful for a list or
' text box that should stretch to the bottom of the page.
Public Property Get RemainingHeight(Optional ByVal bottomMargin As Single = 14) As Single
    RemainingHeight = ContainerHeight() - m_top - S(bottomMargin)
    If RemainingHeight < 0 Then RemainingHeight = 0
End Property

' The usable inside area. A Page reports InsideWidth; a UserForm does
' too, but some containers don't - fall back to Width so a missing
' property degrades to "slightly wrong" rather than a crash.
Private Function ContainerWidth() As Single
    On Error Resume Next
    ContainerWidth = m_container.InsideWidth
    If ContainerWidth = 0 Then ContainerWidth = m_container.width
    On Error GoTo 0
End Function

Private Function ContainerHeight() As Single
    On Error Resume Next
    ContainerHeight = m_container.InsideHeight
    If ContainerHeight = 0 Then ContainerHeight = m_container.height
    On Error GoTo 0
End Function

' =====================================================================
' Internals
' =====================================================================

' Positions the caption label to the left of the field, if there is one.
Private Sub PlaceCaption(ByVal captionName As String, ByVal caption As String, _
                          ByVal rowH As Single)
    Dim lbl As MSForms.Label
    If captionName = "" Then Exit Sub

    Set lbl = Resolve(captionName, True)
    If lbl Is Nothing Then Exit Sub

    With lbl
        .Left = m_left
        .Top = m_top + S(2)       ' aligns the text baseline with the field
        .width = LabelWidth - S(6)
        .height = RowHeight
        If caption <> "" Then .caption = caption
        .Font.Name = FontName
        .Font.Size = FontSize
        .Font.Bold = False
        .TextAlign = fmTextAlignRight
    End With
End Sub

' Applies type-appropriate styling. TypeOf checks rather than a single
' blanket style, because setting .BorderStyle on a CommandButton or
' .MultiLine on a ComboBox raises an error.
Private Sub StyleField(ctl As MSForms.control, ByVal isMultiline As Boolean, _
                        ByVal readOnlyLook As Boolean)
    On Error Resume Next

    ctl.Font.Name = FontName
    ctl.Font.Size = FontSize

    If TypeOf ctl Is MSForms.TextBox Then
        With ctl
            .BorderStyle = fmBorderStyleSingle
            .BackColor = IIf(readOnlyLook, ReadOnlyBackColor, EditableBackColor)
            .Locked = readOnlyLook
            If isMultiline Then
                .Multiline = True
                .WordWrap = True
                .ScrollBars = fmScrollBarsVertical
            End If
        End With

    ElseIf TypeOf ctl Is MSForms.ComboBox Then
        With ctl
            .BackColor = IIf(readOnlyLook, ReadOnlyBackColor, EditableBackColor)
            .Style = fmStyleDropDownList    ' stops typed junk getting in
        End With

    ElseIf TypeOf ctl Is MSForms.ListBox Then
        ctl.BorderStyle = fmBorderStyleSingle

    ElseIf TypeOf ctl Is MSForms.Label Then
        ctl.foreColor = &H80000012          ' window text
    End If

    On Error GoTo 0
End Sub

' Looks a control up by name. If optional is False and it's missing,
' the error names both the control AND the tab - far more useful than
' VBA's generic "Could not find the specified object" on a form with
' 70-odd controls.
Private Function Resolve(ByVal ctrlName As String, ByVal isOptional As Boolean) As Object
    
    On Error GoTo NotFound
    Set Resolve = m_container.Controls(ctrlName)
    Exit Function

NotFound:
    If isOptional Then
        Set Resolve = Nothing
        Exit Function
    End If
    Err.Raise vbObjectError + 21, "clsFormStyler", _
        "Control '" & ctrlName & "' was not found on the '" & m_pageTitle & "' tab." & vbCrLf & vbCrLf & _
        "Check the name in the form designer - it must match exactly."
End Function

Private Sub Class_Terminate()
    Set m_container = Nothing
End Sub


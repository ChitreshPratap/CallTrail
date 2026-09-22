Attribute VB_Name = "ListBoxUtil"
Option Explicit

Public Sub CreateListBoxHeaders(lst As MSForms.ListBox, _
                                headerNames As Variant, _
                                Optional headerHeight As Double = 15)
    
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim currentLeft As Single
    Dim colWidths() As String
    Dim singleWidth As Single
    Dim ctrl As control
    
    Dim lastWidth As Double
    
'    Dim uform As UserFormMain
'    Set uform = UserFormMain

    Dim parContainer As Variant
    Set parContainer = lst.Parent

    ' 1. Delete any existing dynamic headers (prevents duplicates if code runs twice)
    For Each ctrl In parContainer.Controls
        If Left(ctrl.Name, 10) = "dynHeader_" Then parContainer.Controls.Remove ctrl.Name
    Next ctrl

    ' 2. Start positioning at the left edge of the ListBox
    currentLeft = lst.Left

    ' 3. Parse the ColumnWidths property (e.g., "50;100;75")
    If lst.ColumnWidths <> "" Then
        colWidths = Split(lst.ColumnWidths, ";")
    End If

    ' 4. Create a label for each header
    For i = LBound(headerNames) To UBound(headerNames)
        
        ' Determine the width for this specific column
        If lst.ColumnWidths <> "" And i <= UBound(colWidths) Then
            singleWidth = Val(colWidths(i)) ' Val ignores the " pt" text if present
        Else
            ' Fallback: evenly divide the ListBox width if ColumnWidths aren't set
            singleWidth = lst.width / (UBound(headerNames) - LBound(headerNames) + 1)
        End If
        
        If i = UBound(colWidths) Then
            lastWidth = lst.width - currentLeft
        End If

        ' Add the label control to the UserForm dynamically
        Set lbl = parContainer.Controls.Add("Forms.Label.1", "dynHeader_" & i, True)
                
        ' Format and position the label
        With lbl
            
            .caption = " " & headerNames(i)
            .Left = currentLeft
            .Top = lst.Top - 15          ' Place it 15 points above the ListBox
            .width = singleWidth
            .height = headerHeight
            .BackColor = &H8000000F      ' Standard grey button-face color
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleSingle
            .Font.Bold = True
            .Font.Size = 9
        End With

        ' Move the starting position for the next label
        currentLeft = currentLeft + singleWidth
    Next i
End Sub








Public Sub generateListBoxHeaders(lst As MSForms.ListBox, _
                                headerNames As Variant, _
                                Optional HeaderColor As Long = &H4A4643, _
                                Optional boundedHeaderColor As Long = -1, _
                                Optional foreColor As Long = &HFFFFFF, _
                                Optional headerHeight As Double = 15, _
                                Optional headerFontSize As Double = 9)
                                
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim currentLeft As Single
    Dim colWidths() As String
    Dim singleWidth As Single
    Dim totalProvidedWidth As Single
    Dim ctrl As control
    Dim headerCount As Integer
    Dim widthCount As Integer
    
    Dim parContainer As Variant
    Set parContainer = lst.Parent
    
    ' ==========================================
    ' 1. PRE-EXECUTION VALIDATION CHECKS
    ' ==========================================
    
    ' Error 1: The ListBox object doesn't exist
    If lst Is Nothing Then
        Err.Raise vbObjectError + 1, "CreateListBoxHeaders", "The ListBox object provided is Nothing."
    End If
    
    ' Error 2: The headers were not passed as an Array
    If Not IsArray(headerNames) Then
        Err.Raise vbObjectError + 2, "CreateListBoxHeaders", "The headerNames argument must be an Array."
    End If
    
    ' Error 3: Not enough space to draw headers
    If lst.Top < 15 Then
        Err.Raise vbObjectError + 3, "CreateListBoxHeaders", "Not enough space above the ListBox. Move your ListBox down so its .Top property is at least 15."
    End If
    
    ' Calculate Header Count
    headerCount = UBound(headerNames) - LBound(headerNames) + 1
    
    ' Error 4: ColumnCount property doesn't match Header Array count
    If lst.ColumnCount <> headerCount Then
        Err.Raise vbObjectError + 4, "CreateListBoxHeaders", "ColumnCount Mismatch: ListBox.ColumnCount is " & lst.ColumnCount & " but you passed " & headerCount & " headers."
    End If
    
    ' Parse ColumnWidths and calculate Width Count
    If lst.ColumnWidths = "" Then
        widthCount = 0
    Else
        colWidths = Split(lst.ColumnWidths, ";")
        widthCount = UBound(colWidths) - LBound(colWidths) + 1
    End If
    
    ' Error 5: ColumnWidths property doesn't match Header Array count
    If widthCount <> headerCount Then
        Err.Raise vbObjectError + 5, "CreateListBoxHeaders", "Width Mismatch: You provided " & headerCount & " headers, but defined " & widthCount & " column widths. They must be exactly equal."
    End If

    ' Calculate Total Provided Width
    totalProvidedWidth = 0
    For i = LBound(colWidths) To UBound(colWidths)
        totalProvidedWidth = totalProvidedWidth + Val(colWidths(i))
    Next i
    
    ' Error 6: Provided widths exceed ListBox width
    If totalProvidedWidth >= lst.width Then
        Err.Raise vbObjectError + 6, "CreateListBoxHeaders", "Total Width Error: The sum of your column widths (" & totalProvidedWidth & ") is greater than or equal to the ListBox width (" & lst.width & "). Reduce your column widths."
    End If

    ' ==========================================
    ' 2. DRAWING THE HEADERS
    ' ==========================================

    ' Delete any existing dynamic headers (prevents duplicates)
    For Each ctrl In parContainer.Controls
        If ctrl.Name Like lst.Name & "_dynHeader_*" Then parContainer.Controls.Remove ctrl.Name
    Next ctrl

    ' Start positioning at the left edge of the ListBox
    currentLeft = lst.Left
    
    ' Create a label for each header
    For i = LBound(headerNames) To UBound(headerNames)
        
        ' Determine the width for this specific column
        If i = UBound(headerNames) Then
            ' It is the LAST column: stretch it to fill the exact remaining width of the ListBox
            singleWidth = lst.width - (currentLeft - lst.Left)
        Else
            ' It is a normal column: use the exact provided width
            singleWidth = Val(colWidths(i))
            
            ' Error 7: Column width evaluates to 0 or negative
            If singleWidth <= 0 Then
                Err.Raise vbObjectError + 7, "CreateListBoxHeaders", "Invalid Width: The width for column " & i + 1 & " is 0 or negative. Check your ColumnWidths string."
            End If
        End If

        ' Add the label control to the UserForm dynamically
        Set lbl = parContainer.Controls.Add("Forms.Label.1", lst.Name & "_dynHeader_" & i, True)
        
        ' Format and position the label
        With lbl
            .caption = " " & headerNames(i)
            .Left = currentLeft
            .Top = lst.Top - 15
            .width = singleWidth
            .height = headerHeight
            
            If i Mod 2 = 0 Then
                '.BackColor = RGB(1, 97, 148)
                .BackColor = HeaderColor
                
            Else
                If boundedHeaderColor = -1 Then
                    .BackColor = HeaderColor
                Else
                    .BackColor = boundedHeaderColor
                End If
            End If
            
            .SpecialEffect = fmSpecialEffectFlat
            .Font.Bold = True
            .Font.Size = headerFontSize
            .foreColor = foreColor
            
        End With

        ' Move the starting position for the next label
        currentLeft = currentLeft + singleWidth
    Next i
        
    lst.SpecialEffect = fmSpecialEffectFlat
        
    Exit Sub

    ' ==========================================
    ' 3. ERROR HANDLER
    ' ==========================================
ErrorHandler:
    MsgBox "Failed to create ListBox headers." & vbCrLf & vbCrLf & _
           "Error " & (Err.Number - vbObjectError) & ": " & Err.Description, _
           vbCritical, "CreateListBoxHeaders Error"
    Err.Clear
End Sub


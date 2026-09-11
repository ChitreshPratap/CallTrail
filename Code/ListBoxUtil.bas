Attribute VB_Name = "ListBoxUtil"
Option Explicit

Public Sub CreateListBoxHeaders(lst As MSForms.ListBox, headerNames As Variant)
    
    Dim i As Integer
    Dim lbl As MSForms.Label
    Dim CurrentLeft As Single
    Dim colWidths() As String
    Dim singleWidth As Single
    Dim ctrl As control
    
'    Dim uform As UserFormMain
'    Set uform = UserFormMain
        
    Dim parContainer As Variant
    Set parContainer = lst.Parent
        
    ' 1. Delete any existing dynamic headers (prevents duplicates if code runs twice)
    For Each ctrl In parContainer.Controls
    
        If Left(ctrl.Name, 10) = "dynHeader_" Then parContainer.Controls.Remove ctrl.Name
        
    Next ctrl

    ' 2. Start positioning at the left edge of the ListBox
    CurrentLeft = lst.Left
    
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

        ' Add the label control to the UserForm dynamically
        Set lbl = parContainer.Controls.Add("Forms.Label.1", "dynHeader_" & i, True)
        
        ' Format and position the label
        With lbl
            
            .caption = " " & headerNames(i)
            .Left = CurrentLeft
            .Top = lst.Top - 15          ' Place it 15 points above the ListBox
            .width = singleWidth
            .height = 15
            .BackColor = &H8000000F      ' Standard grey button-face color
            .SpecialEffect = fmSpecialEffectFlat
            .BorderStyle = fmBorderStyleSingle
            .Font.Bold = True
            .Font.Size = 9
        End With

        ' Move the starting position for the next label
        CurrentLeft = CurrentLeft + singleWidth
    Next i
End Sub







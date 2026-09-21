Attribute VB_Name = "modPageHelpers"
Option Explicit

' =====================================================================
' modPageHelpers
' Small shared utilities for the page controllers.

' The important one is BindControl: every controller looks its controls
' up by name from its own Page. If a control is missing or misnamed, the
' error names the control AND the page, instead of VBA's generic
' "Could not find the specified object" - which, on a form with 150
' controls across 5 tabs, tells you nothing useful.
' =====================================================================

Public Function BindControl(pg As Msforms.Page, ByVal ctrlName As String, _
                             ByVal pageTitle As String) As Msforms.control
    On Error GoTo NotFound
    Set BindControl = pg.Controls(ctrlName)
    Exit Function

NotFound:
    Err.Raise vbObjectError + 10, "BindControl", _
        "Control '" & ctrlName & "' was not found on the '" & pageTitle & "' tab." & vbCrLf & vbCrLf & _
        "Check the control name in the form designer - it must match exactly."
End Function

' Same, but returns Nothing instead of raising when a control is genuinely
' optional (e.g. an admin-only note label).
Public Function BindOptional(pg As Msforms.Page, ByVal ctrlName As String) As Msforms.control
    On Error Resume Next
    Set BindOptional = pg.Controls(ctrlName)
    On Error GoTo 0
End Function

' Standard list-box setup so every grid on every tab looks the same.
Public Sub ConfigureClaimList(lst As Msforms.ListBox)


    
    
    ' 3. Define your headers in an Array
    With lst
        .ColumnCount = 6
        .ColumnHeads = False
        .ColumnWidths = "70;80;140;60;50;80"
        .MultiSelect = fmMultiSelectSingle
    End With
    
End Sub

Public Sub ConfigureHistoryList(lst As Msforms.ListBox)

    With lst
        .ColumnCount = 4
        .ColumnHeads = False
        .ColumnWidths = "90;120;60;260"
    End With
    
End Sub

Public Function FormatDateOrBlank(ByVal v As Variant) As String
    
    If IsEmpty(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    FormatDateOrBlank = Format$(CDate(v), "dd-mmm-yyyy hh:nn")
    
End Function

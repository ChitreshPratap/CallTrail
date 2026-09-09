Attribute VB_Name = "modMain"
Option Explicit

' =====================================================================
' modMain  (CallTrail app)
' The only module that should be wired to buttons/ribbon/shortcut keys.
' Each Sub here just shows a form; the forms call modDataAccess.
' Add new features by adding one Sub here + one form, without touching
' modDataAccess or modUtils.
' =====================================================================

Public Sub ShowAddClaimForm()
    frmAddClaim.Show
End Sub

Public Sub ShowUpdateClaimForm()
    frmUpdateClaim.Show
End Sub

Public Sub ShowAdminLocationForm()
    Dim wb As Workbook
    Dim isAdmin As Boolean

    Set wb = OpenCentralDB()
    isAdmin = IsCurrentUserAdmin(wb)
    CloseCentralDB wb, False

    If Not isAdmin Then
        MsgBox "This feature is restricted to Admin users.", vbExclamation, "Access Denied"
        Exit Sub
    End If
    frmAdminLocation.Show
End Sub

Public Sub ShowClaimHistoryViewer()
    frmClaimHistory.Show
End Sub

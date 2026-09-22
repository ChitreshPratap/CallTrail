Attribute VB_Name = "modRibbon"
Option Explicit

'-------------------------------------------------------------------
' Single dispatcher for every button
'-------------------------------------------------------------------
Public Sub OnRibbonAction(control As IRibbonControl)

    Select Case control.ID
        
        Case "btnExportAll":
            modMain.DownloadDatabase
        
        Case "btnOpenApp":
            UserFormMain.Show

        ' --- Claim operations ---
        Case "btnAddClaimManual1", "btnAddClaimManual2":
            
            'modMain.ShowAddClaimForm
            modMain.ShowAddClaimForm_tab 1
            'modMain.ShowAddClaimForm_tab 2
        
        Case "btnAddClaimBulk":
            modMain.BulkAddClaims
            
        Case "btnGenerateTempateAddClaimBulk":
            modMain.BulkSheetSetup
        
'        ' --- Search ---
'        Case "btnRetrieveAll":                      RetrieveClaims
'        Case "btnFindClaim":                        FindClaim
'        Case "btnSelectActive":                     SelectHighlighted
'        Case "btnClearFilters":                     ResetView
'
'        ' --- Reporting ---
'        Case "btnExportSummary", "btnExportExcel":  ExportSummary "XLSX"
'        Case "btnExportPdf":                        ExportSummary "PDF"
'        Case "btnExportMail":                       ExportSummary "MAIL"
'        Case "btnExportClip":                       ExportSummary "CLIP"
'        Case "btnDashboard":                        ShowDashboard
'
'        ' --- Session ---
'        Case "btnSyncNow":                          SyncWithCentral
'        Case "btnAdminPanel":                       ShowAdminPanel
'        Case "btnHelp":                             ShowGuide

        Case Else
            MsgBox "No handler mapped for: " & control.ID, vbExclamation, "CallTrail"
    End Select

    Exit Sub
Fail:
    MsgBox "Action failed (" & control.ID & "): " & Err.Description, vbCritical, "CallTrail"
End Sub


Private Sub ShowClaimForm():             MsgBox "ShowClaimForm":            End Sub
Private Sub ImportClaimsFromFile():      MsgBox "ImportClaimsFromFile":     End Sub
Private Sub ImportClaimsFromClipboard(): MsgBox "ImportClaimsFromClipboard": End Sub
Private Sub SaveIntakeTemplate():        MsgBox "SaveIntakeTemplate":       End Sub
Private Sub EditSelectedClaim():         MsgBox "EditSelectedClaim":        End Sub
Private Sub LogCallAttempt():            MsgBox "LogCallAttempt":           End Sub
Private Sub ShowClaimHistory():          MsgBox "ShowClaimHistory":         End Sub
Private Sub DeleteSelectedClaim():       MsgBox "DeleteSelectedClaim":      End Sub
Private Sub PurgeFilteredClaims():       MsgBox "PurgeFilteredClaims":      End Sub
Private Sub ShowSettings():              MsgBox "ShowSettings":             End Sub
Private Sub RetrieveClaims():            MsgBox "RetrieveClaims":           End Sub
Private Sub FindClaim():                 MsgBox "FindClaim":                End Sub
Private Sub SelectHighlighted():         MsgBox "SelectHighlighted":        End Sub
Private Sub ResetView():                 MsgBox "ResetView":                End Sub
Private Sub ShowDashboard():             MsgBox "ShowDashboard":            End Sub
Private Sub SyncWithCentral():           gLastSync = Now:                   End Sub
Private Sub ShowAdminPanel():            MsgBox "ShowAdminPanel":           End Sub
Private Sub ShowGuide():                 MsgBox "ShowGuide":                End Sub
Private Sub ApplyStatusFilter(ByVal S As String): MsgBox "Filter status: " & S: End Sub
Private Sub ApplyView(ByVal v As String):         MsgBox "Apply view: " & v:    End Sub



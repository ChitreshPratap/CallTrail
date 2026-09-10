VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsContAddClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim var_frm As UserFormMain

Public Property Set setForm(uForm As UserFormMain)

    Set var_frm = uForm

End Property


Public Sub cmdSave_Click(var_frm As UserFormMain)
    
    Dim tClaimId As String
    Dim tClaimSite As String
    Dim tProviderName As String
    Dim tDtCreationDate As String
    Dim tClaimQuery As String
    
    Dim frm As UserFormMain
    Set frm = var_frm
        
    tClaimId = Trim(frm.txtClaimID.Value)
    tClaimSite = Trim(frm.txtClaimSite.Value)
    tProviderName = Trim(frm.txtProviderName.Value)
    tClaimQuery = Trim(frm.txtClaimQuery.Value)
    tDtCreationDate = Trim(frm.dtCreationDate.Value)
        
    If tClaimId = "" Then
        MsgBox "'Claim ID' is required.", vbExclamation
        Exit Sub
    End If
    If tClaimQuery = "" Then
        MsgBox "'Claim Query' is required.", vbExclamation
        Exit Sub
    End If

    If tClaimSite = "" Then
        MsgBox "'Claim Site' is required.", vbExclamation
        Exit Sub
    End If

    If tProviderName = "" Then
        MsgBox "'Provider Name' is required.", vbExclamation
        Exit Sub
    End If

    If tDtCreationDate = "" Then
        MsgBox "'Creation Date' is required.", vbExclamation
        Exit Sub
    End If
    
    If Not IsDate(tDtCreationDate) Then
        MsgBox "Creation Date is not a valid date.", vbExclamation
        Exit Sub
    End If

    If AddClaim(tClaimId, tClaimSite, tProviderName, tClaimQuery, CDate(tDtCreationDate)) Then
        
        MsgBox "Claim added.", vbInformation
        resetAddClaim var_frm
        'Unload Me
    End If
    
End Sub

Public Sub cmdCancel_Click(var_frm As UserFormMain)
    
    Unload var_frm
    
End Sub

Public Sub resetAddClaim(var_frm As UserFormMain)
        
    var_frm.txtClaimID.Value = ""
    var_frm.txtClaimSite.Value = ""
    var_frm.txtProviderName.Value = ""
    var_frm.txtClaimQuery.Value = ""
    var_frm.dtCreationDate.Value = ""
    
End Sub



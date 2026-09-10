VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdminLocation 
   Caption         =   "Admin Location"
   ClientHeight    =   4710
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8580.001
   OleObjectBlob   =   "frmAdminLocation.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdminLocation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdUpdate_Click()
    If ChangeUpdatedSite(Trim(txtClaimID.Value), Trim(txtNewSite.Value)) Then
        MsgBox "Site updated.", vbInformation
        Unload Me
    End If
End Sub

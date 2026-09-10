VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsClaim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsClaim
' One claim record as an object. Mirrors the Claims sheet columns 1:1
' by name so it's obvious which property maps to which header - if you
' add a column to the sheet later, add one matching property here.
' =====================================================================

Private pClaimID As String
Private pClaimSite As String
Private pClaimProviderName As String
Private pClaimQuery As String
Private pClaimCreationDate As Date
Private pClaimStatus As String
Private pAttempt As Long
Private pClaimUpdatedSite As String
Private pClaimInsertionDate As Date
Private pClaimInsertedBy As String
Private pLastUpdatedDate As Variant   ' Variant: may be blank until first call
Private pLastUpdatedBy As String
Private pLastComment As String
Private pClaimClosedDate As Variant   ' Variant: blank until Closed

Public Property Get claimID() As String: claimID = pClaimID: End Property
Public Property Let claimID(ByVal v As String): pClaimID = v: End Property

Public Property Get claimSite() As String: claimSite = pClaimSite: End Property
Public Property Let claimSite(ByVal v As String): pClaimSite = v: End Property

Public Property Get ClaimProviderName() As String: ClaimProviderName = pClaimProviderName: End Property
Public Property Let ClaimProviderName(ByVal v As String): pClaimProviderName = v: End Property

Public Property Get claimQuery() As String: claimQuery = pClaimQuery: End Property
Public Property Let claimQuery(ByVal v As String): pClaimQuery = v: End Property

Public Property Get ClaimCreationDate() As Date: ClaimCreationDate = pClaimCreationDate: End Property
Public Property Let ClaimCreationDate(ByVal v As Date): pClaimCreationDate = v: End Property

Public Property Get ClaimStatus() As String: ClaimStatus = pClaimStatus: End Property
Public Property Let ClaimStatus(ByVal v As String): pClaimStatus = v: End Property

Public Property Get attempt() As Long: attempt = pAttempt: End Property
Public Property Let attempt(ByVal v As Long): pAttempt = v: End Property

Public Property Get ClaimUpdatedSite() As String: ClaimUpdatedSite = pClaimUpdatedSite: End Property
Public Property Let ClaimUpdatedSite(ByVal v As String): pClaimUpdatedSite = v: End Property

Public Property Get ClaimInsertionDate() As Date: ClaimInsertionDate = pClaimInsertionDate: End Property
Public Property Let ClaimInsertionDate(ByVal v As Date): pClaimInsertionDate = v: End Property

Public Property Get ClaimInsertedBy() As String: ClaimInsertedBy = pClaimInsertedBy: End Property
Public Property Let ClaimInsertedBy(ByVal v As String): pClaimInsertedBy = v: End Property

Public Property Get LastUpdatedDate() As Variant: LastUpdatedDate = pLastUpdatedDate: End Property
Public Property Let LastUpdatedDate(ByVal v As Variant): pLastUpdatedDate = v: End Property

Public Property Get LastUpdatedBy() As String: LastUpdatedBy = pLastUpdatedBy: End Property
Public Property Let LastUpdatedBy(ByVal v As String): pLastUpdatedBy = v: End Property

Public Property Get LastComment() As String: LastComment = pLastComment: End Property
Public Property Let LastComment(ByVal v As String): pLastComment = v: End Property

Public Property Get ClaimClosedDate() As Variant: ClaimClosedDate = pClaimClosedDate: End Property
Public Property Let ClaimClosedDate(ByVal v As Variant): pClaimClosedDate = v: End Property

Public Property Get IsClosed() As Boolean
    IsClosed = (LCase$(Trim$(pClaimStatus)) = "closed")
End Property

' Convenience: days open so far (or total days-to-close if closed).
' Computed live from dates already on the object - never stored,
' so it can't go stale.
Public Property Get DaysOpen() As Long
    If IsClosed And Not IsEmpty(pClaimClosedDate) Then
        DaysOpen = CLng(CDate(pClaimClosedDate)) - CLng(pClaimCreationDate)
    Else
        DaysOpen = CLng(Date) - CLng(pClaimCreationDate)
    End If
End Property


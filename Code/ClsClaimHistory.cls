VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsClaimHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private var_HistoryId As String
Private var_ClaimID As String
Private var_CallerName As String
Private var_CallDate As Date
Private var_Status As String
Private var_comment As String

Public Property Get historyId() As String: historyId = var_HistoryId: End Property
Public Property Let historyId(ByVal v As String): var_HistoryId = v: End Property

Public Property Get claimID() As String: claimID = var_ClaimID: End Property
Public Property Let claimID(ByVal v As String): var_ClaimID = v: End Property

Public Property Get callerName() As String: callerName = var_CallerName: End Property
Public Property Let callerName(ByVal v As String): var_CallerName = v: End Property

Public Property Get callDate() As String: callDate = var_CallDate: End Property
Public Property Let callDate(ByVal v As String): var_CallDate = v: End Property

Public Property Get callStatus() As String: callStatus = var_Status: End Property
Public Property Let callStatus(ByVal v As String): var_Status = v: End Property

Public Property Get comment() As String: comment = var_comment: End Property
Public Property Let comment(ByVal v As String): var_comment = v: End Property



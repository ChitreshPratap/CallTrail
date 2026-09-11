VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHistoryEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsHistoryEntry
' One row of the History sheet as an object - a single logged call.
' =====================================================================

Private pHistoryId As Long
Private pClaimID As String
Private pCallerName As String
Private pCallDateTime As Date
Private pCallerComment As String
Private pCallerStatus As String

Public Property Get historyId() As Long: historyId = pHistoryId: End Property
Public Property Let historyId(ByVal v As Long): pHistoryId = v: End Property

Public Property Get claimID() As String: claimID = pClaimID: End Property
Public Property Let claimID(ByVal v As String): pClaimID = v: End Property

Public Property Get callerName() As String: callerName = pCallerName: End Property
Public Property Let callerName(ByVal v As String): pCallerName = v: End Property

Public Property Get CallDateTime() As Date: CallDateTime = pCallDateTime: End Property
Public Property Let CallDateTime(ByVal v As Date): pCallDateTime = v: End Property

Public Property Get CallerComment() As String: CallerComment = pCallerComment: End Property
Public Property Let CallerComment(ByVal v As String): pCallerComment = v: End Property

Public Property Get CallerStatus() As String: CallerStatus = pCallerStatus: End Property
Public Property Let CallerStatus(ByVal v As String): pCallerStatus = v: End Property


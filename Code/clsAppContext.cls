VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAppContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' clsAppContext
' One shared object handed to every page controller. This is what makes
' the tabbed shell cheaper than five separate forms rather than more
' expensive.

' Without it, each tab would hit the shared workbook on its own: the
' View tab loads all claims, the Search tab loads them again, the Admin
' tab checks the role again. Every one of those is an open/read/close
' of a file on a network share that other callers are also using.
'
' With it, the claim list and the user's role are loaded ONCE and cached.
' Any page that writes calls InvalidateClaims, and the next page to need
' the list reloads it. Nothing reloads speculatively.
' =====================================================================

Private m_repo As ClsClaimRepository
Private m_claims As Collection
Private m_claimsLoaded As Boolean
Private m_isAdmin As Boolean
Private m_roleChecked As Boolean
Private m_statusSink As Object          ' the shell form, for status-bar messages


Private Sub Class_Initialize()
    Set m_repo = New ClsClaimRepository
    Set m_claims = New Collection
    m_claimsLoaded = False
    m_roleChecked = False
End Sub

Public Property Get repo() As ClsClaimRepository
    Set repo = m_repo
End Property

' ---------------------------------------------------------------------
' The cached claim list. Loads on first use, then serves from memory.
' Pass forceReload:=True after a write, or for an explicit Refresh.
' ---------------------------------------------------------------------
Public Function Claims(Optional ByVal forceReload As Boolean = False) As Collection
    
    If forceReload Or Not m_claimsLoaded Then
        SetStatus "Loading claims..."
        Set m_claims = m_repo.GetAllClaims()
        m_claimsLoaded = True
        SetStatus "Ready."
    End If
    Set Claims = m_claims

End Function

' Call this after ANY write that changes the Claims sheet. It doesn't
' reload immediately - it just marks the cache stale so the next tab
' that needs claims picks up fresh data.
Public Sub InvalidateClaims()
    m_claimsLoaded = False
End Sub

Public Property Get ClaimsAreLoaded() As Boolean
    ClaimsAreLoaded = m_claimsLoaded
End Property

' ---------------------------------------------------------------------
' Admin check, cached. Role doesn't change mid-session, and checking it
' means opening the shared workbook, so we do it once.
' ---------------------------------------------------------------------
Public Property Get isAdmin() As Boolean
    Dim wb As Workbook

    If Not m_roleChecked Then
        On Error Resume Next
        Set wb = OpenCentralDB()
        If Not wb Is Nothing Then
            m_isAdmin = IsCurrentUserAdmin(wb)
            CloseCentralDB wb, False
        End If
        On Error GoTo 0
        m_roleChecked = True
    End If
    isAdmin = m_isAdmin
End Property

Public Property Get userName() As String
    userName = GetWindowsUserName()
End Property

' ---------------------------------------------------------------------
' Status-bar plumbing. Pages call SetStatus without knowing or caring
' that a form is on the other end - keeps controllers decoupled from
' the shell.
' ---------------------------------------------------------------------
Public Sub SetStatusSink(frm As Object)
    Set m_statusSink = frm
End Sub

Public Sub SetStatus(ByVal msg As String)
    On Error Resume Next
    If Not m_statusSink Is Nothing Then
        m_statusSink.ShowStatus msg
    End If
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    Set m_repo = Nothing
    Set m_claims = Nothing
    Set m_statusSink = Nothing
End Sub


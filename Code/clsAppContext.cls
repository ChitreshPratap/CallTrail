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
'
' Without it, each tab would hit the shared workbook on its own: the
' View tab loads all claims, the Search tab loads them again, the Admin
' tab checks the role again. Every one of those is an open/read/close
' of a file on a network share that other callers are also using.
'
' With it, the claim list and the user's role are loaded ONCE and cached.
' Any page that writes calls InvalidateClaims, and the next page to need
' the list reloads it. Nothing reloads speculatively.
' =====================================================================

Private m_repo As clsClaimRepository
Private m_claims As Collection
Private m_claimsLoaded As Boolean
Private m_isAdmin As Boolean
Private m_roleChecked As Boolean
Private m_statusSink As Object      ' the shell form, for status-bar messages

Private Sub Class_Initialize()
    Set m_repo = New clsClaimRepository
    Set m_claims = New Collection
    m_claimsLoaded = False
    m_roleChecked = False
End Sub

Public Property Get Repo() As clsClaimRepository
    Set Repo = m_repo
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
' User identity, cached. Role and registration don't change mid-session,
' and checking means opening the shared workbook, so we do it once.
'
' An unregistered user is a normal state, not an error: they get Guest
' role, the home page explains why, and every other tab is locked.
' ---------------------------------------------------------------------
Private m_registered As Boolean
Private m_role As String
Private m_defaultLocation As String

Private Sub EnsureUserLoaded()
    Dim wb As Workbook
    Dim role As String, loc As String

    If m_roleChecked Then Exit Sub
    m_roleChecked = True
    m_role = "Guest"

    On Error Resume Next
    Set wb = OpenCentralDB()
    If Not wb Is Nothing Then
        m_registered = GetUserRecord(wb, role, loc)
        CloseCentralDB wb, False
        If m_registered Then
            m_role = role
            m_defaultLocation = loc
            m_isAdmin = (LCase$(role) = "admin")
        End If
    End If
    On Error GoTo 0
End Sub

Public Property Get IsAdmin() As Boolean
    EnsureUserLoaded
    IsAdmin = m_isAdmin
End Property

' False means the Windows user isn't in the Users sheet at all.
Public Property Get IsRegistered() As Boolean
    EnsureUserLoaded
    IsRegistered = m_registered
End Property

Public Property Get Role() As String
    EnsureUserLoaded
    Role = m_role
End Property

' The site to pre-fill on the Add Claim form. Empty if the user is
' unregistered or the Users sheet has no DefaultLocation column.
' Clears the cached lookup so a just-registered user can be picked up
' without restarting Excel.
Public Sub ResetUserCache()
    m_roleChecked = False
    m_registered = False
    m_isAdmin = False
    m_role = "Guest"
    m_defaultLocation = ""
End Sub

Public Property Get DefaultLocation() As String
    EnsureUserLoaded
    DefaultLocation = m_defaultLocation
End Property

Public Property Get UserName() As String
    UserName = GetWindowsUserName()
End Property

' ---------------------------------------------------------------------
' UI scale, measured once by the shell and reused by every page.
'
' Calibration creates and measures a temporary label, so doing it in all
' six places (shell + five tabs) would repeat the same measurement six
' times for an identical answer.
' ---------------------------------------------------------------------
Private m_uiScale As Single

Public Property Get UIScale() As Single
    If m_uiScale <= 0 Then m_uiScale = 1    ' not calibrated yet
    UIScale = m_uiScale
End Property

Public Property Let UIScale(ByVal v As Single)
    If v > 0 Then m_uiScale = v
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

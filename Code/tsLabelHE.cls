VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "tsLabelHE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' It implement the hover effect
' It changes the Font Size Change Only Of the provided label on Hover On/Hover Off

Dim var_Name As String
Dim var_mukhLabel As MSForms.Label
Dim var_originalFont As Long
Dim var_originalBackColor As Long
Dim var_originalForeColor As Long
Dim var_incDecFontSize As Long
Dim var_onHoverBackColor As Long
Dim var_onHoverForeColor As Long

Public WithEvents dynaLabelParent As MSForms.UserForm
Attribute dynaLabelParent.VB_VarHelpID = -1
Public WithEvents dynaLabel As MSForms.Label
Attribute dynaLabel.VB_VarHelpID = -1


Property Get classModuleName() As String
    
    Dim tName As String
    tName = var_Name

End Property

Private Sub Class_Initialize()

    var_Name = "tsLabelHE"
    var_incDecFontSize = 1

End Sub

'Sub setMukhLabel(mukhLabel As control)
Sub setMukhLabel(mukhLabel As MSForms.Label)
    
    Dim clor As Variant
    Dim fClor As Variant
    Set var_mukhLabel = mukhLabel
    Set dynaLabel = var_mukhLabel
    Set dynaLabelParent = var_mukhLabel.Parent
    var_originalFont = mukhLabel.Font.Size
    var_originalBackColor = var_mukhLabel.BackColor
    var_originalForeColor = var_mukhLabel.foreColor
    
End Sub

Function getMukhLabel() As MSForms.Label
    
    Dim tLabel As MSForms.Label
    Set tLabel = var_mukhLabel
    If tLabel Is Nothing Then
        Err.Raise vbObjectError + 1001, Me.classModuleName & "." & "getMukhLabel", "LabelNotSetError : Label is not set, Please set the label using function 'setMukhLabel' "
    End If
    Set getMukhLabel = tLabel

End Function

Function setOnHoverBackColor(colorValue As Long)
    var_onHoverBackColor = colorValue
End Function

Function setOnHoverForeColor(colorValue As Long)
    var_onHoverForeColor = colorValue
End Function

Function getOnHoverBackColor() As Long
    Dim tColor As Long
    tColor = var_onHoverBackColor
    getOnHoverBackColor = tColor
End Function

Function getOnHoverForeColor() As Long
    Dim tColor As Long
    tColor = var_onHoverForeColor
    getOnHoverForeColor = tColor
End Function

Function setIncDecInFont(incDecFontSize As Long)
    
    'It set the increament and decrement in font size on hover and hover off
    'incDecFontSize can be positive or negative
    'If positive then font size increased on hover
    'If negative then font size decreased on hover
    
    var_incDecFontSize = incDecFontSize

End Function

Function getIncDecInFont() As Long
    
    Dim tIncDecFontSize As Long
    tIncDecFontSize = var_incDecFontSize
    getIncDecInFont = tIncDecFontSize
    
End Function


Private Sub styleHoverOn()
    
    Dim tMukhLabel As MSForms.Label
    Set tMukhLabel = Me.getMukhLabel()
    Dim diffFontSize As Long
    diffFontSize = Me.getIncDecInFont()
    
    With tMukhLabel
        .Font.Size = var_originalFont + diffFontSize
    End With
    
    Dim onHoverBackColor As Long
    onHoverBackColor = Me.getOnHoverBackColor()
    If onHoverBackColor <> 0 Then
        With tMukhLabel
            .BackColor = onHoverBackColor
        End With
    End If
    
    Dim onHoverForeColor As Long
    onHoverForeColor = Me.getOnHoverForeColor()
    If onHoverForeColor <> 0 Then
        With tMukhLabel
            .foreColor = onHoverForeColor
        End With
    End If
    
End Sub

Private Sub styleHoverOff()
    
    Dim tMukhLabel As MSForms.Label
    Set tMukhLabel = Me.getMukhLabel()
    With tMukhLabel
        .Font.Size = var_originalFont
        .BackColor = var_originalBackColor
        .foreColor = var_originalForeColor
    End With
End Sub

Private Sub dynaLabel_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    styleHoverOn

End Sub

Private Sub dynaLabelParent_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    styleHoverOff

End Sub


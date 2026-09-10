Attribute VB_Name = "VarnahUtil"

Public Function isLightColor(ByVal color As Long) As Boolean
    Dim isDark As Boolean
    isDark = VarnahUtil.isDarkColor(color)
    isLightColor = Not (isDark)
End Function

Public Function isDarkColor(ByVal color As Long) As Boolean
    
    Dim r As Long, g As Long, b As Long
    r = (color And &HFF)
    g = ((color \ &H100) And &HFF)
    b = ((color \ &H10000) And &HFF)

    Dim brightness As Long
    brightness = 0.2126 * r + 0.7152 * g + 0.0722 * b
    
    If brightness < 128 Then
        isDark = True
    Else
        isDark = False
    End If
    
    
End Function

Public Function getFadeColor(ByVal color As Long, ByVal fadePercent As Double) As Long
    
    Dim r As Long, g As Long, b As Long
    Dim targetR As Long, targetG As Long, targetB As Long
    Dim fadeFactor As Double
    
    ' Extract RGB components from the input color
    r = (color And &HFF)
    g = ((color \ &H100) And &HFF)
    b = ((color \ &H10000) And &HFF)
    
    ' Ensure fadePercent is within -1 to 1 (-100% to +100%)
    fadePercent = WorksheetFunction.Min(WorksheetFunction.Max(fadePercent, -1), 1)
    
    If fadePercent >= 0 Then
        ' Fade toward white (255, 255, 255)
        targetR = 255
        targetG = 255
        targetB = 255
        fadeFactor = fadePercent
    Else
        ' Fade toward black (0, 0, 0)
        targetR = 0
        targetG = 0
        targetB = 0
        fadeFactor = -fadePercent
    End If
    
    ' Calculate the new RGB values by blending
    r = r + (targetR - r) * fadeFactor
    g = g + (targetG - g) * fadeFactor
    b = b + (targetB - b) * fadeFactor
    
    ' Combine the adjusted RGB components into a single color
    getFadeColor = RGB(r, g, b)
End Function


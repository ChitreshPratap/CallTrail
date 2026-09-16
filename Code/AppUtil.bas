Attribute VB_Name = "AppUtil"

Function getAppName() As String
    Dim appName As String
    appName = "Call Trail"
    getAppName = appName
End Function

Function getThemeColor() As Long
    Dim baseColor As Long
    'baseColor = RGB(97, 68, 229)
    baseColor = RGB(68, 188, 229)
    'baseColor = RGB(168, 229, 68)
    baseColor = RGB(0, 44, 58)
    getThemeColor = baseColor
End Function

Function getMenuItemsColor() As Long
End Function

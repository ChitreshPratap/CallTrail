Attribute VB_Name = "ColorUtil"


Function AdjustColor(ByVal Col As Long, ByVal Factor As Double) As Long
    ' Factor > 1 ? lighter
    ' Factor < 1 ? darker
    ' Example: 1.2 = 20% lighter, 0.8 = 20% darker
    
    Dim r As Long, g As Long, b As Long
    
    ' Extract RGB
    r = Col Mod 256
    g = (Col \ 256) Mod 256
    b = (Col \ 65536) Mod 256
    
    ' Scale with factor and clamp to [0,255]
    r = Application.Min(255, Application.Max(0, r * Factor))
    g = Application.Min(255, Application.Max(0, g * Factor))
    b = Application.Min(255, Application.Max(0, b * Factor))
    
    ' Return adjusted color
    AdjustColor = RGB(r, g, b)
End Function

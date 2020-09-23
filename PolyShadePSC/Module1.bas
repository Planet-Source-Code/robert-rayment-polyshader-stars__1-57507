Attribute VB_Name = "Module1"
' Module.bas

Option Explicit

Public Const pi# = 3.1415926535898
Public Const r2d# = 180 / pi# ' rad to deg
Public Const d2r# = pi# / 180 ' deg to rad

Public Function zATan2(ByVal yd As Single, ByVal xd As Single) As Single
' Public Const pi# = 3.1415926535898
' Input: (yd, xd) = (y - ycen, x - xcen)
   If yd = 0 Then
      If xd >= 0 Then zATan2 = 0 Else zATan2 = pi#
   Else
      zATan2 = Atn(xd / yd) + pi# / 2
      If yd > 0 Then zATan2 = zATan2 + pi#
   End If
End Function

Public Sub LngToRGB(LCul As Long, bred As Byte, bgreen As Byte, bblue As Byte)
'Convert Long LCul to RGB components
bred = (LCul And &HFF&)
bgreen = (LCul And &HFF00&) / &H100&
bblue = (LCul And &HFF0000) / &H10000
End Sub

Public Sub RGB2HSL(ByVal zR As Single, ByVal zG As Single, ByVal zB As Single, _
   zH As Single, zS As Single, zL As Single)
' In: zRGB  Out: zHSL
Dim ColMax As Long, ColMin As Long
Dim MmM As Long, MpM As Long
Dim zMul As Single
Dim zRD As Single, zGD As Single, zBD As Single
   ColMax = zR
   If zG > zR Then ColMax = zG
   If zB > ColMax Then ColMax = zB
   ColMin = zR
   If zG < zR Then ColMin = zG
   If zB < ColMin Then ColMin = zB
   MmM = ColMax - ColMin
   MpM = ColMax + ColMin
   zL = MpM / 2
   
   If ColMax = ColMin Then
      zS = 0
      zH = 170
   Else
      If zL <= 127.5 Then
         zS = MmM * 255 / MpM
      Else
         zS = MmM * 255 / (510 - MpM)
      End If
      zMul = 255 / (MmM * 6)
      zRD = (ColMax - zR) * zMul
      zGD = (ColMax - zG) * zMul
      zBD = (ColMax - zB) * zMul
      Select Case ColMax
      Case zR: zH = zBD - zGD
      Case zG: zH = 85 + zRD - zBD
      Case zB: zH = 170 + zGD - zRD
      End Select
      If zH < 0 Then zH = zH + 255
   End If
End Sub

Public Sub HSL2RGB(ByVal zH As Single, ByVal zS As Single, ByVal zL As Single, _
   zR As Single, zG As Single, zB As Single)
' In: zHSL   Out: zRGB
Dim zFactA As Single, zFactB As Single

   If zH > 255 Then zH = 255
   If zS > 255 Then zS = 255
   If zL > 255 Then zS = 255
   If zH < 0 Then zH = 0
   If zS < 0 Then zS = 0
   If zL < 0 Then zS = 0
   If zS = 0 Then
      zR = zL
      zG = zR
      zB = zR
   Else
      If zL <= 127.5 Then
         zFactA = zL * (255 + zS) / 255
      Else
         zFactA = zL + zS - zL * zS / 255
      End If
      zFactB = zL + zL - zFactA
            
      zR = (Hue2RGB(zFactA, zFactB, zH + 85)) And 255
      zG = (Hue2RGB(zFactA, zFactB, zH)) And 255
      zB = (Hue2RGB(zFactA, zFactB, zH - 85)) And 255
   End If
End Sub

Function Hue2RGB(zFA As Single, zFB As Single, ByVal zH As Single) As Long
' Called by HSL2RGB
   Select Case zH
   Case Is < 0: zH = zH + 255
   Case Is > 255: zH = zH - 255
   End Select
       
   Select Case zH
   Case Is < 42.5
      Hue2RGB = zFB + 6 * (zFA - zFB) * zH / 255
   Case Is < 127.5
      Hue2RGB = zFA
   Case Is < 170
      Hue2RGB = zFB + 6 * (zFA - zFB) * (170 - zH) / 255
   Case Else
      Hue2RGB = zFB
   End Select
End Function

Public Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   
   p = InStr(1, FSpec$, ".")
   
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub


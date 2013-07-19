Imports System.Drawing

Public Class ColorHandler
    ' Handle conversions between RGB and HSV _
    ' (and Color types, as well).

  Public Structure RGB
    ' All values are between 0 and 255.
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer

    Public Sub New( _
     ByVal R As Integer, ByVal G As Integer, ByVal B As Integer)
      Red = R
      Green = G
      Blue = B
    End Sub

    Public Overrides Function ToString() As String
      Return String.Format("({0}, {1}, {2})", _
       Red, Green, Blue)
    End Function
  End Structure

  Public Structure HSV
    ' All values are between 0 and 255.
    Dim Hue As Integer
    Dim Saturation As Integer
    Dim Value As Integer

    Public Sub New( _
     ByVal H As Integer, ByVal S As Integer, ByVal V As Integer)
      Hue = H
      Saturation = S
      Value = V
    End Sub

    Public Overrides Function ToString() As String
      Return String.Format("({0}, {1}, {2})", _
       Hue, Saturation, Value)
    End Function
  End Structure

  Public Shared Function HSVtoRGB(ByVal H As Integer, ByVal S As Integer, ByVal V As Integer) As RGB
    ' H, S, and V must all be between 0 and 255.
    Return HSVtoRGB(New HSV(H, S, V))
  End Function

  Public Shared Function HSVtoColor(ByVal HSV As HSV) As Color
    Dim RGB As RGB = HSVtoRGB(HSV)
    Return Color.FromArgb(RGB.Red, RGB.Green, RGB.Blue)
  End Function

  Public Shared Function HSVtoColor(ByVal H As Integer, ByVal S As Integer, ByVal V As Integer) As Color
    Return HSVtoColor(New HSV(H, S, V))
  End Function

  Public Shared Function HSVtoRGB(ByVal HSV As HSV) As RGB
    ' HSV contains values scaled as in the color wheel:
    ' that is, all from 0 to 255. 

    ' For this code to work, HSV.Hue needs
    ' to be scaled from 0 to 360 (it's the angle of the selected
    ' point within the circle). HSV.Saturation and HSV.Value must be 
    ' scaled to be between 0 and 1.

    Dim h As Decimal
    Dim s As Decimal
    Dim v As Decimal

    Dim r As Decimal
    Dim g As Decimal
    Dim b As Decimal

    ' Scale Hue to be between 0 and 360. Saturation
    ' and Value scale to be between 0 and 1.
    h = (HSV.Hue / 255D * 360) Mod 360
    s = HSV.Saturation / 255D
    v = HSV.Value / 255D

    If s = 0 Then
      ' If s is 0, all colors are the same.
      ' This is some flavor of gray.
      r = v
      g = v
      b = v
    Else
      Dim p As Decimal
      Dim q As Decimal
      Dim t As Decimal

      Dim fractionalSector As Decimal
      Dim sectorNumber As Integer
      Dim sectorPos As Decimal

      ' The color wheel consists of 6 sectors.
      ' Figure out which sector you're in.
      sectorPos = h / 60
      sectorNumber = CInt(Math.Floor(sectorPos))

      ' Get the fractional part of the sector.
      ' That is, how many degrees into the sector
      ' are you?
      fractionalSector = sectorPos - sectorNumber

      ' Calculate values for the three axes
      ' of the color. 
      p = v * (1 - s)
      q = v * (1 - (s * fractionalSector))
      t = v * (1 - (s * (1 - fractionalSector)))

      ' Assign the fractional colors to r, g, and b
      ' based on the sector the angle is in.
      Select Case sectorNumber
        Case 0
          r = v
          g = t
          b = p

        Case 1
          r = q
          g = v
          b = p

        Case 2
          r = p
          g = v
          b = t

        Case 3
          r = p
          g = q
          b = v

        Case 4
          r = t
          g = p
          b = v

        Case 5
          r = v
          g = p
          b = q
      End Select
    End If
    ' Return an RGB structure, with values scaled
    ' to be between 0 and 255.
    Return New RGB(CInt(r * 255), CInt(g * 255), CInt(b * 255))
  End Function

  Public Shared Function RGBtoHSV(ByVal RGB As RGB) As HSV
    ' In this function, R, G, and B values must be scaled 
    ' to be between 0 and 1.
    ' HSV.Hue will be a value between 0 and 360, and 
    ' HSV.Saturation and Value are between 0 and 1.
    ' The code must scale these to be between 0 and 255 for
    ' the purposes of this application.

    Dim min As Decimal
    Dim max As Decimal
    Dim delta As Decimal

    Dim r As Decimal = RGB.Red / 255D
    Dim g As Decimal = RGB.Green / 255D
    Dim b As Decimal = RGB.Blue / 255D

    Dim h As Decimal
    Dim s As Decimal
    Dim v As Decimal

    min = Math.Min(Math.Min(r, g), b)
    max = Math.Max(Math.Max(r, g), b)
    v = max
    delta = max - min
    If max = 0 Or delta = 0 Then
      ' R, G, and B must be 0, or all the same.
      ' In this case, S is 0, and H is undefined.
      ' Using H = 0 is as good as any...
      s = 0
      h = 0
    Else
      s = delta / max
      If r = max Then
        ' Between Yellow and Magenta
        h = (g - b) / delta
      ElseIf g = max Then
        ' Between Cyan and Yellow
        h = 2 + (b - r) / delta
      Else
        ' Between Magenta and Cyan
        h = 4 + (r - g) / delta
      End If

    End If
    ' Scale h to be between 0 and 360. 
    ' This may require adding 360, if the value
    ' is negative.
    h *= 60
    If h < 0 Then
      h += 360
    End If

    ' Scale to the requirements of this 
    ' application. All values are between 0 and 255.
    Return New HSV(CInt(h / 360 * 255), CInt(s * 255), CInt(v * 255))
  End Function
End Class


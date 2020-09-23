Attribute VB_Name = "modTYPES"
Public Type tpoint
    X As Single
    Y As Single
    ToMove As Boolean
End Type

Public Type tGRID
    T() As New clsTRIANG
End Type


Public SCALA As Single

Public Const MAXOUT = 400





Public Function IsInside(U, V, w)

'IsInside = IIf((U > 0) And (V > 0) And (U + V < 1), True, False)
IsInside = IIf((U > -0.00005) And (V > -0.00005) And (U + V < 1.0001), True, False)

End Function


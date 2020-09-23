VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior  = 0 'vbNone
MTSTransactionMode  = 0 'NotAnMTSObject
END
Attribute VB_Name = "clsTRIANG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False




Private P(1 To 3) As tpoint

Public Area As Single




Public Sub CalcArea()

Dim M As Single, Bm As Single
' Dim vX As Long, vY As Long
Dim vX As Single, vY As Single



vX = P(1).X - P(2).X
vY = P(1).Y - P(2).Y
If P(1).X <> P(2).X Then
    M = vY / vX
Else
    M = 1E+20
End If
Bm = Abs(-M * P(3).X + P(3).Y - P(1).Y + M * P(1).X) / Sqr(M * M + 1)

Area = 0.5 * Sqr(vX * vX + vY * vY) * Bm

End Sub

'Public Property Get GetArea() As Single
'GetArea = Area
'End Property

Public Property Get GetCoordX(wP) As Single
'If wP < 1 Or wP > 3 Then Exit Property
GetCoordX = P(wP).X

End Property
Public Property Get GetCoordY(wP) As Single
'If wP < 1 Or wP > 3 Then Exit Property
GetCoordY = P(wP).Y

End Property

Public Property Let SetCoordX(wP, Value As Single)
'If wP < 1 Or wP > 3 Then Exit Property
P(wP).X = Value
End Property
Public Property Let SetCoordY(wP, Value As Single)
'If wP < 1 Or wP > 3 Then Exit Property
P(wP).Y = Value
End Property

Public Property Get GetToMOVE(wP) As Boolean
GetToMOVE = P(wP).ToMove
End Property
Public Property Let SetToMOVE(wP, YesNo As Boolean)
P(wP).ToMove = YesNo
End Property

Public Function bToPOINT(T As clsTRIANG, U, V, W, ByRef rX As Single, ByRef rY As Single)

With T
    
    rX = (U * .GetCoordX(1) + V * .GetCoordX(2) + W * .GetCoordX(3))
    rY = (U * .GetCoordY(1) + V * .GetCoordY(2) + W * .GetCoordY(3))
    
End With

End Function




Public Function InterMedTria(T1 As clsTRIANG, T2 As clsTRIANG, Perc As Single) As clsTRIANG
Dim T As New clsTRIANG

Dim Perc2 As Single
Perc2 = 1 - Perc

For I = 1 To 3
    T.SetCoordX(I) = T1.GetCoordX(I) * Perc2 + T2.GetCoordX(I) * Perc
    T.SetCoordY(I) = T1.GetCoordY(I) * Perc2 + T2.GetCoordY(I) * Perc
Next

Set InterMedTria = T
Set T = Nothing

End Function


Public Sub DRAW(hdc As Long)
Dim X1 As Long
Dim Y1 As Long
Dim X2 As Long
Dim y2 As Long
Dim X3 As Long
Dim Y3 As Long

X1 = CLng(P(1).X)
Y1 = CLng(P(1).Y)
X2 = CLng(P(2).X)
y2 = CLng(P(2).Y)
X3 = CLng(P(3).X)
Y3 = CLng(P(3).Y)

FastLine hdc, X1, Y1, X2, y2, 1, vbRed
FastLine hdc, X2, y2, X3, Y3, 1, vbRed
FastLine hdc, X3, Y3, X1, Y1, 1, vbRed



End Sub


Public Function pToBARICENTRIC(T As clsTRIANG, X As Single, Y As Single, ByRef rU As Single, ByRef rV As Single, ByRef rW As Single)
Dim A As Single
Dim B As Single
Dim C As Single
Dim D As Single
Dim E As Single
Dim F As Single
Dim G As Single
Dim H As Single
Dim I As Single

Dim X1 As Single
Dim Y1 As Single
Dim X2 As Single
Dim y2 As Single
Dim X3 As Single
Dim Y3 As Single


X1 = T.GetCoordX(1)
Y1 = T.GetCoordY(1)
X2 = T.GetCoordX(2)
y2 = T.GetCoordY(2)
X3 = T.GetCoordX(3)
Y3 = T.GetCoordY(3)

A = X1 - X3
B = X2 - X3
C = X3 - X
D = Y1 - Y3
E = y2 - Y3
F = Y3 - Y
G = 0 'z1 - z3
H = 0 'z2 - z3
I = 0 'z3 - z

rU = (B * (F + I) - C * (E + H)) / (A * (E + H) - B * (D + G))
rV = (A * (F + I) - C * (D + G)) / (B * (D + G) - A * (E + H))


rW = 1 - rU - rV



End Function

Public Function GetBB(ByRef X1 As Single, ByRef Y1 As Single, ByRef X2 As Single, ByRef y2 As Single)
X1 = 9999999999#
X2 = -99999999999999#
Y1 = 999999999999#
y2 = -999999999999#

For I = 1 To 3
    
    If P(I).X < X1 Then X1 = P(I).X
    If P(I).Y < Y1 Then Y1 = P(I).Y
    If P(I).X > X2 Then X2 = P(I).X
    If P(I).Y > y2 Then y2 = P(I).Y
    
Next



End Function


VERSION 1.0 CLASS
BEGIN
MultiUse = -1 'True
Persistable = 0 'NotPersistable
DataBindingBehavior = 0 'vbNone
DataSourceBehavior  = 0 'vbNone
MTSTransactionMode  = 0 'NotAnMTSObject
END
Attribute VB_Name = "clsFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, ByRef lpBits As Any) As Long

Private T1Bytes() As Byte
Private T2Bytes() As Byte
Private OutBytes() As Byte

Private hBmp As Bitmap

Private pH As Integer
Private pW As Integer

Dim T1X As Single
Dim T2X As Single
Dim T1Y As Single
Dim T2Y As Single

Dim T3X As Single
Dim T3y As Single

Dim kX As Single
Dim kY As Single

Dim kX2 As Single
Dim kY2 As Single

Dim K2 As Single


Public Sub SetT1(pBoxPictureHandle)
Dim iRet As Long
'Get the bitmap header
iRet = GetObject(pBoxPictureHandle, Len(hBmp), hBmp)
'Resize to hold image data
ReDim T1Bytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
'Get the image data and store into bmpBytes array
iRet = GetBitmapBits(pBoxPictureHandle, hBmp.bmWidthBytes * hBmp.bmHeight, T1Bytes(0, 0, 0))

pW = hBmp.bmWidth - 1
pH = hBmp.bmHeight - 1

T1X = hBmp.bmWidth '- 1
T1Y = hBmp.bmHeight '- 1

'ReDim OutBytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To pW, 0 To pH) As Byte
'   Stop

End Sub
Public Sub SetT2(pBoxPictureHandle)
Dim iRet As Long
'Get the bitmap header
iRet = GetObject(pBoxPictureHandle, Len(hBmp), hBmp)
'Resize to hold image data
ReDim T2Bytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
'Get the image data and store into bmpBytes array
iRet = GetBitmapBits(pBoxPictureHandle, hBmp.bmWidthBytes * hBmp.bmHeight, T2Bytes(0, 0, 0))

T2X = hBmp.bmWidth '- 1
T2Y = hBmp.bmHeight '- 1


End Sub

Public Sub SetT3(pBoxPictureHandle)
Dim iRet As Long
'Get the bitmap header
iRet = GetObject(pBoxPictureHandle, Len(hBmp), hBmp)
'Resize to hold image data
ReDim OutBytes(0 To (hBmp.bmBitsPixel \ 8) - 1, 0 To hBmp.bmWidth - 1, 0 To hBmp.bmHeight - 1) As Byte
'Get the image data and store into bmpBytes array
iRet = GetBitmapBits(pBoxPictureHandle, hBmp.bmWidthBytes * hBmp.bmHeight, T2Bytes(0, 0, 0))


T3X = hBmp.bmWidth '- 1
T3y = hBmp.bmHeight '- 1

End Sub
Public Sub GetOut(pBoxPictureHandle)
Dim iRet As Long
'Set the new image data back onto pBox
iRet = SetBitmapBits(pBoxPictureHandle, hBmp.bmWidthBytes * hBmp.bmHeight, OutBytes(0, 0, 0))


End Sub



Private Sub Class_Terminate()
'Erase bmpBytes because we finished with it now
Erase T1Bytes
Erase T2Bytes
Erase OutBytes

End Sub





Public Sub MORPH(T1 As clsTRIANG, T2 As clsTRIANG, Perc As Single, Optional BLUR = False)
Dim X As Long
Dim Y As Long
Dim X2 As Long
Dim y2 As Long
Dim sX As Single
Dim sY As Single
Dim sX2 As Single
Dim sy2 As Single

Dim X3 As Long
Dim Y3 As Long
Dim U As Single
Dim V As Single
Dim W As Single
Dim R1 As Integer
Dim G1 As Integer
Dim B1 As Integer
Dim R2 As Integer
Dim G2 As Integer
Dim B2 As Integer
Dim R3 As Integer
Dim G3 As Integer
Dim B3 As Integer
Dim Perc2 As Single
Dim MED As New clsTRIANG
Dim XB1 As Single
Dim XB2 As Single
Dim yB1 As Single
Dim yB2 As Single

Perc2 = 1 - Perc

kX = T1X / T3X
kY = T1Y / T3y
kX2 = T2X / T3X
kY2 = T2Y / T3y


Set MED = T1.InterMedTria(T1, T2, Perc)

MED.GetBB XB1, yB1, XB2, yB2

For X3 = XB1 To XB2
    For Y3 = yB1 To yB2
        
        
        MED.pToBARICENTRIC MED, CSng(X3), CSng(Y3), U, V, W
        
        If IsInside(U, V, W) Then
            
            '                        Stop
            
            T1.bToPOINT T1, U, V, W, sX, sY
            T2.bToPOINT T2, U, V, W, sX2, sy2
            sX = sX * kX
            sY = sY * kY
            '                        Stop
            
            sX2 = sX2 * kX2
            sy2 = sy2 * kY2
            
            R1 = T1Bytes(2, sX, sY)
            G1 = T1Bytes(1, sX, sY)
            B1 = T1Bytes(0, sX, sY)
            
            R2 = T2Bytes(2, sX2, sy2)
            G2 = T2Bytes(1, sX2, sy2)
            B2 = T2Bytes(0, sX2, sy2)
            
            R3 = R1 * Perc2 + R2 * Perc
            G3 = G1 * Perc2 + G2 * Perc
            B3 = B1 * Perc2 + B2 * Perc
            
            OutBytes(2, X3, Y3) = R3
            OutBytes(1, X3, Y3) = G3
            OutBytes(0, X3, Y3) = B3
            
        End If
    Next
Next

Set MED = Nothing
End Sub

Public Sub DEFORM(T1 As clsTRIANG, T2 As clsTRIANG, Perc As Single, Optional BLUR = False)
Dim X As Long
Dim Y As Long
Dim X2 As Long
Dim y2 As Long
Dim X3 As Long
Dim Y3 As Long
Dim U(0 To 9) As Single
Dim V(0 To 9) As Single
Dim W(0 To 9) As Single
Dim R(0 To 9) As Integer
Dim G(0 To 9) As Integer
Dim B(0 To 9) As Integer
Dim Perc2 As Single
Dim MED As New clsTRIANG
Dim XB1 As Single
Dim XB2 As Single
Dim yB1 As Single
Dim yB2 As Single
Dim sX As Single
Dim sY As Single


'Stop

Perc2 = 1 - Perc

kX = T1X / T3X
kY = T1Y / T3y




Set MED = T1.InterMedTria(T1, T2, Perc)

MED.GetBB XB1, yB1, XB2, yB2

For X3 = XB1 To XB2
    For Y3 = yB1 To yB2
        
        
        MED.pToBARICENTRIC MED, CSng(X3), CSng(Y3), U(0), V(0), W(0)
        
        If BLUR Then
            MED.pToBARICENTRIC MED, CSng(X3 - 1), CSng(Y3 - 1), U(1), V(1), W(1)
            MED.pToBARICENTRIC MED, CSng(X3), CSng(Y3 - 1), U(2), V(2), W(2)
            MED.pToBARICENTRIC MED, CSng(X3 + 1), CSng(Y3 - 1), U(3), V(3), W(3)
            MED.pToBARICENTRIC MED, CSng(X3 + 1), CSng(Y3), U(4), V(4), W(4)
            MED.pToBARICENTRIC MED, CSng(X3 + 1), CSng(Y3 + 1), U(5), V(5), W(5)
            MED.pToBARICENTRIC MED, CSng(X3), CSng(Y3 + 1), U(6), V(6), W(6)
            MED.pToBARICENTRIC MED, CSng(X3 - 1), CSng(Y3 + 1), U(7), V(7), W(7)
            MED.pToBARICENTRIC MED, CSng(X3 - 1), CSng(Y3), U(8), V(8), W(8)
        End If
        
        
        
        If IsInside(U(0), V(0), W(0)) Then
            
            T1.bToPOINT T1, U(0), V(0), W(0), sX, sY
            'T2.bToPOINT T2, U, V, W, X2, y2
            X = CLng(sX * kX)
            Y = CLng(sY * kY)
            R(9) = T1Bytes(2, X, Y)
            G(9) = T1Bytes(1, X, Y)
            B(9) = T1Bytes(0, X, Y)
            
            If BLUR Then
                T1.bToPOINT T1, U(1), V(1), W(1), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(1) = T1Bytes(2, X, Y)
                G(1) = T1Bytes(1, X, Y)
                B(1) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(2), V(2), W(2), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(2) = T1Bytes(2, X, Y)
                G(2) = T1Bytes(1, X, Y)
                B(2) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(3), V(3), W(3), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(3) = T1Bytes(2, X, Y)
                G(3) = T1Bytes(1, X, Y)
                B(3) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(4), V(4), W(4), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(4) = T1Bytes(2, X, Y)
                G(4) = T1Bytes(1, X, Y)
                B(4) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(5), V(5), W(5), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(5) = T1Bytes(2, X, Y)
                G(5) = T1Bytes(1, X, Y)
                B(5) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(6), V(6), W(6), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(6) = T1Bytes(2, X, Y)
                G(6) = T1Bytes(1, X, Y)
                B(6) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(7), V(7), W(7), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(7) = T1Bytes(2, X, Y)
                G(7) = T1Bytes(1, X, Y)
                B(7) = T1Bytes(0, X, Y)
                T1.bToPOINT T1, U(8), V(8), W(8), sX, sY
                X = CLng(sX * kX)
                Y = CLng(sY * kY)
                R(8) = T1Bytes(2, X, Y)
                G(8) = T1Bytes(1, X, Y)
                B(8) = T1Bytes(0, X, Y)
                
                R(9) = (R(1) + R(2) + R(3) + R(4) + R(5) + R(6) + R(7) + R(8) + R(9)) \ 9
                G(9) = (G(1) + G(2) + G(3) + G(4) + G(5) + G(6) + G(7) + G(8) + G(9)) \ 9
                B(9) = (B(1) + B(2) + B(3) + B(4) + B(5) + B(6) + B(7) + B(8) + B(9)) \ 9
                
                
            End If
            
            
            
            OutBytes(2, X3, Y3) = R(9)
            OutBytes(1, X3, Y3) = G(9)
            OutBytes(0, X3, Y3) = B(9)
            
            
        End If
    Next
Next

Set MED = Nothing
End Sub


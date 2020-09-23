VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fMAIN 
   Caption         =   "Baricentric Coordinates  Morph and Deform"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   720
   ClientWidth     =   14190
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   946
   Begin VB.CommandButton Command5 
      Caption         =   "DEFORM ALL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12240
      TabIndex        =   8
      ToolTipText     =   "DEFORM ALL Sequence"
      Top             =   4920
      Width           =   1575
   End
   Begin VB.TextBox tITERA 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   12720
      TabIndex        =   15
      Text            =   "12"
      ToolTipText     =   "How Many Frames '"
      Top             =   2490
      Width           =   615
   End
   Begin VB.CommandButton cmdLOAD 
      Caption         =   "Load Project"
      Height          =   615
      Left            =   12960
      TabIndex        =   14
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton cmdSAVE 
      Caption         =   "Save Project"
      Height          =   615
      Left            =   11880
      TabIndex        =   13
      Top             =   7320
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   " DEFROM Last"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   12
      Top             =   4440
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CMD2 
      Left            =   6720
      Top             =   6240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pORG2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   6960
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   11
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.PictureBox pORG1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   5280
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   105
      TabIndex        =   10
      Top             =   7200
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CheckBox chBLUR 
      Caption         =   "Deform BLUR"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12240
      TabIndex        =   9
      Top             =   5760
      Width           =   1575
   End
   Begin VB.TextBox tLATE 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   12600
      TabIndex        =   7
      Text            =   "1"
      ToolTipText     =   "GRID LATE"
      Top             =   0
      Width           =   855
   End
   Begin VB.CommandButton Command4 
      Caption         =   "MORPH ALL"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   12240
      TabIndex        =   6
      ToolTipText     =   "MORPH ALL Sequence"
      Top             =   2760
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Create GRID"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   12600
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.PictureBox PIC3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3255
      Left            =   120
      ScaleHeight     =   217
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   337
      TabIndex        =   4
      Top             =   6360
      Width           =   5055
   End
   Begin VB.PictureBox PIC2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   4800
      Picture         =   "fMAIN.frx":0000
      ScaleHeight     =   401
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   321
      TabIndex        =   3
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "MORPH Middle"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12240
      TabIndex        =   2
      ToolTipText     =   "Test the Middle Frame"
      Top             =   1560
      Width           =   1575
   End
   Begin VB.PictureBox PIC1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   120
      ScaleHeight     =   225
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   169
      TabIndex        =   1
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   13560
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Menu cmdLOADP1 
      Caption         =   "Load Picture 1"
   End
   Begin VB.Menu cmdLOADP2 
      Caption         =   "Load Picture 2"
   End
End
Attribute VB_Name = "fMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Base 1

Dim T() As New clsTRIANG

Dim MED As New clsTRIANG

Dim P As tpoint
Dim FX As New clsFX

Dim GRID1() As New clsTRIANG

Dim GRID2() As New clsTRIANG

Dim moveX As Integer
Dim moveY As Integer
Dim MOVER As Boolean

Public gridLate As Integer

Dim K(1 To 2) As Single

Dim FilePic(1 To 2) As String


Private Sub cmdLOAD_Click()
Dim UB As Long
Dim V As String
Dim S As String

Open App.Path & "\PRJ.TXT" For Input As 1


Input #1, FilePic(1)
Input #1, FilePic(2)
Input #1, UB

ReDim GRID1(UB)
ReDim GRID2(UB)


For tr = 1 To UB
    For I = 1 To 3
        Input #1, V
        GRID1(tr).SetCoordX(I) = Val(V)
        Input #1, V
        GRID1(tr).SetCoordY(I) = Val(V)
        Input #1, V
        GRID2(tr).SetCoordX(I) = Val(V)
        Input #1, V
        GRID2(tr).SetCoordY(I) = Val(V)
        
        GRID1(tr).SetToMOVE(I) = False
        GRID2(tr).SetToMOVE(I) = False
        
        
    Next I
Next

Close 1

If InStr(1, FilePic(1), "\") = 0 Then
    FilePic(1) = App.Path & "\" & FilePic(1)
End If
If InStr(1, FilePic(2), "\") = 0 Then
    FilePic(2) = App.Path & "\" & FilePic(2)
End If


cmdLOADP1_A
DoEvents
cmdLOADP2_A
DoEvents

DrawGRIDS
End Sub

Private Sub cmdLOADP1_Click()
CMD2.Filter = "Image 1|*.bmp;*.jpg"
CMD2.DialogTitle = "Load P 1"
CMD2.Action = 1
If CMD2.FileName = "" Then Exit Sub

FilePic(1) = CMD2.FileName
cmdLOADP1_A
End Sub

Sub cmdLOADP1_A()

pORG1.Cls

pORG1.Picture = LoadPicture(FilePic(1))
pORG1.Refresh

If pORG1.Width > pORG1.Height Then
    PIC1.Width = MAXOUT
    PIC1.Height = MAXOUT / pORG1.Width * pORG1.Height
Else
    PIC1.Height = MAXOUT
    PIC1.Width = MAXOUT / pORG1.Height * pORG1.Width
End If

PIC1.Cls
SetStretchBltMode PIC1.hdc, STRETCHMODE
StretchBlt PIC1.hdc, 0, 0, PIC1.Width, PIC1.Height, _
        pORG1.hdc, 0, 0, pORG1.Width - 1, pORG1.Height, vbSrcCopy

PIC1.Refresh
SavePicture PIC1.Image, App.Path & "\tmpPIC1.bmp"
PIC1.Cls: PIC2.Cls: PIC3.Cls
PIC1.Picture = LoadPicture(App.Path & "\tmpPIC1.bmp")
PIC2.Width = PIC1.Width
PIC2.Height = PIC1.Height
PIC3.Width = PIC1.Width
PIC3.Height = PIC1.Height
PIC2.Left = PIC1.Left + PIC1.Width + 10
PIC3.Top = PIC1.Top + PIC1.Height + 10
PIC3.Left = PIC1.Left

pORG2.Refresh
SetStretchBltMode PIC2.hdc, STRETCHMODE
StretchBlt PIC2.hdc, 0, 0, PIC2.Width, PIC2.Height, _
        pORG2.hdc, 0, 0, pORG2.Width - 1, pORG2.Height, vbSrcCopy

PIC2.Refresh
SavePicture PIC2.Image, App.Path & "\tmpPIC2.bmp"
PIC2.Cls
PIC2.Picture = LoadPicture(App.Path & "\tmpPIC2.bmp")

End Sub

Private Sub cmdLOADP2_Click()
CMD2.Filter = "Image 1|*.bmp;*.jpg"
CMD2.DialogTitle = "Load P 1"
CMD2.Action = 1
If CMD2.FileName = "" Then Exit Sub

FilePic(2) = CMD2.FileName

cmdLOADP2_A

End Sub

Sub cmdLOADP2_A()

pORG2.Cls
pORG2.Picture = LoadPicture(FilePic(2))
pORG2.Refresh

If pORG2.Width > pORG2.Height Then
    PIC2.Width = MAXOUT
    PIC2.Height = MAXOUT / pORG2.Width * pORG2.Height
Else
    PIC2.Height = MAXOUT
    PIC2.Width = MAXOUT / pORG2.Height * pORG2.Width
End If

PIC2.Cls
SetStretchBltMode PIC2.hdc, STRETCHMODE
StretchBlt PIC2.hdc, 0, 0, PIC2.Width, PIC2.Height, _
        pORG2.hdc, 0, 0, pORG2.Width - 1, pORG2.Height, vbSrcCopy

PIC2.Refresh
SavePicture PIC2.Image, App.Path & "\tmpPIC2.bmp"
PIC1.Cls: PIC2.Cls: PIC3.Cls
PIC2.Picture = LoadPicture(App.Path & "\tmpPIC2.bmp")
PIC1.Width = PIC2.Width
PIC1.Height = PIC2.Height
PIC3.Width = PIC2.Width
PIC3.Height = PIC2.Height
PIC2.Left = PIC1.Left + PIC1.Width + 10
PIC3.Top = PIC1.Top + PIC1.Height + 10
PIC3.Left = PIC1.Left

pORG1.Refresh
SetStretchBltMode PIC1.hdc, STRETCHMODE
StretchBlt PIC1.hdc, 0, 0, PIC1.Width, PIC1.Height, _
        pORG1.hdc, 0, 0, pORG1.Width - 1, pORG1.Height, vbSrcCopy

PIC1.Refresh
SavePicture PIC1.Image, App.Path & "\tmpPIC1.bmp"
PIC1.Cls
PIC1.Picture = LoadPicture(App.Path & "\tmpPIC1.bmp")


End Sub
Private Sub cmdSAVE_Click()
Open App.Path & "\PRJ.TXT" For Output As 1

Print #1, FilePic(1)
Print #1, FilePic(2)
Print #1, UBound(GRID1())

For tr = 1 To UBound(GRID1())
    For I = 1 To 3
        Print #1, Replace(GRID1(tr).GetCoordX(I), ",", ".")
        Print #1, Replace(GRID1(tr).GetCoordY(I), ",", ".")
        Print #1, Replace(GRID2(tr).GetCoordX(I), ",", ".")
        Print #1, Replace(GRID2(tr).GetCoordY(I), ",", ".")
    Next I
Next

Close 1

End Sub

Private Sub Command1_Click()
'BARICENTRIC TEST
Dim U As Single
Dim V As Single
Dim W As Single
Dim XX As Single
Dim YY As Single

For I = 1 To 2
    For C = 1 To 3
        T(I).SetCoordX(C) = Rnd * PIC1.Width - 1
        T(I).SetCoordY(C) = Rnd * PIC1.Height - 1
        T(I).CalcArea
    Next C
Next


T(1).DRAW PIC3.hdc
T(2).DRAW PIC3.hdc

PIC3.Cls

P.X = T(1).GetCoordX(1) + 10
P.Y = T(1).GetCoordY(1) + 10

T(1).pToBARICENTRIC T(1), P.X, P.Y, U, V, W


For Pe = 0 To 1 Step 0.2
    
    Set MED = T(1).InterMedTria(T(1), T(2), CSng(Pe))
    MED.CalcArea
    
    MED.DRAW PIC3.hdc
    
    MED.bToPOINT MED, U, V, W, XX, YY
    
    MyCircle PIC3.hdc, XX, YY, 2, 2, vbGreen
    
    
Next Pe
PIC3.Refresh


End Sub

Private Sub Command2_Click()
'MIDDLe morph



PIC1.Cls
PIC2.Cls

FX.SetT1 pORG1.Picture.Handle

FX.SetT2 pORG2.Picture.Handle

FX.SetT3 PIC1.Image.Handle



DrawIntermed 0.5, 0



DrawGRIDS

End Sub




Private Sub Command3_Click()

Dim StepX
Dim StepY

ReDim GRID1(gridLate * gridLate * 2)
ReDim GRID2(gridLate * gridLate * 2)
StepX = (PIC1.Width - 14) / gridLate
StepY = (PIC1.Height - 14) / gridLate

tr = 0
For Y = 7 To PIC1.Height - 7 - 1 Step StepY
    For X = 7 To PIC1.Width - 7 - 1 Step StepX
        
        tr = tr + 1
        With GRID1(tr)
            .SetCoordX(1) = X
            .SetCoordY(1) = Y
            .SetCoordX(2) = X + StepX
            .SetCoordY(2) = Y
            .SetCoordX(3) = X
            .SetCoordY(3) = Y + StepY
        End With
        tr = tr + 1
        With GRID1(tr)
            .SetCoordX(1) = X + StepX
            .SetCoordY(1) = Y
            .SetCoordX(2) = X + StepX
            .SetCoordY(2) = Y + StepY
            .SetCoordX(3) = X
            .SetCoordY(3) = Y + StepY
        End With
        
        
    Next X
Next Y

PIC1.Cls
PIC2.Cls
For tr = 1 To UBound(GRID1)
    For I = 1 To 3
        GRID2(tr).SetCoordX(I) = GRID1(tr).GetCoordX(I)
        GRID2(tr).SetCoordY(I) = GRID1(tr).GetCoordY(I)
    Next
    GRID1(tr).DRAW PIC1.hdc
    GRID2(tr).DRAW PIC2.hdc
Next
PIC1.Refresh
PIC2.Refresh


End Sub

Private Sub Command4_Click()
Dim P As Single
Dim nFR As Single


If Dir(App.Path & "\frames\*.bmp") <> "" Then Kill App.Path & "\frames\*.bmp"

nFR = Val(tITERA)

PIC1.Cls
PIC2.Cls

FX.SetT1 pORG1.Picture.Handle

FX.SetT2 pORG2.Picture.Handle

FX.SetT3 PIC1.Image.Handle


For P = 0 To 1.0001 Step 1 / (nFR - 1)
    
    DrawIntermed P, 0
    SavePicture PIC3.Image, App.Path & "\frames\F" & Format(P * 1000, "000000") & ".bmp"
Next

DrawGRIDS



End Sub

Private Sub Command5_Click()
Dim P As Single
Dim nFR As Single

If Dir(App.Path & "\frames\*.bmp") <> "" Then Kill App.Path & "\frames\*.bmp"

nFR = Val(tITERA)

PIC1.Cls
PIC2.Cls

FX.SetT1 pORG1.Image.Handle
FX.SetT2 pORG2.Image.Handle
FX.SetT3 PIC1.Image.Handle

For P = 0 To 1.0001 Step 1 / (nFR - 1)
    
    DrawIntermed P, 1
    SavePicture PIC3.Image, App.Path & "\frames\F" & Format(P * 1000, "000000") & ".bmp"
Next

DrawGRIDS


End Sub

Private Sub Command6_Click()
'view last
Dim P As Single

PIC1.Cls
PIC2.Cls
FX.SetT1 pORG1.Image.Handle
FX.SetT2 pORG2.Image.Handle

FX.SetT3 PIC1.Image.Handle
P = 1
DrawIntermed P, 1
SavePicture PIC3.Image, App.Path & "\frames\F" & Format(P * 1000, "000000") & ".bmp"


DrawGRIDS

End Sub

Private Sub Form_Load()
SCALA = 1
Me.WindowState = 2


Randomize Timer

ReDim T(2)



If Dir(App.Path & "\Frames", vbDirectory) = "" Then MkDir App.Path & "\Frames"

tLATE_Change
Command3_Click

PIC3.Width = PIC1.Width
PIC3.Height = PIC1.Height

cmdLOAD_Click

End Sub

Private Sub Form_Unload(Cancel As Integer)
'End

End Sub

Private Sub PIC1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dmin
Dim dx
Dim dy
Dim DIS() As Single

Dmin = 1E+20

MOVER = False


If Button = 1 Then
    
    
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            GRID1(tria).SetToMOVE(I) = False
            X2 = GRID1(tria).GetCoordX(I)
            y2 = GRID1(tria).GetCoordY(I)
            dx = X2 - X
            dy = y2 - Y
            D = Sqr(dx * dx + dy * dy)
            If D < Dmin Then Dmin = D
        Next
    Next
    
    If Dmin > 15 Then Exit Sub
    MOVER = True
    
    
    'Stop
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            X2 = GRID1(tria).GetCoordX(I)
            y2 = GRID1(tria).GetCoordY(I)
            dx = X2 - X
            dy = y2 - Y
            D = Sqr(dx * dx + dy * dy)
            If D = Dmin Then
                GRID1(tria).SetToMOVE(I) = True
                Debug.Print tria & "  " & I
                moveX = X
                moveY = Y
            End If
        Next
    Next
End If 'button=1
'-------------------------------------------------------------
If Button = 2 Then
    Dim MyT As Integer
    Dim MyT2 As Integer
    
    Dim bU As Single
    Dim bV As Single
    Dim bW As Single
    
    For tria = 1 To UBound(GRID1)
        GRID1(tria).pToBARICENTRIC GRID1(tria), CSng(X), CSng(Y), bU, bV, bW
        
        '    If IsInside(bU, bV, bW) Then MyT = tria: MsgBox bU & vbTab & bV & vbTab & bW
        If IsInside(bU, bV, bW) Or Abs(bU) < 0.05 Or Abs(bV) < 0.05 Or Abs(bW) < 0.05 Then
            If MyT = 0 Then
                MyT = tria
            Else
                MyT2 = tria
            End If
            
        End If
        
    Next
    MsgBox MyT & vbTab & MyT2
    
    Debug.Print MyT
    If MyT <> 0 Then
        If MyT2 = 0 Then
            U = UBound(GRID1)
            ReDim Preserve GRID1(U + 2)
            ReDim Preserve GRID2(U + 2)
            
            GRID1(U + 1).SetCoordX(1) = GRID1(MyT).GetCoordX(1)
            GRID1(U + 1).SetCoordY(1) = GRID1(MyT).GetCoordY(1)
            GRID1(U + 1).SetCoordX(2) = GRID1(MyT).GetCoordX(2)
            GRID1(U + 1).SetCoordY(2) = GRID1(MyT).GetCoordY(2)
            GRID1(U + 1).SetCoordX(3) = CSng(X)
            GRID1(U + 1).SetCoordY(3) = CSng(Y)
            
            GRID2(U + 1).SetCoordX(1) = GRID2(MyT).GetCoordX(1)
            GRID2(U + 1).SetCoordY(1) = GRID2(MyT).GetCoordY(1)
            GRID2(U + 1).SetCoordX(2) = GRID2(MyT).GetCoordX(2)
            GRID2(U + 1).SetCoordY(2) = GRID2(MyT).GetCoordY(2)
            GRID2(U + 1).SetCoordX(3) = CSng(X)
            GRID2(U + 1).SetCoordY(3) = CSng(Y)
            '-
            GRID1(U + 2).SetCoordX(1) = GRID1(MyT).GetCoordX(2)
            GRID1(U + 2).SetCoordY(1) = GRID1(MyT).GetCoordY(2)
            GRID1(U + 2).SetCoordX(2) = GRID1(MyT).GetCoordX(3)
            GRID1(U + 2).SetCoordY(2) = GRID1(MyT).GetCoordY(3)
            GRID1(U + 2).SetCoordX(3) = CSng(X)
            GRID1(U + 2).SetCoordY(3) = CSng(Y)
            GRID2(U + 2).SetCoordX(1) = GRID2(MyT).GetCoordX(2)
            GRID2(U + 2).SetCoordY(1) = GRID2(MyT).GetCoordY(2)
            GRID2(U + 2).SetCoordX(2) = GRID2(MyT).GetCoordX(3)
            GRID2(U + 2).SetCoordY(2) = GRID2(MyT).GetCoordY(3)
            GRID2(U + 2).SetCoordX(3) = CSng(X)
            GRID2(U + 2).SetCoordY(3) = CSng(Y)
            '-
            GRID1(MyT).SetCoordX(2) = CSng(X)
            GRID1(MyT).SetCoordY(2) = CSng(Y)
            GRID2(MyT).SetCoordX(2) = CSng(X)
            GRID2(MyT).SetCoordY(2) = CSng(Y)
            '-
            
        Else
            
        End If 'MyT2 = 0
        
    Else 'myt=0
        Dmin = 1000000
        
        For tria = 1 To UBound(GRID1)
            GRID1(tria).pToBARICENTRIC GRID1(tria), CSng(X), CSng(Y), bU, bV, bW
            '    If IsInside(bU, bV, bW) Then MyT = tria: MsgBox bU & vbTab & bV & vbTab & bW
            MsgBox bU & vbTab & bV & vbTab & bW
            If Abs(bU) + Abs(bV) + Abs(bW) < Dmin Then
                Dmin = Abs(bU) + Abs(bV) + Abs(bW)
                MyT = tria
            End If
        Next
        MsgBox MyT
    End If 'MyT <> 0
    
    PIC1.Cls
    PIC2.Cls
    For tria = 1 To UBound(GRID1)
        GRID1(tria).DRAW PIC1.hdc
        GRID2(tria).DRAW PIC2.hdc
    Next
    PIC1.Refresh
    PIC2.Refresh
    
End If 'Button = 2

End Sub



Private Sub PIC1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)




If Button = 1 And MOVER Then
    
    If X < 5 Then X = 5
    If Y < 5 Then Y = 5
    If X > PIC1.Width - 5 Then X = PIC1.Width - 5
    If Y > PIC1.Height - 5 Then Y = PIC1.Height - 5
    
    
    moveX = X
    moveY = Y
    
    PIC1.Cls
    PIC2.Cls
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            If GRID1(tria).GetToMOVE(I) Then
                
                GRID1(tria).SetCoordX(I) = moveX
                GRID1(tria).SetCoordY(I) = moveY
                
                MyCircle PIC2.hdc, GRID2(tria).GetCoordX(I), GRID2(tria).GetCoordY(I), 4, 4, vbGreen
                
            End If
        Next
        GRID1(tria).DRAW PIC1.hdc
        GRID2(tria).DRAW PIC2.hdc
        MyCircle PIC1.hdc, moveX, moveY, 4, 4, vbGreen
        
        
    Next
    PIC1.Refresh
    PIC2.Refresh
End If

End Sub

Private Sub PIC1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVER = False

End Sub

''''''''
Private Sub PIC2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Dmin
Dim dx
Dim dy
Dmin = 1E+20
MOVER = False
If Button = 1 Then
    
    
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            GRID2(tria).SetToMOVE(I) = False
            X2 = GRID2(tria).GetCoordX(I)
            y2 = GRID2(tria).GetCoordY(I)
            dx = X2 - X
            dy = y2 - Y
            D = Sqr(dx * dx + dy * dy)
            If D < Dmin Then Dmin = D
        Next
    Next
    If Dmin > 15 Then Exit Sub
    MOVER = True
    
    
    'Stop
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            X2 = GRID2(tria).GetCoordX(I)
            y2 = GRID2(tria).GetCoordY(I)
            dx = X2 - X
            dy = y2 - Y
            D = Sqr(dx * dx + dy * dy)
            If D = Dmin Then
                GRID2(tria).SetToMOVE(I) = True
                Debug.Print tria & "  " & I
                moveX = X
                moveY = Y
            End If
        Next
    Next
End If 'button=1

End Sub



Private Sub PIC2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)


If Button = 1 And MOVER Then
    
    
    If X < 5 Then X = 5
    If Y < 5 Then Y = 5
    If X > PIC2.Width - 5 Then X = PIC2.Width - 5
    If Y > PIC2.Height - 5 Then Y = PIC2.Height - 5
    
    moveX = X
    moveY = Y
    
    PIC1.Cls
    PIC2.Cls
    For tria = 1 To UBound(GRID1)
        For I = 1 To 3
            If GRID2(tria).GetToMOVE(I) Then
                
                GRID2(tria).SetCoordX(I) = moveX
                GRID2(tria).SetCoordY(I) = moveY
                
                MyCircle PIC1.hdc, GRID1(tria).GetCoordX(I), GRID1(tria).GetCoordY(I), 4, 4, vbGreen
                
            End If
        Next
        GRID1(tria).DRAW PIC1.hdc
        GRID2(tria).DRAW PIC2.hdc
        MyCircle PIC2.hdc, moveX, moveY, 4, 4, vbGreen
        
    Next
    PIC1.Refresh
    PIC2.Refresh
End If

End Sub

Private Sub PIC2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MOVER = False

End Sub

Public Sub DrawIntermed(Perc As Single, MODE)
Me.Caption = Perc
DoEvents

If MODE = 0 Then
    For tr = 1 To UBound(GRID1)
        FX.MORPH GRID1(tr), GRID2(tr), Perc
    Next
Else
    For tr = 1 To UBound(GRID1)
        'Stop
        
        FX.DEFORM GRID1(tr), GRID2(tr), Perc, chBLUR
    Next
End If


FX.GetOut PIC3.Image.Handle
PIC3.Refresh


End Sub

Private Sub tLATE_Change()
gridLate = Val(tLATE)

End Sub

Public Sub SETSCALA(sV)
PIC1.Width = PIC1.Width * sV
PIC1.Height = PIC1.Height * sV
PIC2.Width = PIC2.Width * sV
PIC2.Height = PIC2.Height * sV
PIC3.Width = PIC3.Width * sV
PIC3.Height = PIC3.Height * sV

PIC1.Refresh

PIC3.Refresh

For tr = 1 To UBound(GRID1)
    For I = 1 To 3
        GRID1(tr).SetCoordX(I) = GRID1(tr).GetCoordX(I) * SCALA
        GRID1(tr).SetCoordY(I) = GRID1(tr).GetCoordY(I) * SCALA
        GRID2(tr).SetCoordX(I) = GRID2(tr).GetCoordX(I) * SCALA
        GRID2(tr).SetCoordY(I) = GRID2(tr).GetCoordY(I) * SCALA
    Next
Next

End Sub

Sub DrawGRIDS()
For tr = 1 To UBound(GRID1)
    GRID1(tr).DRAW PIC1.hdc
    GRID2(tr).DRAW PIC2.hdc
Next
PIC1.Refresh
PIC2.Refresh
End Sub

VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bubbulator"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   385
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   621
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsv 
      Caption         =   "<---"
      Height          =   255
      Left            =   4680
      TabIndex        =   8
      Top             =   4680
      Width           =   495
   End
   Begin VB.PictureBox stats 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   4680
      ScaleHeight     =   615
      ScaleWidth      =   4575
      TabIndex        =   10
      Top             =   5040
      Width           =   4575
   End
   Begin VB.PictureBox dest 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   4680
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   9
      Top             =   120
      Width           =   4500
   End
   Begin MSComDlg.CommonDialog cdc 
      Left            =   840
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open 300x300 Image"
      Filter          =   "Still Image (*.bmp; *.jpg)|*.bmp;*.jpg"
   End
   Begin VB.Frame frmopt 
      Caption         =   "Options"
      Height          =   975
      Left            =   120
      TabIndex        =   1
      Top             =   4680
      Width           =   4455
      Begin VB.CheckBox chkinout 
         Alignment       =   1  'Right Justify
         Caption         =   "Bubble Out"
         Height          =   195
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox txtrs 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "10"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox txtrn 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Text            =   "50"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Create 
         Caption         =   "Create Bubble"
         Height          =   375
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Bubble Strength:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bubble Size:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.PictureBox source 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4500
      Left            =   120
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   300
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   300
      TabIndex        =   0
      Top             =   120
      Width           =   4500
      Begin VB.CommandButton cmdnewimage 
         Caption         =   "Load Image"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   4200
         Width           =   1215
      End
      Begin VB.Line xpos 
         DrawMode        =   6  'Mask Pen Not
         X1              =   72
         X2              =   72
         Y1              =   0
         Y2              =   300
      End
      Begin VB.Line ypos 
         DrawMode        =   6  'Mask Pen Not
         X1              =   0
         X2              =   300
         Y1              =   224
         Y2              =   224
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Pi As Double = 3.14159265358979
Const Pi2 As Double = 6.28318530717958
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

Dim TickCount As Long
Dim t As Long, fn As Long

Private Sub cmdnewimage_Click()
    cdc.ShowOpen
    If cdc.FileName <> "" Then source.Picture = LoadPicture(cdc.FileName)
End Sub

Private Sub cmdsv_Click()
    source.Picture = dest.Image
End Sub

Private Sub Create_Click()
    stats.Cls
    DoEvents
    Dim X As Long, Y As Long, a As Double, r As Double, re As Double, w As Long, s As String, rn As Double, rs As Double, RCX As Single, RCY As Single, inout As Boolean
    Dim x2 As Single, y2 As Single
    RCX = xpos.x1
    RCY = ypos.y1
    rn = txtrn.Text
    rs = txtrs.Text
    inout = chkinout.Value
    d = 300
    TickCount = GetTickCount
    ReDim SourceBuffer.Bits(3, 299, 299)
    With SourceBuffer.Header
      .biSize = 40
      .biWidth = 300
      .biHeight = -300
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 3 * d * d
    End With
     
    ReDim DestBuffer.Bits(3, 299, 299)
    With DestBuffer.Header
      .biSize = 40
      .biWidth = 300
      .biHeight = -300
      .biPlanes = 1
      .biBitCount = 32
      .biSizeImage = 3 * d * d
    End With

    GetDIBits source.hdc, source.Image.Handle, 0, 300, SourceBuffer.Bits(0, 0, 0), SourceBuffer, 0&
            For Y = 0 To 299
                For X = 0 To 299
                    a = ATan2(RCY - Y, RCX - X)
                    r = Sqr((RCX - X) * (RCX - X) + (RCY - Y) * (RCY - Y))
                    If r <= rn Then
                        If inout = True Then
                            re = r + Sin((r / rn) * Pi + Pi) * rs
                        ElseIf inout = False Then
                            re = r - Sin((r / rn) * Pi + Pi) * rs
                        End If
                    Else
                        re = r
                    End If
                    x2 = Within(RCX - (re * Cos(a)), 0, 299)
                    y2 = Within(RCY + (re * Sin(a)), 0, 299)
                    DestBuffer.Bits(0, X, Y) = SourceBuffer.Bits(0, x2, y2)
                    DestBuffer.Bits(1, X, Y) = SourceBuffer.Bits(1, x2, y2)
                    DestBuffer.Bits(2, X, Y) = SourceBuffer.Bits(2, x2, y2)
                Next X
            Next Y
        SetDIBits dest.hdc, dest.Image.Handle, 0, 300, DestBuffer.Bits(0, 0, 0), DestBuffer, 0&
        DoEvents
    stats.Print "Operation took " & (GetTickCount - TickCount) / 1000 & " seconds."
End Sub

Private Sub source_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    source_MouseMove Button, Shift, X, Y
End Sub

Private Sub source_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        xpos.x1 = X
        xpos.x2 = X
        ypos.y1 = Y
        ypos.y2 = Y
    End If
End Sub

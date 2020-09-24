VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   DrawWidth       =   3
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "ÁâÐÎ"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3840
      TabIndex        =   4
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "¾íÖá"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1980
      TabIndex        =   3
      Top             =   7320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "°ÙÒ¶´°"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   120
      TabIndex        =   2
      Top             =   7320
      Width           =   1815
   End
   Begin VB.PictureBox picdest 
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   4980
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   9660
   End
   Begin VB.PictureBox picsour 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   7260
      Left            =   0
      Picture         =   "Form1.frx":9BCC
      ScaleHeight     =   7200
      ScaleWidth      =   9600
      TabIndex        =   1
      Top             =   0
      Width           =   9660
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'×÷Õß: WXJ_Lake
'ÍøÖ·£ºwww.archtide.com


Sub Form_Load()
picsour.Move 0, 0
picsour.ScaleMode = 3
picsour.AutoRedraw = True

picdest.Move 0, 0, picsour.Width, picsour.Height
picdest.AutoRedraw = False
End Sub


Private Sub Command1_Click()
Dim i As Long, j As Long
Dim H As Long, W As Long
Dim T As Long
Dim ScanLines As Long

picdest.Cls
H = picsour.ScaleHeight
W = picsour.ScaleWidth
ScanLines = 25
T = GetTickCount

For i = 0 To (ScanLines - 1)
  For j = i To W Step ScanLines
    BitBlt picdest.hdc, j, 0, 1, H, picsour.hdc, j, 0, vbSrcCopy
  Next j
  DoEvents
  Do: Loop Until GetTickCount - T > 50
  T = GetTickCount
Next i

End Sub

Private Sub Command2_Click()
Dim i As Long, j As Long
Dim H As Long, W As Long
Dim T As Long
Dim ScanLines As Long

picdest.Cls
H = picsour.ScaleHeight
W = picsour.ScaleWidth
ScanLines = 25
T = GetTickCount

While i < W
  BitBlt picdest.hdc, i, 0, ScanLines, H, picsour.hdc, i, 0, vbSrcCopy
  i = i + ScanLines
  DoEvents
  Do: Loop Until GetTickCount - T > 50
  T = GetTickCount
Wend

End Sub

Private Sub Command3_Click()
Dim Bitmap1 As Long, Bitmap2 As Long
Dim Buffer1 As Long, Buffer2 As Long
Dim hDC1 As Long, hDC2 As Long
Dim hBuffer1DC As Long, hBuffer2DC As Long
Dim hBrush1 As Long, hBrush2 As Long
Dim hRgn1 As Long, hRgn2 As Long
Dim Axis(3) As POINTAPI
Dim i As Long, T As Long
Dim ScanLines As Long

picdest.Cls

hDC1 = CreateCompatibleDC(0)
Bitmap1 = CreateCompatibleBitmap(picsour.hdc, 640, 480)
SelectObject hDC1, Bitmap1
hBrush1 = CreateSolidBrush(0)
SelectObject hDC1, hBrush1
SetPolyFillMode hDC1, ALTERNATE

hDC2 = CreateCompatibleDC(0)
Bitmap2 = CreateCompatibleBitmap(picdest.hdc, 640, 480)
SelectObject hDC2, Bitmap2
hBrush2 = CreateSolidBrush(0)
SelectObject hDC2, hBrush2
SetPolyFillMode hDC2, ALTERNATE

hBuffer1DC = CreateCompatibleDC(0)
Buffer1 = CreateCompatibleBitmap(picdest.hdc, 640, 480)
SelectObject hBuffer1DC, Buffer1

hBuffer2DC = CreateCompatibleDC(0)
Buffer2 = CreateCompatibleBitmap(picdest.hdc, 640, 480)
SelectObject hBuffer2DC, Buffer2

BitBlt hDC1, 0, 0, 640, 480, picsour.hdc, 0, 0, vbSrcCopy
BitBlt hBuffer2DC, 0, 0, 640, 480, picdest.hdc, 0, 0, vbSrcCopy

Axis(0).X = 0
Axis(0).Y = 0
Axis(1).X = 0
Axis(1).Y = 480
Axis(2).X = 640
Axis(2).Y = 480
Axis(3).X = 640
Axis(3).Y = 0
hRgn1 = CreatePolygonRgn(Axis(0), 4, ALTERNATE)

Axis(0).X = 320
Axis(0).Y = 240
ScanLines = 25
i = 1

While Axis(0).X > -240
  Axis(0).X = 320 - ScanLines * i
  Axis(0).Y = 240
  Axis(1).X = 320
  Axis(1).Y = 240 - ScanLines * i
  Axis(2).X = 320 + ScanLines * i - 1
  Axis(2).Y = 240
  Axis(3).X = 320
  Axis(3).Y = 240 + ScanLines * i
  Polygon hDC1, Axis(0), 4

  BitBlt hDC2, 0, 0, 640, 480, hBuffer2DC, 0, 0, vbSrcCopy
  hRgn2 = CreatePolygonRgn(Axis(0), 4, ALTERNATE)
  CombineRgn hRgn2, hRgn1, hRgn2, RGN_XOR
  FillRgn hDC2, hRgn2, hBrush2
  DeleteObject hRgn2
  
  BitBlt hBuffer1DC, 0, 0, 640, 480, hDC1, 0, 0, vbSrcCopy
  BitBlt hBuffer1DC, 0, 0, 640, 480, hDC2, 0, 0, vbSrcPaint
  BitBlt picdest.hdc, 0, 0, 640, 480, hBuffer1DC, 0, 0, vbSrcCopy
  
  i = i + 1
  
  DoEvents
  Do: Loop Until GetTickCount - T > 50
  T = GetTickCount

Wend

DeleteObject hBrush1
DeleteObject hBrush2
DeleteObject Bitmap1
DeleteObject Bitmap2
DeleteObject Buffer1
DeleteObject Buffer2

DeleteDC hDC1
DeleteDC hDC2
DeleteDC hBuffer1DC
DeleteDC hBuffer2DC

DeleteObject hRgn1
End Sub



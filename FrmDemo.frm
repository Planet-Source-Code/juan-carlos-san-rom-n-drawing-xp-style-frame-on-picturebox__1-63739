VERSION 5.00
Begin VB.Form FrmDemo 
   Caption         =   "XP style frame"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5745
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2505
      Left            =   150
      ScaleHeight     =   2505
      ScaleWidth      =   4995
      TabIndex        =   0
      Top             =   210
      Width           =   4995
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3270
         TabIndex        =   8
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3270
         TabIndex        =   6
         Top             =   540
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   4
         Top             =   1020
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1110
         TabIndex        =   2
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1:"
         Height          =   195
         Left            =   2670
         TabIndex        =   9
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1:"
         Height          =   195
         Left            =   2670
         TabIndex        =   7
         Top             =   600
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1:"
         Height          =   195
         Left            =   510
         TabIndex        =   5
         Top             =   1080
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "The simplest way of drawing xp style frame on picturebox"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   450
         TabIndex        =   3
         Top             =   1980
         Width           =   4005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Label1:"
         Height          =   195
         Left            =   510
         TabIndex        =   1
         Top             =   600
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function RoundRect Lib "gdi32" _
    (ByVal hDC As Long, ByVal Left As Long, ByVal Top As Long, _
     ByVal Right As Long, ByVal Bottom As Long, ByVal EllipseWidth As Long, _
     ByVal EllipseHeight As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hDC As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, lpDrawTextParams As Any) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private lpp As POINTAPI

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_RIGHT = &H2
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_SINGLELINE = &H20
Private Const DT_NOCLIP = &H100
Private Const DT_LEFT = &H0


Private Sub Form_Load()
    'just put a picturebox in your form and add other controls in it
    'then put this line and that's all
    PaintFrame Picture1, "XP style frame" 'name of picturbox and caption
End Sub

Private Sub PaintFrame(Pic As PictureBox, p_Caption As String)
    Dim p_LenghOfCaption As Long, p_HeightOfCaption As Long
    Dim p_indentation As Long, p_offset As Long, R As RECT, space As Long
    
    Pic.ScaleMode = vbPixels
    Pic.Picture = LoadPicture()
    p_LenghOfCaption = Pic.TextWidth(p_Caption)
    p_HeightOfCaption = Pic.TextHeight(p_Caption)
    p_offset = 1
    p_indentation = 12
    space = 4
    
    'define border color
    Pic.ForeColor = RGB(170, 170, 170)
    
    'paint rounded rectangle
    RoundRect Pic.hDC, p_offset, p_offset + p_HeightOfCaption / 2, Pic.ScaleWidth - p_offset, Pic.ScaleHeight - p_offset, 8&, 8&
    
    'define caption rectangle
    SetRect R, p_offset + p_indentation + space, p_offset, p_offset + p_indentation + p_LenghOfCaption + 2 * space, p_offset + p_HeightOfCaption
    
    'Set line color
    Pic.ForeColor = Pic.BackColor
    MoveToEx Pic.hDC, p_offset + p_indentation + space, p_offset + p_HeightOfCaption / 2, lpp
    LineTo Pic.hDC, p_offset + p_indentation + p_LenghOfCaption + 2 * space, p_offset + p_HeightOfCaption / 2
    
    'Set text color
    Pic.ForeColor = &HCF3603
    
    'Draw frame caption
    DrawTextEx Pic.hDC, p_Caption, Len(p_Caption), R, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER, ByVal 0&

    Pic.ScaleMode = vbTwips
    Pic.Picture = Pic.Image
End Sub

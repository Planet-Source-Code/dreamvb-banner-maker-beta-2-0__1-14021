VERSION 5.00
Begin VB.Form frmChooseColour 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Change ForeColour"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4890
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   201
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   326
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Apply"
      Height          =   405
      Left            =   3330
      TabIndex        =   8
      Top             =   2505
      Width           =   1305
   End
   Begin VB.PictureBox DisplayCol 
      Height          =   1035
      Left            =   2760
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   7
      Top             =   1320
      Width           =   1920
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   2
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "0"
      Top             =   825
      Width           =   690
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   1
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0"
      Top             =   495
      Width           =   690
   End
   Begin VB.TextBox RgbVal 
      Height          =   285
      Index           =   0
      Left            =   3960
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "0"
      Top             =   195
      Width           =   690
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H80000008&
      BorderStyle     =   0  'None
      Height          =   2550
      Left            =   75
      Picture         =   "frmChooseColour.frx":0000
      ScaleHeight     =   170
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   164
      TabIndex        =   0
      Top             =   75
      Width           =   2460
      Begin VB.Shape Shape1 
         Height          =   120
         Left            =   1215
         Shape           =   3  'Circle
         Top             =   1260
         Width           =   75
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   4
      X2              =   4
      Y1              =   5
      Y2              =   175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   4
      X2              =   170
      Y1              =   5
      Y2              =   5
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000003&
      Height          =   2565
      Left            =   60
      Top             =   75
      Width           =   2490
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Blue"
      Height          =   195
      Left            =   3465
      TabIndex        =   3
      Top             =   855
      Width           =   315
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Green"
      Height          =   195
      Left            =   3450
      TabIndex        =   2
      Top             =   540
      Width           =   435
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Red"
      Height          =   195
      Left            =   3465
      TabIndex        =   1
      Top             =   225
      Width           =   300
   End
End
Attribute VB_Name = "frmChooseColour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload frmChooseColour
    Form1.Show
    Form1.PicCol.BackColor = RGB(RGB_Col.R, RGB_Col.G, RGB_Col.B)
    
End Sub

Private Sub Form_Load()
Dim X, Y, Res As Long

    X = 160
    Y = 160
    Res = CreateEllipticRgn(6, 6, X, Y)
    SetWindowRgn Picture1.hWnd, Res, True
    RGB_Col.R = Val(RgbVal(0))
    RGB_Col.G = Val(RgbVal(1))
    RGB_Col.B = Val(RgbVal(2))
    DisplayCol.BackColor = RGB(RGB_Col.R, RGB_Col.G, RGB_Col.B)
    
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 1 Then
        BannerMod.Col_Move_Now = True
  End If
  
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Col As Long
On Error Resume Next
    
    If Not Col_Move_Now Then
        Exit Sub
    Else
        Shape1.Top = Y - 5
        Shape1.Left = X
        If Y > 160 Then
            Y = 87
            Shape1.Top = 87
        ElseIf Y < 10 Then
            Shape1.Left = 9
            Y = 9
        ElseIf X > 160 Then
            X = 150
            Shape1.Left = 150
        ElseIf X < 10 Then
            X = 9
            Shape1.Left = 9
        End If
    End If
    
    BannerMod.Long_Colour = Picture1.Point(Shape1.Left, Shape1.Top)
    BannerMod.ConvertLongToRGB
    RgbVal(0) = RGB_Col.R ' Red Value
    RgbVal(1) = RGB_Col.G ' Green Value
    RgbVal(2) = RGB_Col.B ' Blue Value
    DisplayCol.BackColor = RGB(RGB_Col.R, RGB_Col.G, RGB_Col.B)
    
    If Err Then Err.Clear
    
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    BannerMod.Col_Move_Now = False
    
End Sub

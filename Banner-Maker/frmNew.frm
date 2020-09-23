VERSION 5.00
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Banner"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "E&xit"
      Height          =   345
      Left            =   1680
      TabIndex        =   5
      Top             =   1980
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "360 X 60 Full Banner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   4
      Top             =   1200
      Width           =   2130
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   2910
      TabIndex        =   3
      Top             =   1980
      Width           =   1155
   End
   Begin VB.OptionButton opt392 
      Caption         =   "392 X 72 Full Banner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   990
      Width           =   2130
   End
   Begin VB.OptionButton opt486 
      Caption         =   "486 X 60 Standard Size"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   765
      Width           =   2475
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   90
      X2              =   2385
      Y1              =   645
      Y2              =   645
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000009&
      X1              =   90
      X2              =   2385
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Choose Your Banner Size"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   105
      TabIndex        =   0
      Top             =   300
      Width           =   2295
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Form1.PicBanner.Height = Ban_Size.Ban_Height
    Form1.PicBanner.Width = Ban_Size.Ban_Width
    Unload frmNew
    Form1.Show
    Form1.CenterBanner Form1.PicBanner
    
End Sub

Private Sub opt234_Click()
    Ban_Size.Ban_Height = 60
    Ban_Size.Ban_Width = 234
    
End Sub

Private Sub Command2_Click()
    Unload frmNew: End
    
End Sub

Private Sub Form_Load()
    CenterForm frmNew
    
End Sub

Private Sub opt392_Click()
    Ban_Size.Ban_Height = 72
    Ban_Size.Ban_Width = 392
    
End Sub

Private Sub opt486_Click()
    Ban_Size.Ban_Height = 60
    Ban_Size.Ban_Width = 486
    
    
    
End Sub

Private Sub Option1_Click()
    Ban_Size.Ban_Height = 60
    Ban_Size.Ban_Width = 360
    
End Sub

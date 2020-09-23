VERSION 5.00
Begin VB.Form FrmTexture 
   Caption         =   "Texture Broswer"
   ClientHeight    =   3705
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   ScaleHeight     =   247
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   165
      TabIndex        =   27
      Top             =   3195
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3060
      Left            =   165
      ScaleHeight     =   3000
      ScaleWidth      =   4305
      TabIndex        =   0
      Top             =   45
      Width           =   4365
      Begin VB.VScrollBar VScroll1 
         Height          =   3030
         Left            =   3990
         Max             =   3000
         TabIndex        =   26
         Top             =   -15
         Width           =   315
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6090
         Left            =   30
         ScaleHeight     =   406
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   272
         TabIndex        =   1
         Top             =   15
         Width           =   4080
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   1
            Left            =   0
            Picture         =   "FrmTexture.frx":0000
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   25
            Top             =   0
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   2
            Left            =   990
            Picture         =   "FrmTexture.frx":3042
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   24
            Top             =   0
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   3
            Left            =   1980
            Picture         =   "FrmTexture.frx":6084
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   23
            Top             =   0
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   4
            Left            =   2970
            Picture         =   "FrmTexture.frx":90C6
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   22
            Top             =   0
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   5
            Left            =   0
            Picture         =   "FrmTexture.frx":C108
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   21
            Top             =   1005
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   6
            Left            =   990
            Picture         =   "FrmTexture.frx":F14A
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   20
            Top             =   1005
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   7
            Left            =   1980
            Picture         =   "FrmTexture.frx":1218C
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   19
            Top             =   1005
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            BackColor       =   &H8000000B&
            Height          =   960
            Index           =   8
            Left            =   2970
            Picture         =   "FrmTexture.frx":151CE
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   18
            Top             =   1005
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   9
            Left            =   0
            Picture         =   "FrmTexture.frx":18210
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   17
            Top             =   2025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   10
            Left            =   990
            Picture         =   "FrmTexture.frx":1B252
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   16
            Top             =   2025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   11
            Left            =   1980
            Picture         =   "FrmTexture.frx":1E294
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   15
            Top             =   2025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   12
            Left            =   2970
            Picture         =   "FrmTexture.frx":212D6
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   14
            Top             =   2025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   13
            Left            =   0
            Picture         =   "FrmTexture.frx":24318
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   13
            Top             =   3030
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   14
            Left            =   990
            Picture         =   "FrmTexture.frx":2735A
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   12
            Top             =   3030
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   15
            Left            =   1980
            Picture         =   "FrmTexture.frx":2A39C
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   11
            Top             =   3030
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   16
            Left            =   2970
            Picture         =   "FrmTexture.frx":2D31E
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   10
            Top             =   3030
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   17
            Left            =   0
            Picture         =   "FrmTexture.frx":30360
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   9
            Top             =   4035
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   18
            Left            =   990
            Picture         =   "FrmTexture.frx":332E2
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   8
            Top             =   4035
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   19
            Left            =   1980
            Picture         =   "FrmTexture.frx":36324
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   7
            Top             =   4035
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   20
            Left            =   2970
            Picture         =   "FrmTexture.frx":39366
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   6
            Top             =   4035
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   21
            Left            =   0
            Picture         =   "FrmTexture.frx":3C3A8
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   5
            Top             =   5025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   22
            Left            =   990
            Picture         =   "FrmTexture.frx":3F3EA
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   4
            Top             =   5025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   23
            Left            =   1980
            Picture         =   "FrmTexture.frx":4236C
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   3
            Top             =   5025
            Width           =   960
         End
         Begin VB.PictureBox dest 
            AutoRedraw      =   -1  'True
            Height          =   960
            Index           =   24
            Left            =   2970
            Picture         =   "FrmTexture.frx":453AE
            ScaleHeight     =   60
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   60
            TabIndex        =   2
            Top             =   5010
            Width           =   960
         End
      End
   End
End
Attribute VB_Name = "FrmTexture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload FrmTexture
    Form1.picTx.Picture = Form1.Picture1.Picture
    Form1.Show
    
End Sub

Private Sub dest_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    dest(Index).DrawStyle = 2 ' Just added for  an effect
    dest(Index).Line (dest(Index).Width - 1, dest(Index).Height - 1)-(0, 0), vBlueed, B
    dest(Index).Refresh ' Refresh here
    Form1.Picture1.Picture = dest(Index).Picture
    
End Sub

Private Sub dest_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
     dest(Index).DrawStyle = 0
     dest(Index).Line (dest(Index).Width - 2, dest(Index).Height - 2)-(0, 0), &H8000000C, B  ' Just for an effect around the pictures
     dest(Index).Line (dest(Index).Width, dest(Index).Height - 1)-(0, 0), vbWhite, B '  Just for an effect around the pictures
     dest(Index).Refresh ' Refresh here
     
End Sub

Private Sub Form_Load()
Dim i As Integer
    CenterForm FrmTexture
    For i = 1 To 24
        dest(i).BorderStyle = 0 ' Just set the border style of all 24 pictures to 0
        dest(i).Refresh ' Refresh here
        dest(i).Line (dest(i).Width - 1, dest(i).Height - 1)-(0, 0), &H8000000C, B ' Just for an effect around the pictures
        dest(i).Line (dest(i).Width, dest(i).Height - 1)-(0, 0), vbWhite, B ' Just for an effect around the pictures
    Next
    i = 0
    
End Sub

Private Sub VScroll1_Change()
    Picture3.Top = -VScroll1.Value ' Scrolls picture box up or down
    
    
End Sub

Private Sub VScroll1_Scroll()
    VScroll1_Change
    
End Sub

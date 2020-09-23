VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Back Styles"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   4065
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Height          =   480
      Left            =   3135
      TabIndex        =   10
      Top             =   3090
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sytles"
      Height          =   3675
      Left            =   75
      TabIndex        =   0
      Top             =   150
      Width           =   4470
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   6
         Left            =   3315
         Picture         =   "Form2.frx":0000
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   9
         Top             =   285
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   7
         Left            =   3330
         Picture         =   "Form2.frx":3042
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   8
         Top             =   1365
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   8
         Left            =   120
         Picture         =   "Form2.frx":6084
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   7
         Top             =   2475
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   5
         Left            =   2250
         Picture         =   "Form2.frx":90C6
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   6
         Top             =   1380
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   4
         Left            =   1185
         Picture         =   "Form2.frx":C108
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   5
         Top             =   1380
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   3
         Left            =   105
         Picture         =   "Form2.frx":F14A
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   4
         Top             =   1365
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   2
         Left            =   2235
         Picture         =   "Form2.frx":1218C
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   3
         Top             =   285
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   1
         Left            =   1140
         Picture         =   "Form2.frx":151CE
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   2
         Top             =   285
         Width           =   1020
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         Height          =   1020
         Index           =   0
         Left            =   135
         Picture         =   "Form2.frx":18210
         ScaleHeight     =   960
         ScaleWidth      =   960
         TabIndex        =   1
         Top             =   285
         Width           =   1020
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 Unload Form2
 Form1.Show
 
End Sub

Private Sub Form_Load()
Form1.CenterForm Form2

End Sub

Private Sub Picture1_Click(Index As Integer)
Form1.Picture1.Picture = Picture1(Index).Picture

End Sub

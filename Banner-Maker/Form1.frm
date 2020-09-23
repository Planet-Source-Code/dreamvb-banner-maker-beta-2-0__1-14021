VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Banner Maker Beta 2.0"
   ClientHeight    =   6555
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   437
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   693
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   4095
      TabIndex        =   47
      Top             =   4035
      Width           =   1155
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   330
      Left            =   1245
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   330
      ScaleWidth      =   1380
      TabIndex        =   40
      Top             =   4050
      Width           =   1380
      Begin VB.Label Label7 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Height          =   300
         Index           =   3
         Left            =   1035
         TabIndex        =   45
         Top             =   15
         Width           =   315
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000007&
         BackStyle       =   0  'Transparent
         Height          =   270
         Index           =   2
         Left            =   645
         TabIndex        =   44
         Top             =   45
         Width           =   270
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   43
         Top             =   45
         Width           =   240
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000008&
         BackStyle       =   0  'Transparent
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   42
         Top             =   45
         Width           =   270
      End
      Begin VB.Shape shp 
         BorderColor     =   &H8000000B&
         Height          =   315
         Index           =   3
         Left            =   1005
         Top             =   15
         Width           =   375
      End
      Begin VB.Shape shp 
         BorderColor     =   &H8000000A&
         Height          =   315
         Index           =   2
         Left            =   660
         Top             =   15
         Width           =   315
      End
      Begin VB.Shape shp 
         BorderColor     =   &H8000000A&
         Height          =   315
         Index           =   1
         Left            =   330
         Top             =   15
         Width           =   330
      End
      Begin VB.Shape shp 
         BorderColor     =   &H8000000A&
         Height          =   315
         Index           =   0
         Left            =   0
         Top             =   15
         Width           =   315
      End
   End
   Begin VB.PictureBox picTx 
      Height          =   300
      Left            =   3345
      ScaleHeight     =   240
      ScaleWidth      =   315
      TabIndex        =   36
      Top             =   2955
      Width           =   375
   End
   Begin VB.CommandButton Command5 
      Height          =   405
      Left            =   3315
      TabIndex        =   35
      Top             =   2910
      Width           =   1035
   End
   Begin VB.PictureBox PicCol 
      BackColor       =   &H00000000&
      Height          =   300
      Left            =   1035
      ScaleHeight     =   240
      ScaleWidth      =   300
      TabIndex        =   32
      Top             =   2955
      Width           =   360
   End
   Begin VB.CommandButton Command4 
      Height          =   405
      Left            =   990
      TabIndex        =   31
      Top             =   2910
      Width           =   1035
   End
   Begin VB.PictureBox Cover 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC78C&
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   60
      Picture         =   "Form1.frx":17FA
      ScaleHeight     =   164
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   515
      TabIndex        =   27
      Top             =   120
      Width           =   7725
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFC78C&
         BorderStyle     =   0  'None
         Height          =   600
         Left            =   6405
         ScaleHeight     =   600
         ScaleWidth      =   1155
         TabIndex        =   29
         Top             =   1815
         Width           =   1155
         Begin VB.Image Image1 
            Height          =   480
            Left            =   330
            Picture         =   "Form1.frx":5C88
            Top             =   105
            Width           =   480
         End
      End
      Begin VB.PictureBox PicBanner 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         DragIcon        =   "Form1.frx":6552
         DragMode        =   1  'Automatic
         ForeColor       =   &H0000FFFF&
         Height          =   1050
         Left            =   255
         ScaleHeight     =   70
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   476
         TabIndex        =   28
         Top             =   555
         Width           =   7140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Beval"
      Height          =   1095
      Left            =   7830
      TabIndex        =   11
      Top             =   3195
      Width           =   2490
      Begin VB.OptionButton Option6 
         Caption         =   "No Bevel"
         Height          =   195
         Left            =   90
         TabIndex        =   26
         Top             =   720
         Width           =   1005
      End
      Begin VB.OptionButton optBevelLowerd 
         Caption         =   "Bevel Lowerd"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   495
         Width           =   1950
      End
      Begin VB.OptionButton OptBevUp 
         Caption         =   "Bevel Raised"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   255
         Width           =   1950
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Font Allignment"
      Height          =   1770
      Left            =   7830
      TabIndex        =   10
      Top             =   1335
      Width           =   2490
      Begin VB.HScrollBar HSY 
         Height          =   255
         Left            =   1200
         Max             =   30
         TabIndex        =   25
         Top             =   2355
         Width           =   1170
      End
      Begin VB.HScrollBar HSX 
         Height          =   255
         Left            =   1200
         Max             =   204
         TabIndex        =   24
         Top             =   2010
         Width           =   1170
      End
      Begin VB.TextBox txtYPos 
         Height          =   285
         Left            =   690
         MaxLength       =   3
         TabIndex        =   23
         Top             =   2325
         Width           =   435
      End
      Begin VB.TextBox txtXPos 
         Height          =   285
         Left            =   690
         MaxLength       =   3
         TabIndex        =   22
         Top             =   1995
         Width           =   435
      End
      Begin VB.CheckBox chkCustom 
         Caption         =   "Use Custom Values"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1515
         Width           =   2310
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Allign Center"
         Height          =   225
         Left            =   120
         TabIndex        =   16
         Top             =   1230
         Width           =   1980
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Allign Right"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1980
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Allign Left"
         Height          =   225
         Left            =   120
         TabIndex        =   14
         Top             =   750
         Width           =   1980
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Allign Bottom"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   525
         Width           =   1980
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Allign Top"
         Height          =   225
         Left            =   120
         TabIndex        =   12
         Top             =   270
         Width           =   1980
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Y Pos"
         Height          =   195
         Index           =   1
         Left            =   105
         TabIndex        =   21
         Top             =   2355
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "X Pos"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   420
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Properties"
      Height          =   1200
      Left            =   7830
      TabIndex        =   6
      Top             =   30
      Width           =   2430
      Begin VB.CheckBox chkStrike 
         Caption         =   "Font Strikeout"
         Height          =   195
         Left            =   150
         TabIndex        =   30
         Top             =   930
         Width           =   1875
      End
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Font &Underline"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   690
         Width           =   1875
      End
      Begin VB.CheckBox chkItalic 
         Caption         =   "Font &Italic"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   465
         Width           =   1470
      End
      Begin VB.CheckBox chkBold 
         Caption         =   "Font &Bold"
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   225
         Width           =   1050
      End
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   1035
      TabIndex        =   5
      Text            =   "Banner Maker 2.0"
      Top             =   3525
      Width           =   2760
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   4635
      TabIndex        =   4
      Top             =   3525
      Width           =   540
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   960
      Left            =   9120
      Picture         =   "Form1.frx":6994
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   3
      Top             =   4440
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   6090
      TabIndex        =   2
      Top             =   3510
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Save"
      Height          =   435
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Preview"
      Height          =   435
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4860
      Width           =   1395
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Drawing Colours"
      Height          =   195
      Left            =   2895
      TabIndex        =   46
      Top             =   4095
      Width           =   1155
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   5
      X1              =   0
      X2              =   82
      Y1              =   1
      Y2              =   1
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   4
      X1              =   0
      X2              =   82
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   3
      X1              =   -3
      X2              =   79
      Y1              =   307
      Y2              =   307
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   2
      X1              =   -3
      X2              =   79
      Y1              =   308
      Y2              =   308
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00E0E0E0&
      X1              =   82
      X2              =   82
      Y1              =   269
      Y2              =   296
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   83
      X2              =   185
      Y1              =   269
      Y2              =   269
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H8000000C&
      Height          =   405
      Left            =   1230
      Top             =   4035
      Width           =   1545
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Drawing Tools"
      Height          =   195
      Index           =   0
      Left            =   105
      TabIndex        =   41
      Top             =   4110
      Width           =   1020
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Font Name"
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   39
      Top             =   3555
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Font Size"
      Height          =   195
      Index           =   0
      Left            =   3870
      TabIndex        =   38
      Top             =   3555
      Width           =   660
   End
   Begin VB.Label Label5 
      Caption         =   "BannerTexture"
      Height          =   195
      Left            =   2190
      TabIndex        =   37
      Top             =   2985
      Width           =   1050
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Banner Text"
      Height          =   195
      Left            =   75
      TabIndex        =   34
      Top             =   3555
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Text Colour"
      Height          =   195
      Left            =   75
      TabIndex        =   33
      Top             =   2985
      Width           =   810
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   1
      X2              =   83
      Y1              =   186
      Y2              =   186
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   1
      X2              =   83
      Y1              =   185
      Y2              =   185
   End
   Begin VB.Menu mnuF 
      Caption         =   "&File"
      Begin VB.Menu mnuNewBanner 
         Caption         =   "&New Banner"
      End
      Begin VB.Menu mnuBitMap 
         Caption         =   "&Save To Bitmap"
      End
      Begin VB.Menu mnuJpeg 
         Caption         =   "&Save As Jpeg"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuAddText 
         Caption         =   "&Add Text"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Dim BannerText_Allignment As String
Dim DrawingTools As String

Dim BevelOnOff As Boolean
Dim LastChange As Boolean
Dim IsDone As Boolean
Dim Xpos, Ypos As Long


Dim Answer


Sub BevelOp(BevelType As Integer)
    Select Case BevelType
        Case 0 ' Lowerd
            PicBanner.Line (-1, PicBanner.Height)-(PicBanner.Width, 0), &H808080, B
            PicBanner.Line (0, PicBanner.Height)-(PicBanner.Width, -1), vbWhite, B
        Case 1 ' Rasied
            PicBanner.Line (-1, PicBanner.Height)-(PicBanner.Width, 0), vbWhite, B
            PicBanner.Line (0, PicBanner.Height)-(PicBanner.Width, -1), &H808080, B
    End Select
    
End Sub
Sub CenterBanner(BannerWindow As PictureBox)
    With BannerWindow
        .Top = (Cover.Height - Ban_Size.Ban_Height) / 2
        .Left = (Cover.Width - Ban_Size.Ban_Width) / 2
    End With

End Sub

Sub MakeBanner(Xpos, Ypos As Integer, Title, TitleFont As String, TextSize As Integer, bsBold, bsItalic, bsUnLine, bsStrike, TextColour As Long)
Dim X, Y, i, J As Single

    i = Picture1.ScaleWidth
    J = Picture1.ScaleHeight
    
    Y = 0
    Do While Y < PicBanner.ScaleHeight
        X = 0
        Do While X < PicBanner.ScaleWidth
            PicBanner.PaintPicture Picture1.Picture, X, Y, i, J
            PicBanner.FontName = TitleFont
            PicBanner.FontBold = bsBold
            PicBanner.FontItalic = bsItalic
            PicBanner.FontUnderline = bsUnLine
            PicBanner.FontStrikethru = bsStrike
            
            PicBanner.ForeColor = TextColour
            PicBanner.FontSize = TextSize
            
            Select Case BannerText_Allignment
                Case "Top"
                    TextOut PicBanner.hdc, PicBanner.Width / 3.5, 5 - 10, Title, Len(Title) ' top
                Case "Bottom"
                    TextOut PicBanner.hdc, PicBanner.Width / 3.5, 5 * 5, Title, Len(Title) ' Bottom
                Case "Center"
                    TextOut PicBanner.hdc, PicBanner.Width / 3.5, 5 * 2, Title, Len(Title) ' Center
                Case "Right"
                    TextOut PicBanner.hdc, PicBanner.Width - 12 * 5 * 3, 5 * 2, Title, Len(Title) ' Right
                Case "Left"
                    TextOut PicBanner.hdc, 10 - 10, 5 * 2, Title, Len(Title) ' Left
                Case "Custom"
                    TextOut PicBanner.hdc, Xpos, Ypos, Title, Len(Title)
            End Select
            X = X + i
        Loop
        Y = Y + J
    Loop
    Y = 0
    J = 0
    X = 0
    LastChange = True
    
End Sub


Private Sub Check1_Click()

    
End Sub

Private Sub chkCustom_Click()
    If chkCustom Then
        BannerText_Allignment = "Custom"
        Frame2.Height = 207
        Frame3.Top = 327
        Option1.Enabled = False
        Option2.Enabled = False
        Option3.Enabled = False
        Option4.Enabled = False
        Option5.Enabled = False
    Else
        BannerText_Allignment = "Center"
        Option1.Enabled = True
        Option2.Enabled = True
        Option3.Enabled = True
        Option4.Enabled = True
        Option5.Enabled = True
        Frame2.Height = 120
        Frame3.Top = 238
    End If
    
End Sub

Private Sub Combo3_Click()
    Select Case Combo1.ListIndex
        Case 0
            Drawing_Colour = vbRed
        Case 1
            Drawing_Colour = vbBlue
        Case 2
            Drawing_Colour = vbGreen
        Case 4
            Drawing_Colour = vbBlack
        Case 5
            Drawing_Colour = vbWhite
        End Select
        
End Sub

Private Sub Command1_Click()
    If Not IsDone Then
        MsgBox "You have not Selected a texture for your banner please selet one now and try agian", vbInformation
    Else
    
        MakeBanner Val(txtXPos.Text), Val(txtYPos.Text), Text4, Combo1.Text, Val(Combo2), _
        chkBold, chkItalic, chkUnderline, chkStrike, PicCol.BackColor
        Command2.Enabled = True
        If BevelOnOff Then
            OptBevUp_Click
        End If
    End If
    
End Sub

Private Sub Command2_Click()
Dim BanSave As String
    
    BanSave = SaveBanner
    If Len(BanSave) = 0 Then Exit Sub
    
    If InStr(BanSave, ".") Then
        MsgBox BanSave
        Exit Sub
    Else
        BanSave = Left(BanSave, Len(BanSave) - 1) ' Removes Null string
        BanSave = BanSave & ".bmp"
    End If
        SavePicture PicBanner.Image, BanSave
        MsgBox "You Banner have now been saved", vbInformation
 
End Sub

Private Sub Command4_Click()
    frmChooseColour.Show
    Form1.Hide
    
End Sub

Private Sub Command5_Click()
    Form1.Hide
    FrmTexture.Show
    IsDone = True
    
End Sub

Private Sub Form_Load()
    CenterForm Form1
    CenterBanner PicBanner
    
    For i = 1 To Screen.FontCount - 1
        Combo1.AddItem Screen.Fonts(i)
    Next
    i = 0
    
        Combo2.AddItem 9
        Combo2.AddItem 10
        Combo2.AddItem 12
        Combo2.AddItem 14
        Combo2.AddItem 16
        Combo2.AddItem 18
        Combo2.AddItem 20
        Combo2.AddItem 24
        Combo2.AddItem 30
        
        Combo1.ListIndex = 4
        Combo2.ListIndex = 5
        
        Combo3.AddItem "Red"
        Combo3.AddItem "Blue"
        Combo3.AddItem "Green"
        Combo3.AddItem "Black"
        Combo3.AddItem "White"
        
        Combo3.ListIndex = 3
        
        PicCol.BackColor = vbBlack
        BannerText_Allignment = "Center"
        If Ban_Size.Ban_Width = 486 Then HSX.Max = 300: HSY.Max = 25
        If Ban_Size.Ban_Width = 392 Then HSX.Max = 210: HSY.Max = 38
        If Ban_Size.Ban_Width = 360 Then HSX.Max = 178: HSY.Max = 26
    
        HSX.Value = 100
        HSY.Value = 5
        LastChange = False
        IsDone = False
        
End Sub

Private Sub Form_Resize()
    Line1(0).X2 = Me.ScaleWidth - Frame2.Width
    Line1(1).X2 = Me.ScaleWidth - Frame2.Width
    '
    Line1(2).X2 = Me.ScaleWidth
    Line1(3).X2 = Me.ScaleWidth
    '
    Line1(4).X2 = Me.ScaleWidth
    Line1(5).X2 = Me.ScaleWidth
    
End Sub

Private Sub HScroll1_Change()
    If HScroll1.Value < 1 Then
        HScroll1.Value = 1
    End If
    txtThick.Text = HScroll1.Value
           
End Sub

Private Sub HScroll1_Scroll()
    HScroll1_Change
    
End Sub

Private Sub HSX_Change()
    txtXPos.Text = HSX.Value
    
End Sub

Private Sub HSX_Scroll()
    HSX_Change
    
End Sub

Private Sub HSY_Change()
    txtYPos.Text = HSY.Value
    
End Sub

Private Sub HSY_Scroll()
    HSY_Change
    
End Sub

Private Sub Image1_DragDrop(Source As Control, X As Single, Y As Single)
    If LastChange = False Then
        Exit Sub
    Else
        Answer = _
            MsgBox("Are you sure you want to delete your banner", _
            vbYesNo, "Delete....")
            
            If Answer = vbNo Then
                Exit Sub
            Else
                PicBanner.Picture = Nothing
                Command2.Enabled = False
                LastChange = False
                IsDone = False
                
                Beep
            End If
        End If
        
End Sub

Private Sub Label7_Click(Index As Integer)
    Select Case Index
        Case 0
            DrawingTools = "Box"
        Case 1
            DrawingTools = "Circle"
        Case 2
            DrawingTools = "Tri1"
        Case 3
            DrawingTools = "Tri2"
        End Select
        
End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shp(Index).BorderColor = vbRed
    
End Sub

Private Sub Label7_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    shp(Index).BorderColor = &HC0C0C0
    
End Sub

Private Sub mnuAbout_Click()
        MsgBox "Banner Maker Beta 2.0" & vbNewLine & "  By Ben Jones", vbInformation
        
End Sub

Private Sub mnuBitMap_Click()
    Command2_Click
    
End Sub

Private Sub mnuExit_Click()
    Answer = _
        MsgBox("Are you sure that you want to quit now", _
        vbYesNo, "Exit......")
    If Answer = vbNo Then
        Exit Sub
    Else
        End
    End If
    
End Sub

Private Sub mnuJpeg_Click()
    MsgBox "Comming Soon", vbInformation
    
End Sub

Private Sub mnuNewBanner_Click()
Dim Answer

    If Not LastChange Then
        Form1.Hide
        frmNew.Show
    Else
        Answer = _
        MsgBox("Do you want to save your changes", _
        vbYesNo)
        If Answer = vbNo Then
            Unload Form1
            frmNew.Show
        Else
        ' Save picture here
        Command2_Click
        End If
    End If

    
End Sub

Private Sub optBevelLowerd_Click()
    BevelOp 0
    BevelOnOff = True
    
End Sub

Private Sub OptBevUp_Click()
    BevelOp 1
    BevelOnOff = True
    
End Sub

Private Sub Option1_Click()
    BannerText_Allignment = "Top"
    
End Sub

Private Sub Option2_Click()
    BannerText_Allignment = "Bottom"
    
End Sub

Private Sub Option3_Click()
    BannerText_Allignment = "Left"
    
End Sub

Private Sub Option4_Click()
    BannerText_Allignment = "Right"
    
End Sub

Private Sub Option5_Click()
    BannerText_Allignment = "Center"
End Sub


Private Sub Picture5_Click()
    Picture1.Picture = Picture5.Picture
    
End Sub

Private Sub Option6_Click()
    BevelOnOff = False
    Command1_Click
    
End Sub


Private Sub PicBanner_Click()
    If DrawingTools = "Circle" Then
        PicBanner.Circle (Xpos, Ypos), 5
 ElseIf DrawingTools = "Box" Then
            PicBanner.Line (Xpos, Ypos)-(Xpos + 5, Ypos + 5), vbRed, BF
        ElseIf DrawingTools = "Tri1" Then
            PicBanner.Circle (Xpos, Ypos), 10, , , -1, -1
        ElseIf DrawingTools = "Tri2" Then
            PicBanner.Circle (Xpos, Ypos), 10, , , -1 * 5
        End If
        
        
End Sub

Private Sub PicBanner_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Xpos = X
    Ypos = Y
    CurrentX = X
    CurrentY = Y
    
End Sub


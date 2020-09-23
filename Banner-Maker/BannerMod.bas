Attribute VB_Name = "BannerMod"
Type RGB_Col
    R As Integer
    G As Integer
    B As Integer
End Type

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal Blueedraw As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    FLAGS As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type



Type BannerSize
    Ban_Height As Integer
    Ban_Width As Integer
End Type

Public Long_Colour As Long
Public Col_Move_Now As Boolean
Public RGB_Col As RGB_Col
Public Ban_Size As BannerSize

Public Function SaveBanner() As String
 Dim ofn As OPENFILENAME
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Form1.hWnd
    ofn.hInstance = App.hInstance
    ofn.lpstrFilter = ""
        ofn.lpstrFile = Space$(254)
        ofn.nMaxFile = 255
        ofn.lpstrFileTitle = Space$(254)
        ofn.nMaxFileTitle = 255
        ofn.lpstrInitialDir = App.Path & "\"
        ofn.lpstrTitle = "Save Banner"
        ofn.FLAGS = 0
        
        A = GetSaveFileName(ofn)
        If (A) Then
                SaveBanner = Trim$(ofn.lpstrFile)
        End If
        
 End Function
Public Sub ConvertLongToRGB()
Dim Red, Green, Blue As Long
Dim Extra As Long

    Red = 1
    Green = 256
    Blue = 65536

    Extra = Long_Colour \ Blue
    RGB_Col.B = Extra
    Long_Colour = Long_Colour Mod Blue
    If RGB_Col.B < 0 Then RGB_Col.B = 0
    
    Extra = Long_Colour \ Green
    RGB_Col.G = Extra
    Long_Colour = Long_Colour Mod Green
    If RGB_Col.G < 0 Then RGB_Col.G = 0
    
    
    Extra = Long_Colour \ Red
    RGB_Col.R = Extra
    Long_Colour = Long_Colour Mod Red
    If RGB_Col.R < 0 Then RGB_Col.R = 0

End Sub

Public Function CenterForm(Frm As Form)
    With Frm
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
    End With

End Function

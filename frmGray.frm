VERSION 5.00
Begin VB.Form frmGray 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmGray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_EX_RIGHT As Long = &H1000
Private Const WS_EX_LEFTSCROLLBAR As Long = &H4000
Private Const WS_EX_LAYERED As Long = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_COLORKEY As Long = &H1
Private Const LWA_ALPHA As Long = &H2

Private Declare Function GetWindowLong Lib "User32" _
  Alias "GetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "User32" _
  Alias "SetWindowLongA" _
  (ByVal hwnd As Long, _
  ByVal nIndex As Long, _
  ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "User32" _
  (ByVal hwnd As Long, _
  ByVal crKey As Long, _
  ByVal bAlpha As Long, _
  ByVal dwFlags As Long) As Long

Dim formOpening As Boolean

Private Sub Form_Activate()

  If formOpening Then
    Me.WindowState = vbMaximized
    frmModalForm.Show vbModal
    Unload Me
    formOpening = False
  End If

End Sub

Private Sub Form_Load()

  formOpening = True
  Me.BackColor = vbBlack
  Call AlphaForm

End Sub

Private Function AlphaForm()

  Dim style As Long
  Dim alpha As Integer

  alpha = 192
  style = GetWindowLong(Me.hwnd, GWL_EXSTYLE)

  If Not (style And WS_EX_LAYERED = WS_EX_LAYERED) Then
    style = style Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, style
    SetLayeredWindowAttributes Me.hwnd, 0&, alpha, LWA_ALPHA
  End If

End Function


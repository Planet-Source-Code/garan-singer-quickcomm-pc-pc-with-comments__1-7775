VERSION 5.00
Begin VB.Form frmNotify 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   450
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   1350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmNotify.frx":0000
   ScaleHeight     =   30
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'The API stuff here is all for keeping the form on top
'of other forms.

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" _
            (ByVal hwnd As Long, _
            ByVal hWndInsertAfter As Long, _
            ByVal X As Long, _
            ByVal Y As Long, _
            ByVal cx As Long, _
            ByVal cy As Long, _
            ByVal wFlags As Long) As Long

Private Sub Form_Load()
    'Sets the notify form on top of all others.
    SetTopMost Me.hwnd
    'Ensures that it is hidden.
    Me.Hide
    'Sets the leftmost and topmost positions of the notify
    'form correct to keep it in the bottom-left position
    'of the user's screen.
    Me.Left = Screen.Width - Me.Width
    Me.Top = Screen.Height - Me.Height
End Sub

Private Sub SetTopMost(hwnd As Long)
    'Simply a simple way to call the OnTop code. Simply
    'pass the hwnd of the
    Call SetWindowPos(hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub


VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmConWindow 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2415
   ClientLeft      =   5100
   ClientTop       =   15
   ClientWidth     =   6870
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   161
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   458
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdChangeSide 
      Caption         =   "<<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtIP 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   3
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Timer timHideNotify 
      Interval        =   50
      Left            =   0
      Top             =   960
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5640
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin MSWinsockLib.Winsock sckSocket 
      Left            =   0
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemoteHost      =   "127.0.0.1"
      LocalPort       =   888
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5640
      TabIndex        =   1
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtToSend 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   5415
   End
   Begin VB.TextBox txtChatWindow 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   360
      Width           =   5415
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "System"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6600
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   285
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   0
      Picture         =   "frmConWindow.frx":0000
      ScaleHeight     =   255
      ScaleWidth      =   6870
      TabIndex        =   5
      Top             =   0
      Width           =   6870
   End
   Begin VB.Label Label2 
      Caption         =   "IP:"
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Name:"
      Height          =   255
      Left            =   5640
      TabIndex        =   7
      Top             =   360
      Width           =   1095
   End
   Begin VB.Line linBorderJazzy 
      X1              =   264
      X2              =   152
      Y1              =   17
      Y2              =   17
   End
   Begin NiftySoft_QuickComm.Form_TaskBar Form_TaskBar1 
      Left            =   0
      Top             =   1920
      _ExtentX        =   1931
      _ExtentY        =   450
      NumSteps        =   5
   End
End
Attribute VB_Name = "frmConWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*********************************************************
'*     QuickComm PC - PC UDP Communications Program      *
'*********************************************************
'*                                                       *
'*  This complete  application has been  released to the *
'* public via Planet-Source-Code.com  for public use. If *
'* anyone  wished to use  any portion of  the code found *
'* here, they may  do so with all  rights. Only the name *
'* and the graphics are not included in this.            *
'*  Also,  the usercontrol Form_TaskBar  must be granted *
'* permission    by    its    author,    David    Newcum *
'* (newcumdb@cs.purdue.edu).  Otherwise,  feel  free  to *
'* use the  code contained  herein  to  enhance or  make *
'* functional your own work.                             *
'*********************************************************

'This statement informs Visual Basic that you wish to
'dimension all variables before use. VERY GOOD IDEA to use
'this, as it helps deter using Variant data types, and
'helps keep the code clean.
Option Explicit

'API declaration used to play a wave file.
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Constants used with the PlaySound API.
Private Const SND_ASYNC = &H1
Private Const SND_NODEFAULT = &H2
Private Const SND_FLAG = SND_ASYNC Or SND_NODEFAULT

'Boolean flag determining which side of the screen the
'frmConWindow form is on.
Dim Side As Boolean

Private Sub cmdChangeSide_Click()
    'This subroutine occurs when the cmdChangeSide button
    'has been clicked. It is used to change the side of the
    'screen the frmConWindow form is placed.
    
    If Side = False Then
        'If Side is False, the window is on the right side.
        'Here we switch that and place it on the left side.
        'Just set the form's left-most edge to 0.
        frmConWindow.Left = 0
        'And change the caption to reflect the change.
        cmdChangeSide.Caption = ">>"
    Else
        'If Side is True, the window is on the left side.
        'Switch that around to place it on the right side.
        'Screen.Width is the screen's width in twips.
        'To find the position needed to place the left-most
        'edge of the form, we subtract the form's width
        'in twips from the screen's width in twips.
        frmConWindow.Left = Screen.Width - frmConWindow.Width
        'Change the caption to reflect the change.
        cmdChangeSide.Caption = "<<"
    End If
    
    'Toggle the Side boolean flag.
    'The logic here is that "Not True" = "False" and
    '"Not False" = "True". Makes sense, right? :)
    Side = Not Side
End Sub

Private Sub cmdClose_Click()
    'This subroutine occurs when the cmdClose button
    'has been clicked. It is used to close the program.
    
    'This ensures that if there are any errors the user
    'will not be notified. Erronious IP addresses can
    'simply be ignored here, since the user is exiting
    'the program.
    On Error Resume Next
    
    'These two commands save the program's settings to the
    'registry so they can be retrieved at the beginning
    'of the program.
    SaveSetting "QuickComm", "Settings", "Name", txtName.Text
    SaveSetting "QuickComm", "Settings", "LastIP", txtIP.Text
    
    'This sends the remote computer (the receiving computer)
    'information telling it that this computer is
    'disconnecting.
    'Set the Winsock control's RemoteHost property to the
    'IP address specified in txtIP.Text.
    sckSocket.RemoteHost = txtIP.Text
    'Send data through the socket.
    sckSocket.SendData txtName.Text + " has quit QuickComm..."
    
    'This is the best way possible to close a program.
    'It uses a For-Each-Next loop to cycle through each
    'object in a collection, namely each Form in the Forms
    'collection. For Each Form in the Forms collection,
    'it closes it.
    Dim loopf As Form
    
    For Each loopf In Forms
        Unload loopf
    Next
    
    End
End Sub

Private Sub cmdSend_Click()
    'This subroutine occurs when the cmdSend button is
    'clicked. It is used to send text to the receiving
    'computer.
    
    'Make sure that any errors occurring don't interrupt
    'the user.
    On Error Resume Next
    
    'Set the Winsock control's RemoteHost property to the
    'IP address specified in txtIP.Text.
    sckSocket.RemoteHost = txtIP.Text
    'Bind the socket to this port.
    sckSocket.Bind
    
    'Check for commands sent, if any.
    If txtToSend.Text = "/beep" Then
        'MUST be "/beep". Cannot be "/Beep" or etc.
        'This command plays the "Default Sound" wave file.
        Beep
        'This adds the text notifying the user that they
        '/beeped the receiving user, instead of just showing
        '/beep. More user friendly I think.
        txtChatWindow.Text = txtChatWindow + "You have /beeped the receiver." + vbCrLf
    Else
        'If there were no commands sent, simply add the
        'user's name, followed by a colon and a space,
        'then the user's text in the chat window.
        txtChatWindow.Text = txtChatWindow + txtName.Text + ": " + txtToSend.Text + vbCrLf
    End If
    
    'Send the data. All of it. :)
    sckSocket.SendData txtName.Text + ": " + txtToSend.Text
    'Set the text box containing the text to be sent to
    'empty.
    txtToSend.Text = ""
    'And give it focus so the user can start typing right
    'away.
    txtToSend.SetFocus
End Sub

Private Sub Form_Load()
    'This subroutine occurs every time the frmConWindow
    'form is loaded, which only occurs upon startup.
    
    'This sets the line below the title-bar graphic to the
    'width of the form. So I was lazy. Sue me. :)
    linBorderJazzy.x1 = 0
    linBorderJazzy.x2 = frmConWindow.Width
    
    'These two commands retrieve the respective name and
    'Last IP settings from the system registry.
    txtName.Text = GetSetting("QuickComm", "Settings", "Name")
    txtIP.Text = GetSetting("QuickComm", "Settings", "LastIP", "127.0.0.1")
    
    'This initializes the TaskBar simulation control by
    'David Newcum.
    Form_TaskBar1.InitTB
    
    
    'Make sure that any errors occurring don't interrupt
    'the user.
    On Error Resume Next
    
    'Set the remote port, the remote host (by txtIP.Text),
    'and bind the socket to the port.
    sckSocket.RemotePort = 888
    sckSocket.RemoteHost = txtIP.Text
    sckSocket.Bind
    
    'Inform the other user that we've come online. Don't
    'worry if they're online or not, or if the IP is
    'invalid or empty, because, if the IP isn't running
    'this program, nothing will happen because this is
    'UDP -- a connectionless protocol. And if the IP is
    'invalid, nothing will happen because of the
    '"On Error Resume Next" statement.
    sckSocket.SendData txtName.Text + " has entered QuickComm chat."
End Sub

Private Sub sckSocket_DataArrival(ByVal bytesTotal As Long)
    'This subroutine occurs when data arrives in the
    'sckSocket control's buffer.
    
    'We need some place to put the data.
    Dim incoming As String
    
    'Reserve the string SystemCommand for use with, well,
    'System Commands.
    Dim SystemCommand As String
    
    'Retrieve the data from the buffer and place it into
    'the string we reserved.
    sckSocket.GetData incoming
    
    'Check for system commands.
    
    'Really long and stupid string here. This simply parses
    'off the bytes after the ": ", which takes off the name
    'of the sender, and places it in SystemCommand.
    SystemCommand = Right(incoming, Len(incoming) - (InStr(1, incoming, ": ") + 1))
    
    If SystemCommand = "/beep" Then
        'If the last bytes are "/beep" then play the
        '"Default Sound" wave and display that the sending
        'user has /beeped us.
        Beep
        txtChatWindow.Text = txtChatWindow + Left(incoming, (InStr(1, incoming, ": ") - 1)) + " has /Beeped you." + vbCrLf
    Else
        'Else, just display the data they sent us.
        txtChatWindow.Text = txtChatWindow + incoming + vbCrLf
    End If
    
    'This checks to see if the form is sitting completely
    'out (i.e. Visible).
    If Form_TaskBar1.IsTaskbarOut = False Then
        'If not, play a notification sound...
        Playwav App.Path + "\gnid.wav"
        '... and show the frmNotify form.
        frmNotify.Show
    End If

End Sub

Private Sub timHideNotify_Timer()
    'This subroutine hides the Notify form if the
    'form is completely visible.
    
    If Form_TaskBar1.IsTaskbarOut = True Then
        frmNotify.Hide
    End If
End Sub

Private Sub txtChatWindow_Change()
    'This subroutine places the chat window's selection
    'area at the end of the text so that the text box will
    'scroll down when text enters it.
    
    txtChatWindow.SelStart = Len(txtChatWindow)
End Sub

Private Sub txtToSend_KeyPress(KeyAscii As Integer)
    'This checks keypresses in the txtToSend text box.
    'If the key pressed is character code 13, the Enter
    'key (actually the Carriage Return code), then set
    'the key pressed to 0 (null so it doesn't Ding), and
    'click the cmdSend button to execute the code in there.
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdSend_Click
    End If
End Sub

Public Sub Playwav(WavFile As String)
    'This subroutine plays the specified wave file.
    'Circumvents having to use the API all the time.
    
    'Reserve a place for the file existance verification.
    Dim SafeFile As String
    
    'The Dir command creates a listing of files using its
    'input as a pattern. If the file pattern listed is
    'matched, it returns all the files that match it. If
    'a file is specified, and it exists, it returns the
    'filename. If the file doesn't exist, it returns null.
    SafeFile$ = Dir(WavFile$)
    If SafeFile$ <> "" Then
        'So if the return isn't null, we know the file
        'actually exists. Thus, we can call the
        'sndPlaySound API function to, well, play the
        'sound. :)
        Call sndPlaySound(WavFile$, SND_FLAG)
    End If
End Sub

'That's all for this form.

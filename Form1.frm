VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'Kein
   Caption         =   "Working..."
   ClientHeight    =   2655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7095
   ForeColor       =   &H0000FF00&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   2655
   ScaleWidth      =   7095
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer Timer6 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1920
   End
   Begin VB.CommandButton Command6 
      Caption         =   "test"
      Height          =   255
      Left            =   6360
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.PictureBox picProgress6 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   12
      Top             =   1920
      Width           =   6135
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1560
   End
   Begin VB.PictureBox picProgress5 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   11
      Top             =   1560
      Width           =   6135
   End
   Begin VB.CommandButton Command5 
      Caption         =   "test"
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   1200
   End
   Begin VB.CommandButton Command4 
      Caption         =   "test"
      Height          =   255
      Left            =   6360
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.PictureBox picProgress4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   10
      Top             =   1200
      Width           =   6135
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   840
   End
   Begin VB.PictureBox picProgress3 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   840
      Width           =   6135
   End
   Begin VB.CommandButton Command3 
      Caption         =   "test"
      Height          =   255
      Left            =   6360
      TabIndex        =   3
      Top             =   840
      Width           =   615
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   120
      Top             =   480
   End
   Begin VB.CommandButton Command2 
      Caption         =   "test"
      Height          =   255
      Left            =   6360
      TabIndex        =   2
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "go"
      Height          =   255
      Left            =   6360
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.PictureBox picProgress2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   8
      Top             =   480
      Width           =   6135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox picProgress1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   6075
      TabIndex        =   0
      ToolTipText     =   "Nice effect, he? :)"
      Top             =   120
      Width           =   6135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FF00&
      Height          =   2655
      Left            =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abort"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   6240
      MousePointer    =   10  'Aufwärtspfeil
      TabIndex        =   7
      Top             =   2160
      Width           =   690
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessage& Lib "user32" Alias "SendMessageA" (ByVal hWnd&, ByVal wMsg&, ByVal wParam&, lParam As Any)
Private Declare Function ReleaseCapture Lib "user32" () As Long
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
'Declare all needed variables
Dim BitD As Boolean, StartA%, StartB%, StartC%, StartD%, StartE%, StartF%, abc%, Abort As Boolean
Private Sub Command1_Click()
Select Case Command1.Caption
Case "go"
  Abort = False
  Command1.Caption = "stop"
  Timer1.Enabled = True
Case "stop"
  Command1.Caption = "go"
  Timer1.Enabled = False
  picProgress1.Cls
  StartA = 0
End Select
End Sub
Private Sub Command2_Click()
Abort = False
If Timer2.Enabled = False Then Timer2.Enabled = True
End Sub
Private Sub Command3_Click()
Abort = False
If Timer3.Enabled = False Then Timer3.Enabled = True
End Sub
Private Sub Command4_Click()
Abort = False
If Timer4.Enabled = False Then Timer4.Enabled = True
End Sub
Private Sub Command5_Click()
Abort = False
If Timer5.Enabled = False Then
  StartE = 0
  Timer5.Enabled = True
End If
End Sub
Private Sub Command6_Click()
Abort = False
If Timer6.Enabled = False Then
  StartF = 0
  Timer6.Enabled = True
End If
End Sub
Private Sub Form_Load()
SendMessage Command1.hWnd, &HF4&, &H0&, 0&
SendMessage Command2.hWnd, &HF4&, &H0&, 0&
SendMessage Command3.hWnd, &HF4&, &H0&, 0&
SendMessage Command4.hWnd, &HF4&, &H0&, 0&
SendMessage Command5.hWnd, &HF4&, &H0&, 0&
SendMessage Command6.hWnd, &HF4&, &H0&, 0&
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
SetObjectMoveReplace Me.hWnd
End Sub
Private Sub Label5_Click()
If Abort = True Then End
Abort = True
End Sub
Private Sub Timer1_Timer()
'********************************************************'
'Progress with no end. With nice gradients and looks cool'
'Idea taken from Netscape Communicator 5 BETA Installer  '
'********************************************************'
'This is the never-ending progress. On some computers it
'produces some little line in the PictureBox (but looks kewl too :))
On Error Resume Next 'Errors suck :]
If Abort Then End 'If user abort operation exit the program
If Not BitD Then 'If painting-mode is left->right
  For a = 0 To 250 Step 2
    'Start painting the gradient on the left
    picProgress1.Line (StartA + a * 2, 0)-(StartA + a * 2 + 2, picProgress1.Height), RGB(0, a, 0), BF
  Next a
  'Paints the inner box of progress-mark
  picProgress1.Line (StartA + 500, 0)-(StartA + 1500, picProgress1.Height), RGB(0, 255, 0), BF
  For a = 0 To 250 Step 2
    'Start painting the gradient on the right
    picProgress1.Line ((StartA + 1500) + a * 2, 0)-((StartA + 1500) + a * 2 + 2, picProgress1.Height), RGB(0, 255 - a, 0), BF
    Next a
  'Increase marks position by 45
  StartA = StartA + 45
End If
'If painting-mode is right->left
If StartA + 2000 >= picProgress1.Width Or BitD = True Then
  BitD = True 'Must be set to reenter this sub
  For a = 0 To 250 Step 2 'The gradient again
    picProgress1.Line (StartA + a * 2, 0)-(StartA + a * 2 + 2, picProgress1.Height), RGB(0, a, 0), BF
  Next a 'And the block...
  picProgress1.Line (StartA + 500, 0)-(StartA + 1500, picProgress1.Height), RGB(0, 255, 0), BF
  For a = 0 To 250 Step 2 'Next!
    picProgress1.Line ((StartA + 1500) + a * 2, 0)-((StartA + 1500) + a * 2 + 2, picProgress1.Height), RGB(0, 255 - a, 0), BF
  Next a
  StartA = StartA - 45 'Decrease marks position by 45
  If StartA <= 0 Then BitD = False 'Set position back if way's finished
End If
'That's it. Nice one :)
End Sub
Private Sub SetObjectMoveReplace(ObjHWND&)
'This Sub is used to move the form around w/o titlebar
ReleaseCapture
SendMessage ObjHWND, WM_NCLBUTTONDOWN, HTCAPTION, 0
End Sub
Private Sub Timer2_Timer()
'****************************************'
'This is just the boring standard bar :P '
'It's just here to complete the thing.   '
'****************************************'
If Abort = True Then End 'End program if user aborts
picProgress2.Line (0, 0)-((picProgress2.Width / 100) * StartB, picProgress2.Height), vbGreen, BF
If StartB = 100 Then 'If Progress is complete then:
  StartB = 0         'Set progress to zero
  picProgress2.Cls   'Clear the PictureBox
  Timer2.Enabled = False 'Disable Timer
End If
StartB = StartB + 1 'Increase Progress by 1
End Sub
Private Sub Timer3_Timer()
'******************************************'
'This is just the standard bar again, but  '
'it changes color(s) when progress changes '
'This looks some times better than the     '
'boring standard one. The code's the same  '
'as above, but has other rgb-settings      '
'******************************************'
If Abort = True Then End
picProgress3.Line (0, 0)-((picProgress3.Width / 100) * StartC, picProgress3.Height), RGB(0, 2.5 * StartC, 0), BF
If StartC = 100 Then
  StartC = 0
  picProgress3.Cls
  Timer3.Enabled = False
End If
StartC = StartC + 1
End Sub
Private Sub Timer4_Timer()
'******************************************'
'This code creates a gradient over the     '
'complete PictureBox.                      '
'It's as simple as the others.             '
'******************************************'
If Abort = True Then End
picProgress4.Line ((picProgress4.Width / 100) * StartD, 0)-((picProgress4.Width / 100) * (StartD + 5), picProgress4.Height), RGB(0, 2.5 * StartD + 50, 0), BF
If StartD = 100 Then
  StartD = 0
  picProgress4.Cls
  Timer4.Enabled = False
End If
StartD = StartD + 1
End Sub
Private Sub Timer5_Timer()
If Abort = True Then End
picProgress5.Line ((picProgress5.Width / 100) * StartE, 0)-((picProgress5.Width / 100) * (StartE + 3), picProgress5.Height), RGB(0, 255, 0), BF
If StartE = 100 Then
  picProgress5.Cls
  Timer5.Enabled = False
End If
StartE = StartE + 4
End Sub
Private Sub Timer6_Timer()
If Abort = True Then End
picProgress6.Line ((picProgress6.Width / 100) * StartF, 0)-((picProgress6.Width / 100) * (StartF + 3), picProgress6.Height), RGB(255 - (2.5 * StartF), 2.5 * StartF, 0), BF
If StartF = 100 Then
  picProgress6.Cls
  Timer6.Enabled = False
End If
StartF = StartF + 4
End Sub

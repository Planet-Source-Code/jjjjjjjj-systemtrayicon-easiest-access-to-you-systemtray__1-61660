VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "System Tray"
   ClientHeight    =   1590
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   5880
   ClipControls    =   0   'False
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmTest.frx":030A
   ScaleHeight     =   1590
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
      Begin VB.Timer tmrUpdate 
         Enabled         =   0   'False
         Interval        =   1111
         Left            =   2520
         Top             =   240
      End
      Begin VB.PictureBox PicIcon 
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   3
         Left            =   1920
         Picture         =   "frmTest.frx":0614
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   5
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox PicIcon 
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   2
         Left            =   1320
         Picture         =   "frmTest.frx":091E
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox PicIcon 
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   1
         Left            =   720
         Picture         =   "frmTest.frx":0C28
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   3
         Top             =   240
         Width           =   540
      End
      Begin VB.PictureBox PicIcon 
         AutoSize        =   -1  'True
         Height          =   540
         Index           =   0
         Left            =   120
         Picture         =   "frmTest.frx":0F32
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   2
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "We can animate these cute ones..."
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3045
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add_Tray / Hide_Me"
      Height          =   975
      Left            =   3360
      MouseIcon       =   "frmTest.frx":123C
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Adding the tray]
Private Sub cmdAdd_Click()
    ' Add it
    TrayAdd hwnd, Me.Icon, "System Tray", MouseMove
    ' For animated icon
    tmrUpdate.Enabled = True
    Me.WindowState = vbMinimized
    Me.Hide
    MsgBox "The events on tray icon will send to immediate window" & _
            vbCrLf & "(Right click on tray-icon to popup menu)", vbInformation
End Sub

'[Checking The mouse event]
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cEvent As Single
cEvent = X / Screen.TwipsPerPixelX
Select Case cEvent
    Case MouseMove
        Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown"
    Case LeftDbClick
        Debug.Print "LeftDbClick"
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        Debug.Print "RightUp": PopupMenu mnuForm
    Case RightDown
        Debug.Print "RightDown"
    Case RightDbClick
        Debug.Print "RightDbClick"
End Select
End Sub

Private Sub mnuShow_Click()
    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete  '[Deleting Tray]
    tmrUpdate.Enabled = False
End Sub

Private Sub tmrUpdate_Timer()
Static mIcon As Long
    If mIcon = 4 Then mIcon = 0
    TrayModify Tray_Icon, PicIcon(mIcon).Picture
    mIcon = mIcon + 1
End Sub

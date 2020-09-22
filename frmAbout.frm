VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About this application"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   4935
      Begin VB.Label Label2 
         Caption         =   "For Updates or others cool applications, source code, ActiveX ...  Go to VTech http://vtech.ifrance.com or to the Dev Zone"
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4575
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "http://vtech.ifrance.com/vtech/devzone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   2280
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   2655
      Left            =   120
      Picture         =   "frmAbout.frx":0000
      Top             =   120
      Width           =   5010
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com/vtech/devzone", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label1_Click()
    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com/vtech/devzone", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub Label2_Click()

    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com/vtech/devzone", vbNullString, vbNullString, vbNormalFocus)

End Sub

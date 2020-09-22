VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMail 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "HMTL Mailer"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
   Icon            =   "frmEmail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   7320
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   20
      Top             =   7080
      Width           =   7095
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6960
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lvlVTech 
         Alignment       =   2  'Center
         Caption         =   "For professionals ActiveX, go to VTech - CLICK HERE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   6855
      End
      Begin VB.Label lblDevZone 
         Alignment       =   2  'Center
         Caption         =   "For more applicattions, VB Tools, ActiveX and free source code : go to Dev Zone - CLICK HERE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   120
         TabIndex        =   21
         Top             =   120
         Width           =   6855
      End
   End
   Begin VB.TextBox txtEmailFromMail 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   840
      Width           =   2655
   End
   Begin VB.CommandButton cmdFile 
      Height          =   375
      Left            =   6720
      Picture         =   "frmEmail.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Add attachment ..."
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "&About / Updates"
      Height          =   495
      Left            =   4920
      Picture         =   "frmEmail.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "&Quit"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   8400
      Width           =   2295
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "&Send mail"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   8400
      Width           =   2295
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   6360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtEMailTo 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   4
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   5
      Top             =   1560
      Width           =   5775
   End
   Begin VB.TextBox txtEMailFrom 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   2
      Top             =   480
      Width           =   2655
   End
   Begin VB.TextBox txtSMTP 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   7095
      Begin VB.TextBox txtLog 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Contents"
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7095
      Begin VB.CommandButton cmdRemoveAttachment 
         Height          =   375
         Left            =   6600
         Picture         =   "frmEmail.frx":109E
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Remove attachment"
         Top             =   840
         Width           =   375
      End
      Begin VB.ListBox lstAttachments 
         Height          =   840
         ItemData        =   "frmEmail.frx":1628
         Left            =   1080
         List            =   "frmEmail.frx":162A
         TabIndex        =   23
         Top             =   360
         Width           =   5415
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         Height          =   2655
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Text            =   "frmEmail.frx":162C
         Top             =   1320
         Width           =   6855
      End
      Begin VB.Label Label6 
         Caption         =   "Note : Press CTRL+Enter to go to next line in the message area."
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4080
         Width           =   6855
      End
      Begin VB.Label Label5 
         Caption         =   "Attachments"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Email from (mail)"
      Height          =   195
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label4 
      Caption         =   "Subject"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1560
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "Email to"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1200
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Email from (name)"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "SMTP server"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1005
   End
End
Attribute VB_Name = "frmMail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'---- API
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
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

'---- Const
Const err_SMTP = "No SMTP server"
Const err_FROM = "No Email from"
Const err_TO = "No Email to"
Const err_SUBJECT = "No subject"

Dim response As String


'----

Private Function ValidateFormData() As Boolean

    Dim temp As Boolean
    
    temp = False
    If Len(txtEmailFromMail.Text) = 0 Then
        txtLog.Text = "Error: You need to enter a 'from mail'." & vbCrLf & txtLog.Text & vbCrLf
        temp = True
    End If
    
    If Len(txtEMailTo.Text) = 0 Then
        txtLog.Text = "Error: You need to enter a 'to mail'." & vbCrLf & txtLog.Text & vbCrLf
        temp = True
    End If
    
    If Len(txtEMailFrom.Text) = 0 Then
        txtLog.Text = "Error: You need to enter a 'from name'." & vbCrLf & txtLog.Text & vbCrLf
        temp = True
    End If
    
    If Len(txtSMTP.Text) = 0 Then
        txtLog.Text = "Error: You need to enter a 'smtp server'." & vbCrLf & txtLog.Text & vbCrLf
        temp = True
    End If
    
    ValidateFormData = temp
    
End Function

'------------------------------------------------------------
'           Buttons management
'------------------------------------------------------------

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub cmdFile_Click()

    Dim strFile As String
    
    ' Get a file
    strFile = OpenDialog("Images Files(*.gif;*.jpg)|*.gif;*.jpg", "Choose an image ...", "")
    
    If (strFile <> "") Then
        lstAttachments.AddItem strFile
    End If
    
End Sub

Private Sub cmdRemoveAttachment_Click()
    
    If lstAttachments.ListIndex = -1 Then Exit Sub
    
    lstAttachments.RemoveItem (lstAttachments.ListIndex)
    
End Sub

Private Function OpenDialog(Filter As String, DialogTitle As String, InitialFolder As String) As String
 
    Dim ofn As OPENFILENAME
    Dim a As Long
    
    ofn.lStructSize = Len(ofn)
    ofn.hWndOwner = Me.hwnd
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
    
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitialFolder
    ofn.lpstrTitle = DialogTitle
    ofn.flags = 0 'OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    a = GetOpenFileName(ofn)

    If (a) Then
        OpenDialog = Trim$(ofn.lpstrFile)
    Else
        OpenDialog = ""
    End If

End Function

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSend_Click()
    If ValidateFormData = False Then
        connect_to_smtp_server txtSMTP
    End If
End Sub

'------------------------------------------------------------
'           Form managemement
'------------------------------------------------------------

Sub init_me()

    txtLog.Text = "Ready." & vbCrLf
    response = ""
    
    lstAttachments.AddItem App.Path & "\DZLogo2.jpg"
    lstAttachments.AddItem App.Path & "\HTMLMailSS.jpg"
    
End Sub


Private Sub Form_Load()
    init_me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Winsock1.Close
    Unload Me
End Sub

Private Sub Label8_Click()
    
End Sub

Private Sub lblDevZone_Click()
    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com/vtech/devzone", vbNullString, vbNullString, vbNormalFocus)
End Sub

Private Sub lvlVTech_Click()
    Call ShellExecute(GetDesktopWindow(), vbNullString, "http://vtech.ifrance.com", vbNullString, vbNullString, vbNormalFocus)
End Sub

'------------------------------------------------------------
'           Mail managemement
'------------------------------------------------------------

Private Sub Winsock1_Connect()

    txtLog.Text = "Connected to: " & txtSMTP & "." & vbCrLf & txtLog.Text & vbCrLf
    
    ' Send the mail
    SendMail txtEMailTo, txtEMailFrom, txtEmailFromMail, txtSubject, txtMessage
    
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

    Winsock1.GetData response
    
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

    txtLog.Text = "Error: " & Description & "." & vbCrLf & txtLog.Text & vbCrLf
    
End Sub

Sub wait_for(winsock_answare As String)

    Do While Left(response, 3) <> winsock_answare
        DoEvents
    Loop
    response = ""
    
End Sub

Function find_date() As String

    Dim temp As String
    Dim fd_day As String
    Dim fd_month As String
    Dim fd_time As String
    
    fd_day = Format(Date, "Dddd")
    Select Case fd_day
        Case "éåí øàùåï": fd_day = "Sun, "
        Case "éåí ùðé": fd_day = "Mon, "
        Case "éåí ùìéùé": fd_day = "Tue, "
        Case "éåí øáéòé": fd_day = "Wed, "
        Case "éåí çîéùé": fd_day = "Thu, "
        Case "éåí ùéùé": fd_day = "Fri, "
        Case "éåí ùáú": fd_day = "Sat, "
    End Select
    fd_month = Month(Date)
    Select Case fd_month
        Case 1: fd_month = "Jan "
        Case 2: fd_month = "Feb "
        Case 3: fd_month = "Mar "
        Case 4: fd_month = "Apr "
        Case 5: fd_month = "May "
        Case 6: fd_month = "Jun "
        Case 7: fd_month = "Jul "
        Case 8: fd_month = "Aug "
        Case 9: fd_month = "Sep "
        Case 10: fd_month = "Oct "
        Case 11: fd_month = "Nov "
        Case 12: fd_month = "Dec "
    End Select
    fd_time = Format(Time) & " +0200"
    temp = fd_day & Day(Format(Date)) & " " & fd_month & Year(Format(Date, "dd/mm/yyyy")) & " " & fd_time
    find_date = temp
    
End Function

Function attach_file(attach_str As String) As String
    Dim s As Integer
    Dim temp As String
    
    s = InStr(1, attach_str, "\")
    temp = attach_str
    Do While s > 0
        temp = Mid(temp, s + 1, Len(temp))
        s = InStr(1, temp, "\")
    Loop
    attach_file = temp
End Function

Function encode_the_file(attach_str As String) As String
    Dim blocksize As Long
    Dim buffer As String
    Dim s As String
    Dim i As Long
    Dim temp As String
    
    Open attach_str For Binary Access Read As #1
        blocksize = 3
        Do While Not EOF(1)
            buffer = Space(blocksize)
            Get 1, , buffer
            s = s & base64_encode_string(buffer)
            DoEvents
        Loop
    Close #1
    For i = 1 To Len(s) Step 76
        temp = temp & Mid(s, i, 76) & vbCrLf
    Next i
    temp = Mid(temp, 1, Len(temp) - 2)
    encode_the_file = temp
End Function

' Send the mail
Private Sub SendMail(aMailTo As String, aMailFromName As String, aMailFrom As String, aSubject As String, aMessage As String)
    
    'Const boundary = "Hapoel_Tel_Aviv"
    Const boundary = "NextMimePart"
    
    Dim se_body As String
    Dim se_date As String
    Dim se_from As String
    Dim se_to As String
    Dim se_mime As String
    Dim se_content_type As String
    Dim se_content_type_message As String
    Dim se_content_type_attach As String
    Dim x_mailer As String
    Dim x_oem As String
     
    se_date = "Date: " & find_date
    se_from = "From: " & aMailFromName
    se_to = "To: " & aMailTo
    aSubject = "Subject: " & aSubject
    
    se_mime = "MIME-Version: 1.0"
    se_content_type = "Content-Type: multipart/related;" & vbCrLf _
        & vbTab & "boundary = " & """" & boundary & """"
    
    x_oem = "X-OEM: zubin"
    x_mailer = "X-Mailer: " & """" & "HTMLMail" & """" & " - by VTech http://vtech.ifrance.com"
    
    se_content_type_message = "This is a multi-part message in MIME format." & vbCrLf _
        & "--" & boundary & vbCrLf _
        & "Content-Type: text/html;" & vbCrLf _
        & vbTab & "charset=" & """" & "iso-8859-1" & """" & vbCrLf _
        & "Content-Transfer-Encoding: 7bit"
        
    
    se_content_type_attach = ""
        
    Dim iIndex As Long
    For iIndex = 1 To lstAttachments.ListCount
        
        se_content_type_attach = se_content_type_attach & "--" & boundary & vbCrLf _
            & "Content-Type: application/octet-stream;" & vbCrLf _
            & vbTab & "name=" & attach_file(lstAttachments.List(iIndex - 1)) & vbCrLf _
            & "Content-Transfer-Encoding: base64" & vbCrLf _
            & "Content-Disposition: attachment;" & vbCrLf _
            & vbTab & "filename=" & attach_file(lstAttachments.List(iIndex - 1)) & vbCrLf _
            & vbCrLf _
            & encode_the_file(lstAttachments.List(iIndex - 1)) & vbCrLf _
            
    Next iIndex
    
    se_body = se_from & vbCrLf _
        & se_to & vbCrLf _
        & aSubject & vbCrLf _
        & se_date & vbCrLf _
        & se_mime & vbCrLf _
        & x_oem & vbCrLf _
        & x_mailer & vbCrLf _
        & se_content_type & vbCrLf _
        & vbCrLf _
        & se_content_type_message & vbCrLf _
        & vbCrLf _
        & aMessage & vbCrLf _
        & vbCrLf _
        & se_content_type_attach & vbCrLf _
        & "." & vbCrLf
    
    txtLog.Text = "Sending message..." & vbCrLf & txtLog.Text & vbCrLf
    Winsock1.SendData "HELO " & Left(txtEmailFromMail, InStr(1, txtEmailFromMail, "@") - 1) & vbCrLf
    wait_for "250"
    Winsock1.SendData "MAIL FROM: " & "<" + txtEmailFromMail + ">" & vbCrLf
    wait_for "250"
    
    Dim i As Long
    Dim strToUnique As String
    
    strToUnique = aMailTo & ";"
    Call Replace(strToUnique, ";;", ";")
    
    While (InStr(1, strToUnique, ";") <> 0 Or InStr(1, strToUnique, ","))
        Dim strTo As String
        
        If (InStr(1, strToUnique, ";") <> 0) Then
            strTo = Mid(strToUnique, 1, InStr(1, strToUnique, ";") - 1)
        Else
            strTo = Mid(strToUnique, 1, InStr(1, strToUnique, ",") - 1)
        End If
        
        Winsock1.SendData "RCPT TO: " & "<" & Trim$(strTo) & ">" & vbCrLf
        wait_for "250"
        
        strToUnique = Mid(strToUnique, InStr(1, strToUnique, ";") + 1)
    Wend

    Winsock1.SendData "DATA" & vbCrLf
    wait_for "354"
    Winsock1.SendData se_body
    wait_for "250"
    Winsock1.SendData "QUIT" & vbCrLf
    wait_for "221"
    txtLog.Text = "Message sent." & vbCrLf & txtLog.Text & vbCrLf
    Winsock1.Close
    
    DoEvents
    
End Sub

' Do the connection with the SMTP server
Sub connect_to_smtp_server(smtp_server As String)

    Winsock1.LocalPort = 0
    Winsock1.RemoteHost = txtSMTP
    Winsock1.RemotePort = 25
    Winsock1.Connect
    
End Sub


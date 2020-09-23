VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00400000&
   Caption         =   "Welcome"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4500
   FillColor       =   &H00808080&
   Icon            =   "frmGetPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmGetPass.frx":0E42
   ScaleHeight     =   2265
   ScaleWidth      =   4500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000003&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1650
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000003&
      Caption         =   "&Exit"
      Height          =   375
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1650
      Width           =   975
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2430
      MaxLength       =   8
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1110
      Width           =   1215
   End
   Begin VB.TextBox txtUserID 
      Height          =   285
      Left            =   2430
      MaxLength       =   8
      TabIndex        =   0
      Top             =   750
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   600
      Picture         =   "frmGetPass.frx":3426
      Top             =   810
      Width           =   480
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   1140
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   4
      Top             =   780
      Width           =   1035
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
   Unload Me
   End
End Sub

Private Sub cmdOK_Click()
  Static nTries As Byte
  
  rstUsers.Seek "=", txtUserId

  If Not rstUsers.NoMatch Then
    If nTries < 2 Then
      If rstUsers("Password") = txtPassword Then
        nLevel = Val(CStr(rstUsers("Level")))
''        nLevel = 7 'for testing only/temporary
        Unload Me
        frmMain.Show 1
      Else
        nTries = nTries + 1
        MsgBox "Invalid Password: " & CStr(nTries), vbCritical
        txtPassword.SetFocus
      End If
    Else
      MsgBox "Access Denied!", vbCritical
      Unload Me
      End
    End If
  Else
    MsgBox "Unauthorized Access!", vbCritical
    txtUserId.SetFocus
  End If
End Sub

Private Sub Form_Load()
  'Go to Application Path
  ChDir (App.Path)
  
  Set dbsInfoSys = OpenDatabase(App.Path & "\infosys.mdb")
  Set rstUsers = dbsInfoSys.OpenRecordset("Users")
  rstUsers.Index = "UserId"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstUsers.Close
  dbsInfoSys.Close
End Sub

'- UDF
Private Sub txtPassword_GotFocus()
  Call FocusMe(txtPassword)
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
    End
  End If
  
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
  
End Sub

Private Sub txtUserId_GotFocus()
  Call FocusMe(txtUserId)
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then
    Unload Me
    End
  End If
  
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: UDF

VERSION 5.00
Begin VB.Form frmUsers 
   BackColor       =   &H00FFFFFF&
   Caption         =   "User Maintenance"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4005
   Icon            =   "Users.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Users.frx":0E42
   ScaleHeight     =   3900
   ScaleWidth      =   4005
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox txtLevel 
      Height          =   315
      ItemData        =   "Users.frx":53D1
      Left            =   1710
      List            =   "Users.frx":53DE
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2730
      Width           =   795
   End
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H80000003&
      Caption         =   "&Remove"
      Height          =   375
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "&Close"
      Height          =   375
      Left            =   1980
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000003&
      Caption         =   "&Save"
      Height          =   375
      Left            =   990
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3480
      Width           =   975
   End
   Begin VB.TextBox txtConfirm 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "aaaaaaaa"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1710
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2340
      Width           =   1305
   End
   Begin VB.TextBox txtPassword 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "aaaaaaaa"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1710
      MaxLength       =   15
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1980
      Width           =   1305
   End
   Begin VB.TextBox txtUserId 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "LLLLLLLL"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      Height          =   285
      Left            =   1710
      MaxLength       =   8
      TabIndex        =   0
      Top             =   1620
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User's Maintenance"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   240
      TabIndex        =   11
      Top             =   840
      Width           =   3525
   End
   Begin VB.Label Label4 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   630
      TabIndex        =   10
      Top             =   2790
      Width           =   855
   End
   Begin VB.Label Label3 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm "
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   630
      TabIndex        =   9
      Top             =   2370
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   630
      TabIndex        =   8
      Top             =   1980
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00400000&
      BackStyle       =   0  'Transparent
      Caption         =   "User ID"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   630
      TabIndex        =   7
      Top             =   1620
      Width           =   615
   End
End
Attribute VB_Name = "frmUsers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSave_Click()
  If ValidateEntries Then
    ' Confirm Saving..
    If MsgBox("Click Ok to Save", vbInformation + vbOKCancel) = vbOK Then
      If Not CheckUserId Then
        rstUsers.AddNew 'New record
      Else
        rstUsers.Edit 'Existing User
      End If
      rstUsers("UserId") = txtUserId.Text
      rstUsers("Password") = txtPassword.Text
      rstUsers("Level") = txtLevel.Text
      rstUsers.Update
      
      txtUserId = ""
      Initialize
      'Set focus to User Id
      txtUserId.SetFocus
    End If
  End If
  
End Sub

Private Sub cmdRemove_Click()
  If CheckUserId Then 'If User exist?
    If MsgBox("Click Ok to confirm deletion", vbInformation + vbOKCancel) = vbOK Then
      rstUsers.Edit
      rstUsers.Delete
      
      txtUserId = ""
      Initialize
      
      'Set focus to User Id
      txtUserId.SetFocus
    End If
  Else
    MsgBox ("UserId does not exist")
  End If
End Sub


Function Initialize()
  txtPassword = ""
  txtConfirm = ""
  txtLevel = 1
  cmdSave.Caption = "Save"
End Function

Private Sub txtUserId_Lostfocus()
  If CheckUserId Then
    txtLevel = rstUsers("Level")
    cmdSave.Caption = "Edit"
  Else
    Initialize
  End If
End Sub

Function CheckUserId()
  CheckUserId = False
  rstUsers.Seek "=", txtUserId.Text
  If Not rstUsers.NoMatch Then
    CheckUserId = True
  End If
End Function

Function ValidateEntries()
  ValidateEntries = True
  If Len(txtUserId.Text) < 4 Then
    MsgBox ("User Id must be at least 4 chars in length")
    ValidateEntries = False
    txtUserId.SetFocus
    Exit Function
  End If
  If Len(txtPassword.Text) < 8 Then
    MsgBox ("Password must be 8 chars in length")
    ValidateEntries = False
    txtPassword.SetFocus
    Exit Function
  End If
  If txtPassword.Text <> txtConfirm Then
    MsgBox ("Password not confirmed correctly")
    ValidateEntries = False
    txtConfirm.SetFocus
    Exit Function
  End If
  If txtLevel.Text <> "1" And txtLevel.Text <> "2" And txtLevel.Text <> "3" Then
    MsgBox ("Pls. enter 1 or 2 or 3 access level")
    ValidateEntries = False
    txtLevel.SetFocus
    Exit Function
  End If
  If Val(txtLevel) > nLevel Then
    MsgBox "You can not create access level which is greater than yours. Pls inform the administrator.", vbCritical
    ValidateEntries = False
    txtLevel.SetFocus
    Exit Function
  End If
End Function

'- Add'l UDF
Private Sub txtUserId_GotFocus()
  Call FocusMe(txtUserId)
End Sub

Private Sub txtPassword_GotFocus()
  Call FocusMe(txtPassword)
End Sub

Private Sub txtConfirm_Gotfocus()
  Call FocusMe(txtConfirm)
End Sub

Private Sub txtUserID_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtLevel_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: Add'l UDF

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstUsers = dbsInfoSys.OpenRecordset("Users")
  rstUsers.Index = "UserId"
  
  Initialize
  
  cmdRemove.Visible = nLevel > 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstUsers.Close
  dbsInfoSys.Close
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


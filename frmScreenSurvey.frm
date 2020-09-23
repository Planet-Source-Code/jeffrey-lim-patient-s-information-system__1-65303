VERSION 5.00
Begin VB.Form frmScreenSurvey 
   Caption         =   "Medical Screen Survey"
   ClientHeight    =   3945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   Icon            =   "frmScreenSurvey.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmScreenSurvey.frx":08CA
   ScaleHeight     =   3945
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H0080FF80&
      Caption         =   "Gen"
      Height          =   345
      Left            =   3300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   900
      Width           =   615
   End
   Begin VB.ComboBox cboDisease 
      Height          =   315
      ItemData        =   "frmScreenSurvey.frx":4E59
      Left            =   840
      List            =   "frmScreenSurvey.frx":4E5B
      TabIndex        =   0
      Text            =   "cboDisease"
      Top             =   900
      Width           =   2385
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H80000003&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   2970
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   330
      X2              =   3690
      Y1              =   2670
      Y2              =   2670
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   0
      X1              =   330
      X2              =   3690
      Y1              =   2070
      Y2              =   2070
   End
   Begin VB.Label lblTransferred 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2310
      TabIndex        =   9
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblAdmitted 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2490
      TabIndex        =   8
      Top             =   2160
      Width           =   1155
   End
   Begin VB.Label lblTreated 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   345
      Left            =   2550
      TabIndex        =   7
      Top             =   1620
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Untreated"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   360
      TabIndex        =   6
      Top             =   2760
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Recovering"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   330
      TabIndex        =   5
      Top             =   2160
      Width           =   2235
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   960
      Width           =   825
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Treated/dischrg"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   330
      TabIndex        =   3
      Top             =   1590
      Width           =   2535
   End
End
Attribute VB_Name = "frmScreenSurvey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdGenerate_Click()
  Dim nTreated, nAdmitted, nTransferred As Long
  
  nTreated = 0
  nTransferred = 0
  nAdmitted = 0
  If rstDiagnos.RecordCount > 0 Then
    rstDiagnos.MoveFirst
    Do Until rstDiagnos.EOF
      If Trim(rstDiagnos("Disease")) = Trim(UCase(cboDisease)) Then
        If rstDiagnos("Disposition") = "T" Then
          nTreated = nTreated + 1 'Treated/Discharged
        ElseIf rstDiagnos("Disposition") = "S" Then
          nTransferred = nTransferred + 1 'Transferred/Discharged
        ElseIf rstDiagnos("Disposition") = "U" Then
          nTransferred = nTransferred + 1 'Untreated/Dead/Discharged
        Else
          nAdmitted = nAdmitted + 1
        End If
      End If
      rstDiagnos.MoveNext
    Loop
  End If
  
  lblTreated = Str(nTreated)
  lblAdmitted = Str(nAdmitted)
  lblTransferred = Str(nTransferred)

End Sub

Private Sub cboDisease_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstDisease = dbsInfoSys.OpenRecordset("Disease")
  Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
  rstDisease.Index = "Disease"
  rstDiagnos.Index = "Disease"
  
  GetDisease
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstDiagnos.Close
  rstDisease.Close
  dbsInfoSys.Close
End Sub

Function GetDisease()
  rstDisease.Index = "Disease"
  'cboDisease.Style = 0
  If rstDisease.RecordCount > 0 Then
  cboDisease = rstDisease("Disease")
  Do Until rstDisease.EOF
    cboDisease.AddItem UCase(rstDisease("Disease"))
    rstDisease.MoveNext
  Loop
  'cboDisease.Style = 2
  End If
End Function

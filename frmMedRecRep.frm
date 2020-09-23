VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmMedRecRep 
   Caption         =   "Medical Record Generator"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8265
   Icon            =   "frmMedRecRep.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMedRecRep.frx":08CA
   ScaleHeight     =   3630
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox nHospNo 
      Height          =   285
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1050
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      MaxLength       =   35
      TabIndex        =   9
      Top             =   1350
      Width           =   3195
   End
   Begin VB.TextBox txtAddress1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      MaxLength       =   50
      TabIndex        =   8
      Top             =   1680
      Width           =   4455
   End
   Begin VB.TextBox txtFamilyDoctor 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1290
      MaxLength       =   35
      TabIndex        =   7
      Top             =   2400
      Width           =   3225
   End
   Begin VB.TextBox nAge 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6810
      MaxLength       =   20
      TabIndex        =   6
      Top             =   1440
      Width           =   645
   End
   Begin VB.ComboBox cboSocialClass 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMedRecRep.frx":6F7E
      Left            =   7080
      List            =   "frmMedRecRep.frx":6F97
      Style           =   1  'Simple Combo
      TabIndex        =   5
      Top             =   2430
      Width           =   1065
   End
   Begin VB.ComboBox cboCivilStatus 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmMedRecRep.frx":6FB4
      Left            =   6810
      List            =   "frmMedRecRep.frx":6FC4
      Style           =   1  'Simple Combo
      TabIndex        =   4
      Top             =   1770
      Width           =   1305
   End
   Begin VB.TextBox dBirthDate 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6810
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H80000003&
      Caption         =   "&Generate"
      Height          =   375
      Left            =   6210
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Generate and Print Medical Record"
      Top             =   3210
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "&Close"
      Height          =   375
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      Width           =   975
   End
   Begin MSMask.MaskEdBox nTelNo 
      Height          =   285
      Left            =   1290
      TabIndex        =   10
      Top             =   2010
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   " "
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's No."
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
      Left            =   210
      TabIndex        =   19
      Top             =   1080
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Left            =   210
      TabIndex        =   18
      Top             =   1410
      Width           =   1095
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
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
      Left            =   210
      TabIndex        =   17
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tel. No"
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
      Left            =   210
      TabIndex        =   16
      Top             =   2070
      Width           =   1095
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Birthdate"
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
      Left            =   5910
      TabIndex        =   15
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
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
      Left            =   5910
      TabIndex        =   14
      Top             =   1470
      Width           =   825
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Family Doctor"
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
      Left            =   210
      TabIndex        =   13
      Top             =   2430
      Width           =   1095
   End
   Begin VB.Label Label19 
      BackStyle       =   0  'Transparent
      Caption         =   "Social Class"
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
      Left            =   6180
      TabIndex        =   12
      Top             =   2490
      Width           =   855
   End
   Begin VB.Label Label20 
      BackStyle       =   0  'Transparent
      Caption         =   "Civil Status"
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
      Left            =   5910
      TabIndex        =   11
      Top             =   1830
      Width           =   885
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchPatient 
         Caption         =   "&Patient"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMedRecRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdGenerate_Click()
  Dim nHN As Integer
  
  If Trim(nHospNo) <> "" Then
    nHN = Val(nHospNo)
    
    denvMedRec.cmdMedRec (nHospNo)
    rptMedRec.DataMember = "cmdMedRec"
    rptMedRec.Show 1
    
    Unload rptMedRec
    Unload denvMedRec
  Else
    MsgBox "No record to generate..", vbCritical
    nHospNo.SetFocus
  End If
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
  rstPatients.Index = "HospNo"

  Init 'Clear all entry fields
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstPatients.Close
  dbsInfoSys.Close
End Sub

'- clear all entry fields
Function Init()
  txtName = ""
  txtAddress1 = ""
  nTelNo = ""
  dBirthDate = ""
  nAge = ""
  cboCivilStatus = ""
  cboSocialClass = ""
  txtFamilyDoctor = ""
End Function

Function CheckHospNo() As Boolean
  CheckHospNo = False
  If Len(Trim(nHospNo)) > 0 Then
    rstPatients.Seek "=", Val(nHospNo)
    If Not rstPatients.NoMatch Then
      CheckHospNo = True
    Else
      MsgBox "Patient's personal record not found.. Pls register the patient first", vbInformation
    End If
  End If
End Function

Private Sub mnuSearchPatient_Click()
  frmSearchPatient.Show 1
  nHospNo = Str(nPatNo)
  nHospNo_LostFocus
  cmdGenerate.SetFocus
End Sub

Private Sub nHospNo_LostFocus()
  If CheckHospNo Then
    txtName = Trim(rstPatients("Lastname")) & ", " & Trim(rstPatients("Firstname")) & " " & Trim(rstPatients("Middlename"))
    txtAddress1 = rstPatients("Address1")
    nTelNo = rstPatients("TelNo")
    cboCivilStatus = rstPatients("CivilStatus")
    cboSocialClass = rstPatients("SocialClass")
    dBirthDate = rstPatients("Birthdate")
    nAge = GetAge(dBirthDate)
    txtFamilyDoctor = rstPatients("FamilyDoctor")
  Else
    Init
  End If
End Sub


Private Sub nHospNo_GotFocus()
  Init
  Call FocusMe(nHospNo)
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

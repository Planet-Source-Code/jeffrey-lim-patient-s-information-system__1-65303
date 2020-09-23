VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmit 
   Caption         =   "Admit Patient"
   ClientHeight    =   4890
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9045
   Icon            =   "frmAdmit.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAdmit.frx":0E42
   ScaleHeight     =   4890
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox cboSocialClass 
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
      Left            =   7140
      MaxLength       =   35
      TabIndex        =   29
      Top             =   2280
      Width           =   1605
   End
   Begin VB.TextBox cbocivilstatus 
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
      Left            =   6900
      MaxLength       =   35
      TabIndex        =   28
      Top             =   1590
      Width           =   1605
   End
   Begin VB.TextBox nHospNo 
      Height          =   285
      Left            =   1380
      MaxLength       =   7
      TabIndex        =   0
      Top             =   870
      Width           =   975
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
      Left            =   6900
      MaxLength       =   20
      TabIndex        =   27
      Top             =   900
      Width           =   1215
   End
   Begin VB.ComboBox cboRoomType 
      Height          =   315
      ItemData        =   "frmAdmit.frx":8D14
      Left            =   1920
      List            =   "frmAdmit.frx":8D21
      TabIndex        =   3
      Top             =   3660
      Width           =   1485
   End
   Begin VB.ComboBox cboRoomNo 
      Height          =   315
      ItemData        =   "frmAdmit.frx":8D3C
      Left            =   1920
      List            =   "frmAdmit.frx":8D4C
      TabIndex        =   2
      Top             =   3300
      Width           =   1275
   End
   Begin VB.ComboBox cboDisease 
      Height          =   315
      ItemData        =   "frmAdmit.frx":8D6C
      Left            =   5430
      List            =   "frmAdmit.frx":8D6E
      TabIndex        =   4
      Top             =   2970
      Width           =   2625
   End
   Begin VB.TextBox txtAttendingPhysician 
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
      Left            =   5430
      MaxLength       =   35
      TabIndex        =   5
      Top             =   3660
      Width           =   3225
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000003&
      Caption         =   "&Admit"
      Height          =   375
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4350
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "&Close"
      Height          =   375
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4350
      Width           =   975
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
      Left            =   6900
      MaxLength       =   20
      TabIndex        =   14
      Top             =   1260
      Width           =   645
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
      Left            =   1380
      MaxLength       =   35
      TabIndex        =   13
      Top             =   2220
      Width           =   3225
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
      Left            =   1380
      MaxLength       =   50
      TabIndex        =   12
      Top             =   1500
      Width           =   4455
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
      Left            =   1380
      MaxLength       =   35
      TabIndex        =   11
      Top             =   1170
      Width           =   3195
   End
   Begin MSMask.MaskEdBox nTelNo 
      Height          =   285
      Left            =   1380
      TabIndex        =   15
      Top             =   1830
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
   Begin MSComCtl2.DTPicker dDate 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   2940
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55836673
      CurrentDate     =   38382
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Diagnosed Disease"
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
      Left            =   3840
      TabIndex        =   26
      Top             =   3030
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Attending Physician"
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
      Left            =   3810
      TabIndex        =   25
      Top             =   3690
      Width           =   1485
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
      Left            =   6000
      TabIndex        =   24
      Top             =   1650
      Width           =   885
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
      Left            =   6270
      TabIndex        =   23
      Top             =   2310
      Width           =   855
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
      Left            =   300
      TabIndex        =   22
      Top             =   2250
      Width           =   1095
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
      Left            =   6000
      TabIndex        =   21
      Top             =   1290
      Width           =   825
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
      Left            =   6000
      TabIndex        =   20
      Top             =   960
      Width           =   855
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
      Left            =   300
      TabIndex        =   19
      Top             =   1890
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
      Left            =   300
      TabIndex        =   18
      Top             =   1560
      Width           =   1095
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
      Left            =   300
      TabIndex        =   17
      Top             =   1230
      Width           =   1095
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
      Left            =   300
      TabIndex        =   16
      Top             =   900
      Width           =   1125
   End
   Begin VB.Label Label31 
      BackStyle       =   0  'Transparent
      Caption         =   "Date Admitted"
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
      Left            =   660
      TabIndex        =   10
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label Label32 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
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
      Left            =   660
      TabIndex        =   9
      Top             =   3360
      Width           =   1185
   End
   Begin VB.Label Label33 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Type"
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
      Left            =   660
      TabIndex        =   8
      Top             =   3720
      Width           =   1035
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchPatient 
         Caption         =   "Patient"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmAdmit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim lDiagnosDiseaseFound As Boolean
  
  If ValidateEntries Then
    If CheckHospNo Then
      ' Confirm Saving..
      If MsgBox("Pls click Ok to confirm admittance of the patient.", vbInformation + vbOKCancel) = vbOK Then
        rstPatients.Edit
        rstPatients("PatientStatus") = "IN"
        rstPatients.Update
      
        'AdmitStatus
        If CheckAdmitStatus Then
          rstAdmitStatus.Edit
        Else
          rstAdmitStatus.AddNew
        End If
        rstAdmitStatus("Date") = dDate
        rstAdmitStatus("HospNo") = nHospNo
        rstAdmitStatus("RoomNo") = UCase(cboRoomNo)
        rstAdmitStatus("RoomType") = UCase(cboRoomType)
        rstAdmitStatus("Disease") = UCase(cboDisease)
        rstAdmitStatus("PatStatus") = "IN"
        rstAdmitStatus("AttendingPhysician") = UCase(txtAttendingPhysician)
        
        'Remove previous discharge info
        rstAdmitStatus("DateDischarge") = ""
        rstAdmitStatus("Disposition") = ""
        rstAdmitStatus("ConditionOfDischarge") = ""
        rstAdmitStatus.Update
        
        'Update Diagnos Disease List
        lDiagnosDiseaseFound = False
        If rstDiagnos.RecordCount > 0 Then
          Do Until rstDiagnos.EOF
            If Trim(rstDiagnos("HospNo")) = Trim(nHospNo) And _
               rstDiagnos("Disease") = UCase(cboDisease) Then
               lDiagnosDiseaseFound = True
               Exit Do
            End If
            rstDiagnos.MoveNext
          Loop
        End If
        If Not lDiagnosDiseaseFound Then
          rstDiagnos.AddNew
          rstDiagnos("Date") = dDate
          rstDiagnos("Disease") = UCase(cboDisease)
          rstDiagnos("HospNo") = nHospNo
          rstDiagnos("PatStatus") = "IN"
          rstDiagnos.Update
        End If
        
        'Clear all fields
        Init
        nHospNo = ""
        nHospNo.SetFocus
      End If
    Else
      MsgBox "Unable to admit unregistered patient. Pls create patient's info first.", vbCritical
    End If
  End If
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
  Set rstAdmitStatus = dbsInfoSys.OpenRecordset("AdmitStatus")
  Set rstDisease = dbsInfoSys.OpenRecordset("Disease")
  Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
  rstPatients.Index = "HospNo"
  rstAdmitStatus.Index = "HospNo"
  rstDiagnos.Index = "HospNo"

  Init 'Clear all entry fields
  
  GetDisease
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstDiagnos.Close
  rstDisease.Close
  rstAdmitStatus.Close
  rstPatients.Close
  dbsInfoSys.Close
End Sub

'- clear all entry fields
Function Init()
  cmdSave.Enabled = True
  txtName = ""
  txtAddress1 = ""
  nTelNo = ""
  dBirthDate = ""
  nAge = ""
  dDate = Date
  cbocivilstatus = ""
  cboSocialClass = ""
  txtFamilyDoctor = ""
  txtAttendingPhysician = ""
End Function

Function ValidateEntries() As Boolean
  If Val(nHospNo) < 1 Then
    ValidateEntries = False
    MsgBox "Invalid Hospital Number", vbCritical
    nHospNo.SetFocus
    Exit Function
  End If
  If dDate > Date Then
    ValidateEntries = False
    MsgBox "Invalid admit date", vbCritical
    dDate.SetFocus
    Exit Function
  End If
  If Trim(cboRoomNo) = "" Then
    ValidateEntries = False
    MsgBox "Room number required", vbCritical
    cboRoomNo.SetFocus
    Exit Function
  End If
  If Trim(cboRoomType) = "" Then
    ValidateEntries = False
    MsgBox "Room type required", vbCritical
    cboRoomType.SetFocus
    Exit Function
  End If
  If Trim(cboDisease) = "" Then
    ValidateEntries = False
    MsgBox "Initial diagnosed disease required", vbCritical
    cboDisease.SetFocus
    Exit Function
  End If
  If Trim(txtAttendingPhysician) = "" Then
    ValidateEntries = False
    MsgBox "Attending physician required", vbCritical
    txtAttendingPhysician.SetFocus
    Exit Function
  End If
  ValidateEntries = True
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

Function CheckAdmitStatus() As Boolean
  CheckAdmitStatus = False
  rstAdmitStatus.Seek "=", Val(nHospNo)
  If Not rstAdmitStatus.NoMatch Then
    CheckAdmitStatus = True
  End If
End Function

Function GetDataFromAdmitStatus()
  If CheckAdmitStatus Then
    If Trim(rstAdmitStatus("DateDischarge")) = "" Then 'not yet discharge
      cmdSave.Caption = "Edit"
      cmdSave.Enabled = nLevel > 1
      dDate = rstAdmitStatus("Date")
      cboRoomNo = rstAdmitStatus("RoomNo")
      cboRoomType = rstAdmitStatus("RoomType")
      cboDisease = rstAdmitStatus("Disease")
      txtAttendingPhysician = rstAdmitStatus("AttendingPhysician")
    Else
      If MsgBox("Patient's was discharged last " & Trim(rstAdmitStatus("DateDischarge")) & ". Would you like to re-confine this patient?", vbInformation + vbOKCancel) = vbOK Then
        dDate = Date
        dDate.SetFocus
      Else
        'Get another patient
        Init
        nHospNo.SetFocus
      End If
    End If
  End If
End Function



Private Sub nHospNo_LostFocus()
  cmdSave.Caption = "Admit"
  If CheckHospNo Then
    txtName = Trim(rstPatients("Lastname")) & ", " & Trim(rstPatients("Firstname")) & " " & Trim(rstPatients("Middlename"))
    txtAddress1 = rstPatients("Address1")
    nTelNo = rstPatients("TelNo")
    cbocivilstatus = rstPatients("CivilStatus")
    cboSocialClass = rstPatients("SocialClass")
    dBirthDate = rstPatients("Birthdate")
    nAge = GetAge(dBirthDate)
    txtFamilyDoctor = rstPatients("FamilyDoctor")
    
    'GetDataFromAdmitStatus
    GetDataFromAdmitStatus
  Else
    Init
  End If
End Sub

Function GetDisease()
  rstDisease.Index = "Disease"
  If rstDisease.RecordCount > 0 Then
  cboDisease = rstDisease("Disease")
  Do Until rstDisease.EOF
    cboDisease.AddItem UCase(rstDisease("Disease"))
    rstDisease.MoveNext
  Loop
  End If
End Function

Private Sub mnuSearchPatient_Click()
  frmSearchPatient.Show 1
  'nHospNo.Mask = ""
  nHospNo = Str(nPatNo)
  'nHospNo.Mask = "#######"
  'MsgBox "Auto-retrieve record of patient number " & Trim(Str(nPatNo)), vbInformation
  nHospNo_LostFocus
  dDate.SetFocus
End Sub

'- UDF
Private Sub nHospNo_GotFocus()
  Init
  Call FocusMe(nHospNo)
End Sub

Private Sub txtAttendingPhysician_Gotfocus()
  Call FocusMe(txtAttendingPhysician)
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtAttendingPhysician_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboRoomNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboRoomType_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboDisease_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: UDF


VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDischarge 
   Caption         =   "Discharge Patient"
   ClientHeight    =   5295
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9075
   Icon            =   "frmDischarge.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmDischarge.frx":014A
   ScaleHeight     =   5295
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtNationality 
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
      Left            =   9360
      MaxLength       =   20
      TabIndex        =   38
      Top             =   3960
      Width           =   2445
   End
   Begin VB.ComboBox cboSex 
      Height          =   315
      ItemData        =   "frmDischarge.frx":82A9
      Left            =   9450
      List            =   "frmDischarge.frx":82B3
      Style           =   1  'Simple Combo
      TabIndex        =   37
      Text            =   "cboSex"
      Top             =   600
      Width           =   1185
   End
   Begin VB.ComboBox cboCivilStatus 
      Height          =   315
      ItemData        =   "frmDischarge.frx":82C5
      Left            =   9450
      List            =   "frmDischarge.frx":82D5
      Style           =   1  'Simple Combo
      TabIndex        =   36
      Text            =   "cboCivilStatus"
      Top             =   960
      Width           =   1425
   End
   Begin VB.TextBox txtLastname 
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
      Left            =   9330
      MaxLength       =   35
      TabIndex        =   34
      Top             =   1950
      Width           =   3195
   End
   Begin VB.TextBox txtFirstname 
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
      Left            =   9300
      MaxLength       =   35
      TabIndex        =   33
      Top             =   2280
      Width           =   3195
   End
   Begin VB.TextBox txtMiddlename 
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
      Left            =   9300
      MaxLength       =   35
      TabIndex        =   32
      Top             =   2580
      Width           =   3195
   End
   Begin VB.TextBox txtAddress1 
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
      Left            =   9210
      MaxLength       =   50
      TabIndex        =   31
      Top             =   3150
      Width           =   4455
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
      Left            =   9300
      MaxLength       =   20
      TabIndex        =   30
      Top             =   1350
      Width           =   645
   End
   Begin VB.TextBox nHospNo 
      Height          =   285
      Left            =   1800
      MaxLength       =   7
      TabIndex        =   0
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox cboDisease 
      Enabled         =   0   'False
      Height          =   285
      Left            =   6150
      TabIndex        =   29
      Top             =   270
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00DBAD8E&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   4830
      TabIndex        =   27
      Top             =   3960
      Width           =   4035
      Begin VB.OptionButton optCritical 
         BackColor       =   &H00DBAD8E&
         Caption         =   "&Critical"
         Height          =   315
         Left            =   3030
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
      Begin VB.OptionButton optSerious 
         BackColor       =   &H00DBAD8E&
         Caption         =   "Serio&us"
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   240
         Width           =   885
      End
      Begin VB.OptionButton optSatisfactory 
         BackColor       =   &H00DBAD8E&
         Caption         =   "Satis&factory"
         Height          =   315
         Left            =   930
         TabIndex        =   7
         Top             =   240
         Width           =   1185
      End
      Begin VB.OptionButton optGood 
         BackColor       =   &H00DBAD8E&
         Caption         =   "&Good"
         Height          =   315
         Left            =   150
         TabIndex        =   6
         Top             =   240
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Condition of Discharge"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DBAD8E&
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   270
      TabIndex        =   25
      Top             =   3960
      Width           =   4395
      Begin VB.OptionButton optDead 
         BackColor       =   &H00DBAD8E&
         Caption         =   "Untreated"
         Height          =   285
         Left            =   3330
         TabIndex        =   5
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton optTransferred 
         BackColor       =   &H00DBAD8E&
         Caption         =   "Tran&sferred"
         Height          =   285
         Left            =   2100
         TabIndex        =   4
         Top             =   240
         Width           =   1305
      End
      Begin VB.OptionButton optTreated 
         BackColor       =   &H00DBAD8E&
         Caption         =   "&Treated and Discharge"
         Height          =   285
         Left            =   90
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   2025
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Disposition"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   30
         TabIndex        =   26
         Top             =   0
         Width           =   1185
      End
   End
   Begin RichTextLib.RichTextBox rtbFinalDiagnosis 
      Height          =   1605
      Left            =   210
      TabIndex        =   1
      Top             =   2310
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2831
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDischarge.frx":82FA
   End
   Begin VB.TextBox cboRoomType 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   19
      Top             =   1590
      Width           =   1335
   End
   Begin VB.TextBox cboRoomNo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   18
      Top             =   1260
      Width           =   1065
   End
   Begin VB.TextBox dDateAdmitted 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7320
      TabIndex        =   17
      Top             =   930
      Width           =   1395
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
      Left            =   1800
      MaxLength       =   35
      TabIndex        =   13
      Top             =   1260
      Width           =   3525
   End
   Begin VB.TextBox txtAttendingPhysician 
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
      Left            =   1800
      MaxLength       =   35
      TabIndex        =   12
      Top             =   1590
      Width           =   3525
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "&Close"
      Height          =   375
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4860
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000003&
      Caption         =   "&Discharge"
      Height          =   375
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4860
      Width           =   975
   End
   Begin RichTextLib.RichTextBox rtbRemarks 
      Height          =   1605
      Left            =   4740
      TabIndex        =   2
      Top             =   2310
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   2831
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDischarge.frx":837C
   End
   Begin MSComCtl2.DTPicker dBirthDate 
      Height          =   315
      Left            =   9300
      TabIndex        =   35
      Top             =   1410
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20512769
      CurrentDate     =   38382
   End
   Begin MSMask.MaskEdBox nTelNo 
      Height          =   285
      Left            =   9330
      TabIndex        =   39
      Top             =   3660
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
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
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Final Diagnosis"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   24
      Top             =   2100
      Width           =   2205
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Special Instructions"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4770
      TabIndex        =   23
      Top             =   2100
      Width           =   2025
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
      Left            =   6060
      TabIndex        =   22
      Top             =   1680
      Width           =   1035
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
      Left            =   6060
      TabIndex        =   21
      Top             =   1320
      Width           =   1185
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
      Left            =   6060
      TabIndex        =   20
      Top             =   990
      Width           =   1185
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
      Left            =   240
      TabIndex        =   16
      Top             =   960
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
      Left            =   240
      TabIndex        =   15
      Top             =   1290
      Width           =   1095
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
      Left            =   240
      TabIndex        =   14
      Top             =   1620
      Width           =   1485
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchPatient 
         Caption         =   "&Patient"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmDischarge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdSave_Click()
  Dim strDisposition, strConditionOfDischarge As String
  Dim lMedRecNew As Boolean
  
  If ValidateEntries Then
    If CheckHospNo Then
      ' Confirm Saving..
      If MsgBox("Pls click Ok to confirm Discharge of the patient.", vbInformation + vbOKCancel) = vbOK Then
        rstPatients.Edit
        rstPatients("PatientStatus") = "OUT"
        rstPatients.Update
      
        'AdmitStatus
        If CheckAdmitStatus Then
          rstAdmitStatus.Edit
          
          'discharge info
          strDisposition = "T" 'Treated
          If optTransferred = True Then
            strDisposition = "S" 'Transferred
          ElseIf optDead = True Then
            strDisposition = "U" 'Untreated/Dead/Discharged
          End If
          
          strConditionOfDischarge = "G"
          If optSatisfactory Then
            strConditionOfDischarge = "F"
          ElseIf optSerious Then
            strConditionOfDischarge = "U"
          ElseIf optCritical Then
            strConditionOfDischarge = "C"
          End If
        
          rstAdmitStatus("DateDischarge") = Date
          rstAdmitStatus("Disposition") = strDisposition
          rstAdmitStatus("ConditionOfDischarge") = strConditionOfDischarge
          rstAdmitStatus("PatStatus") = "OUT"
          rstAdmitStatus.Update
        End If
        
        'Update Diagnos Disease List
        If rstDiagnos.RecordCount > 0 Then
          Do Until rstDiagnos.EOF
            If Trim(rstDiagnos("HospNo")) = Trim(nHospNo) And _
               rstDiagnos("Disease") = UCase(cboDisease) Then
               rstDiagnos.Edit
               rstDiagnos("Disposition") = strDisposition
               rstDiagnos("PatStatus") = "OUT"
               rstDiagnos.Update
               Exit Do
            End If
            rstDiagnos.MoveNext
          Loop
        End If
        
        'Medical Record
        lMedRecNew = True
        If rstMedRec.RecordCount > 0 Then
          rstMedRec.MoveFirst
          Do Until rstMedRec.EOF
          If Trim(rstMedRec("HospNo")) = Trim(nHospNo) And _
            rstMedRec("Date") = Date And _
            rstMedRec("Discharge") = True Then 'edit
            lMedRecNew = False
            Exit Do
          End If
          rstMedRec.MoveNext
         Loop
        End If
        If lMedRecNew Then 'New medical record
          rstMedRec.AddNew
        Else
          rstMedRec.Edit
        End If
        rstMedRec("HospNo") = nHospNo
        rstMedRec("Date") = Date
        rstMedRec("Firstname") = txtFirstname
        rstMedRec("Middlename") = txtMiddlename
        rstMedRec("Lastname") = txtLastname
        rstMedRec("Address") = txtAddress1
        rstMedRec("Age") = nAge
        rstMedRec("Sex") = cboSex
        rstMedRec("Nationality") = txtNationality
        rstMedRec("TelNo") = nTelNo
        rstMedRec("CivilStatus") = cboCivilStatus
        rstMedRec("DateOfArrival") = dDateAdmitted
        rstMedRec("FinalDiagnosis") = UCase(rtbFinalDiagnosis.Text)
        rstMedRec("Remarks") = UCase(rtbRemarks.Text)
        rstMedRec("Disposition") = strDisposition
        rstMedRec("ConditionOfDischarge") = strConditionOfDischarge
        rstMedRec("Discharge") = True
        rstMedRec.Update
        
        'Clear all fields
        Init
        nHospNo = ""
        nHospNo.SetFocus
      End If
    Else
      MsgBox "Unable to discharge unregistered patient. Pls create patient's info first.", vbCritical
    End If
  End If
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
  Set rstAdmitStatus = dbsInfoSys.OpenRecordset("AdmitStatus")
  Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
  Set rstMedRec = dbsInfoSys.OpenRecordset("MedRec")
  rstMedRec.Index = "HospNo"
  rstPatients.Index = "HospNo"
  rstAdmitStatus.Index = "HospNo"
  rstDiagnos.Index = "HospNo"

  Init 'Clear all entry fields
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstMedRec.Close
  rstDiagnos.Close
  rstAdmitStatus.Close
  rstPatients.Close
  dbsInfoSys.Close
End Sub

'- clear all entry fields
Function Init()
  cmdSave.Caption = "Discharge"
  cmdSave.Enabled = True
  txtName = ""
  txtAttendingPhysician = ""
  cboDisease = ""
  cboRoomNo = ""
  cboRoomType = ""
  dDateAdmitted = Date '""
  rtbFinalDiagnosis = ""
  rtbRemarks = ""
  
  txtFirstname = ""
  txtMiddlename = ""
  txtLastname = ""
  txtAddress1 = ""
  nAge = 0
End Function

Function ValidateEntries() As Boolean
  If Val(nHospNo) < 1 Then
    ValidateEntries = False
    MsgBox "Invalid Hospital Number", vbCritical
    nHospNo.SetFocus
    Exit Function
  End If
'  If Trim(cboRoomNo) = "" Then
'    ValidateEntries = False
'    MsgBox "Room number required", vbCritical
'    cboRoomNo.SetFocus
'    Exit Function
'  End If
'  If Trim(cboRoomType) = "" Then
'    ValidateEntries = False
'    MsgBox "Room type required", vbCritical
'    cboRoomType.SetFocus
'    Exit Function
'  End If
  ValidateEntries = True
End Function

Function CheckHospNo() As Boolean
  CheckHospNo = False
  If Len(Trim(nHospNo)) > 0 Then
    rstPatients.Seek "=", Val(nHospNo)
    If Not rstPatients.NoMatch Then
      CheckHospNo = True
      
      txtFirstname = rstPatients("Firstname")
      txtLastname = rstPatients("Lastname")
      txtMiddlename = rstPatients("Middlename")
      txtAddress1 = rstPatients("Address1")
      dBirthDate = rstPatients("Birthdate")
      cboSex = rstPatients("Sex")
      txtNationality = rstPatients("Nationality")
      nTelNo = rstPatients("TelNo")
      nAge = GetAge(dBirthDate)
      cboCivilStatus = rstPatients("CivilStatus")
    Else
      MsgBox "Patient's personal record not found.. Pls register the patient first", vbInformation
    End If
  End If
End Function

Function CheckMedRec() As Boolean
  CheckMedRec = False
  If Len(Trim(nHospNo)) > 0 Then
    rstMedRec.Seek "=", Val(nHospNo)
    If Not rstMedRec.NoMatch Then
      CheckMedRec = True
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
    If Trim(rstAdmitStatus("DateDischarge")) <> "" Then 'not yet discharge
      cmdSave.Caption = "Edit"
    End If
      dDateAdmitted = rstAdmitStatus("Date")
      cboDisease = rstAdmitStatus("Disease")
      cboRoomNo = rstAdmitStatus("RoomNo")
      cboRoomType = rstAdmitStatus("RoomType")
      txtAttendingPhysician = rstAdmitStatus("AttendingPhysician")
  End If
End Function


Function GetDataFromMedRec()
  If rstMedRec.RecordCount > 0 Then
    rstMedRec.MoveFirst
    Do Until rstMedRec.EOF
      If Trim(rstMedRec("HospNo")) = Trim(nHospNo) And _
         rstMedRec("Date") = Date And _
         rstMedRec("Discharge") = True Then 'edit
        
        cmdSave.Caption = "Edit"
        cmdSave.Enabled = nLevel > 1
        'cboDisease = rstMedRec("Disease")
        rtbFinalDiagnosis = rstMedRec("FinalDiagnosis")
        rtbRemarks = rstMedRec("Remarks")
        txtAttendingPhysician = rstMedRec("AttendingPhysician")
        'Disposition
        optTreated = (rstMedRec("Disposition") = "T")
        optTransferred = (rstMedRec("Disposition") = "S")
        optDead = (rstMedRec("Disposition") = "U")
        'Condition of Discharged
        optGood = (rstMedRec("ConditionOfDischarge") = "G")
        optSatisfactory = (rstMedRec("ConditionOfDischarge") = "F")
        optSerious = (rstMedRec("ConditionOfDischarge") = "U")
        optCritical = (rstMedRec("ConditionOfDischarge") = "C")
        Exit Do
        
      End If
      rstMedRec.MoveNext
    Loop
  End If
End Function



Private Sub nHospNo_LostFocus()
  cmdSave.Caption = "Discharge"
  If CheckHospNo Then
    
    txtName = Trim(rstPatients("Lastname")) & ", " & Trim(rstPatients("Firstname")) & " " & Trim(rstPatients("Middlename"))
    
    'GetDataFromAdmitStatus
    GetDataFromAdmitStatus
    GetDataFromMedRec
  Else
    Init
  End If
End Sub

Private Sub mnuSearchPatient_Click()
  frmSearchPatient.Show 1
  nHospNo = Str(nPatNo)
  nHospNo_LostFocus
  rtbFinalDiagnosis.SetFocus
End Sub

'- UDF
Private Sub nHospNo_GotFocus()
  Init
  Call FocusMe(nHospNo)
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: UDF



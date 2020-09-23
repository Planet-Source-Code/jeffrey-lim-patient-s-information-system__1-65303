VERSION 5.00
Begin VB.Form frmInitialize 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Clear all tables"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4020
   Icon            =   "Initialize.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Initialize.frx":014A
   ScaleHeight     =   389
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   268
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkLastNo 
      BackColor       =   &H00DBAD8E&
      Caption         =   "Last Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   5
      Top             =   3900
      Width           =   2505
   End
   Begin VB.CheckBox chkDisease 
      BackColor       =   &H00DBAD8E&
      Caption         =   "Diseases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   4
      Top             =   3480
      Width           =   2475
   End
   Begin VB.CheckBox chkDiagnos 
      BackColor       =   &H00DBAD8E&
      Caption         =   "Diagnos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   3
      Top             =   3060
      Width           =   2325
   End
   Begin VB.CheckBox chkAdmitStatus 
      BackColor       =   &H00DBAD8E&
      Caption         =   "Admit Status"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   2
      Top             =   2640
      Width           =   2385
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000003&
      Caption         =   "&Close"
      Height          =   375
      Left            =   2790
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H80000003&
      Caption         =   "&Ok"
      Height          =   375
      Left            =   1650
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CheckBox chkUsers 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Users"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   6
      Top             =   4320
      Width           =   1965
   End
   Begin VB.CheckBox chkMedRec 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Medical Records"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   780
      TabIndex        =   1
      Top             =   2220
      Width           =   2475
   End
   Begin VB.CheckBox chkPatients 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Patient's Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   780
      TabIndex        =   0
      Top             =   1890
      Width           =   2205
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Clear all tables"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   270
      TabIndex        =   10
      Top             =   900
      Width           =   2745
   End
   Begin VB.Label lblMessage 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   270
      TabIndex        =   9
      Top             =   4680
      Width           =   3435
   End
End
Attribute VB_Name = "frmInitialize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOK_Click()
  
  If chkPatients.Value = 0 And _
     chkMedRec.Value = 0 And _
     chkAdmitStatus.Value = 0 And _
     chkDiagnos.Value = 0 And _
     chkDisease.Value = 0 And _
     chkLastNo.Value = 0 And _
     chkUsers.Value = 0 Then
     
     MsgBox "Please select atleast one.", vbExclamation
     
     
    Exit Sub
  Else
    If MsgBox("Click Ok to confirm deletion", vbInformation + vbOKCancel) = vbOK Then
      Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
      
      If chkPatients.Value Then
        Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
        If rstPatients.RecordCount > 0 Then
          Do Until rstPatients.EOF
            'rstPatients.Edit
            rstPatients.Delete
            rstPatients.MoveNext
            'rstPatients.Update
          Loop
        End If
        rstPatients.AddNew
        rstPatients("HospNo") = 1
        rstPatients("Firstname") = "WALK-IN PATIENT"
        rstPatients("Lastname") = "WALK-IN PATIENT"
        rstPatients("Middlename") = "WALK-IN PATIENT"
        rstPatients("Birthdate") = "12/25/1900"
        rstPatients.Update
        rstPatients.Close
      End If
      
      If chkMedRec.Value Then
        Set rstMedRec = dbsInfoSys.OpenRecordset("MedRec")
        If rstMedRec.RecordCount > 0 Then
          rstMedRec.Edit
          Do Until rstMedRec.EOF
            rstMedRec.Delete
            rstMedRec.MoveNext
          Loop
          'rstMedRec.Update
          rstMedRec.Close
        End If
      End If
      
      If chkAdmitStatus.Value Then
        Set rstAdmitStatus = dbsInfoSys.OpenRecordset("AdmitStatus")
        If rstAdmitStatus.RecordCount > 0 Then
          rstAdmitStatus.Edit
          Do Until rstAdmitStatus.EOF
            rstAdmitStatus.Delete
            rstAdmitStatus.MoveNext
          Loop
          rstAdmitStatus.Close
        End If
      End If
      
      If chkDiagnos.Value Then
        Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
        If rstDiagnos.RecordCount > 0 Then
          rstDiagnos.Edit
          Do Until rstDiagnos.EOF
            rstDiagnos.Delete
            rstDiagnos.MoveNext
          Loop
          rstDiagnos.Close
        End If
      End If
      
      If chkDisease.Value Then
        Set rstDisease = dbsInfoSys.OpenRecordset("Disease")
        If rstDisease.RecordCount > 0 Then
          rstDisease.Edit
          Do Until rstDisease.EOF
            rstDisease.Delete
            rstDisease.MoveNext
          Loop
          rstDisease.Close
        End If
      End If
      
      If chkLastNo.Value Then
        Set rstLastNo = dbsInfoSys.OpenRecordset("LastNo")
        If rstLastNo.RecordCount > 0 Then
          'rstLastNo.Edit
          Do Until rstLastNo.EOF
            rstLastNo.Delete
            rstLastNo.MoveNext
          Loop
        End If
        rstLastNo.AddNew
        rstLastNo("HospNo") = 1
        rstLastNo.Update
        rstLastNo.Close
        
      End If
      
      If chkUsers.Value Then
        Set rstUsers = dbsInfoSys.OpenRecordset("Users")
        If rstUsers.RecordCount > 0 Then
          rstUsers.Edit
          Do Until rstUsers.EOF
            rstUsers.Delete
            rstUsers.MoveNext
          Loop
          rstUsers.AddNew
          rstUsers("UserId") = "admin"
          rstUsers("Password") = "admin"
          rstUsers("Level") = 5
          rstUsers.Update
          rstUsers.Close
          
          MsgBox ("Default: UserID = admin, Password = admin")
          
        End If
      End If
   
      dbsInfoSys.Close
    End If
  End If
  End Sub



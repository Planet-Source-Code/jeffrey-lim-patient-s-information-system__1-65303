VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmSearchPatient 
   Caption         =   "Search patient"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10125
   Icon            =   "frmSearchPatient.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSearchPatient.frx":08CA
   ScaleHeight     =   8040
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox dBirthdate 
      DataField       =   "Birthdate"
      DataSource      =   "adoPatients"
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   44
      ToolTipText     =   "ex.: 5' 11"""
      Top             =   1380
      Width           =   1485
   End
   Begin VB.TextBox nHospNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   8280
      MaxLength       =   7
      TabIndex        =   2
      Top             =   120
      Width           =   1635
   End
   Begin VB.CommandButton cmdPrevious 
      BackColor       =   &H80000003&
      Caption         =   "&Previous"
      Height          =   375
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Display previous record"
      Top             =   7620
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "&OK"
      Height          =   375
      Left            =   8970
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Close this form or select Hospital number"
      Top             =   7620
      Width           =   1095
   End
   Begin VB.CommandButton cmdNext 
      BackColor       =   &H80000003&
      Caption         =   "&Next"
      Height          =   375
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Display next record"
      Top             =   7620
      Width           =   1095
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
      Left            =   210
      MaxLength       =   35
      TabIndex        =   0
      Top             =   1020
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
      Left            =   3450
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1020
      Width           =   3195
   End
   Begin VB.TextBox txtMiddlename 
      DataField       =   "MiddleName"
      DataSource      =   "adoPatients"
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
      Left            =   6690
      MaxLength       =   35
      TabIndex        =   21
      Top             =   1020
      Width           =   3195
   End
   Begin VB.TextBox txtAddress1 
      DataField       =   "Address1"
      DataSource      =   "adoPatients"
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
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   20
      Top             =   1440
      Width           =   4455
   End
   Begin VB.TextBox txtAddress2 
      DataField       =   "Address2"
      DataSource      =   "adoPatients"
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
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   19
      Top             =   1770
      Width           =   4455
   End
   Begin VB.TextBox txtGuardian 
      DataField       =   "Guardian"
      DataSource      =   "adoPatients"
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
      Left            =   1530
      MaxLength       =   35
      TabIndex        =   18
      Top             =   3990
      Width           =   3195
   End
   Begin VB.TextBox txtGAddress 
      DataField       =   "GAddress"
      DataSource      =   "adoPatients"
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
      Left            =   1530
      MaxLength       =   50
      TabIndex        =   17
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox txtRelationship 
      DataField       =   "Relationship"
      DataSource      =   "adoPatients"
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
      Left            =   1530
      MaxLength       =   20
      TabIndex        =   16
      Top             =   4650
      Width           =   2445
   End
   Begin VB.TextBox txtOccupation 
      DataField       =   "Occupation"
      DataSource      =   "adoPatients"
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
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   15
      Top             =   2430
      Width           =   2445
   End
   Begin VB.TextBox txtNationality 
      DataField       =   "Nationality"
      DataSource      =   "adoPatients"
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
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   14
      Top             =   2760
      Width           =   2445
   End
   Begin VB.TextBox txtFamilyDoctor 
      DataField       =   "FamilyDoctor"
      DataSource      =   "adoPatients"
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
      Left            =   1530
      MaxLength       =   35
      TabIndex        =   13
      Top             =   5010
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   12
      Top             =   1710
      Width           =   645
   End
   Begin VB.TextBox txtHeight 
      DataField       =   "Height"
      DataSource      =   "adoPatients"
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   11
      ToolTipText     =   "ex.: 5' 11"""
      Top             =   2040
      Width           =   915
   End
   Begin VB.TextBox nWeight 
      DataField       =   "Weight"
      DataSource      =   "adoPatients"
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
      Left            =   7560
      MaxLength       =   20
      TabIndex        =   10
      ToolTipText     =   "Weight in lbs"
      Top             =   2340
      Width           =   705
   End
   Begin VB.ComboBox cboSex 
      DataField       =   "Sex"
      DataSource      =   "adoPatients"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearchPatient.frx":A535
      Left            =   1350
      List            =   "frmSearchPatient.frx":A53F
      Style           =   1  'Simple Combo
      TabIndex        =   9
      Top             =   3090
      Width           =   1185
   End
   Begin VB.ComboBox cboSocialClass 
      DataField       =   "SocialClass"
      DataSource      =   "adoPatients"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearchPatient.frx":A551
      Left            =   8370
      List            =   "frmSearchPatient.frx":A56A
      Style           =   1  'Simple Combo
      TabIndex        =   8
      Top             =   2850
      Width           =   1125
   End
   Begin VB.ComboBox cboPatientStatus 
      DataField       =   "PatientStatus"
      DataSource      =   "adoPatients"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearchPatient.frx":A587
      Left            =   8370
      List            =   "frmSearchPatient.frx":A591
      Style           =   1  'Simple Combo
      TabIndex        =   7
      Top             =   3210
      Width           =   855
   End
   Begin VB.ComboBox cboCivilStatus 
      DataField       =   "CivilStatus"
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmSearchPatient.frx":A59E
      Left            =   1350
      List            =   "frmSearchPatient.frx":A5AE
      Style           =   1  'Simple Combo
      TabIndex        =   6
      Top             =   3450
      Width           =   1425
   End
   Begin MSMask.MaskEdBox nTelNo 
      DataField       =   "TelNo"
      DataSource      =   "adoPatients"
      Height          =   285
      Left            =   1350
      TabIndex        =   22
      Top             =   2100
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmSearchPatient.frx":A5D3
      Height          =   2025
      Left            =   210
      TabIndex        =   43
      Top             =   5430
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   3572
      _Version        =   393216
      AllowUpdate     =   0   'False
      Appearance      =   0
      BackColor       =   14396814
      ForeColor       =   16777215
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   3
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "lbs"
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
      Left            =   8310
      TabIndex        =   46
      Top             =   2370
      Width           =   735
   End
   Begin VB.Label Label21 
      BackStyle       =   0  'Transparent
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
      Left            =   8520
      TabIndex        =   45
      Top             =   2070
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's No."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   6570
      TabIndex        =   42
      Top             =   180
      Width           =   1605
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Lastname"
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
      TabIndex        =   41
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Firstname"
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
      Left            =   3450
      TabIndex        =   40
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Miiddlename"
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
      Left            =   6720
      TabIndex        =   39
      Top             =   840
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
      TabIndex        =   38
      Top             =   1500
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
      TabIndex        =   37
      Top             =   2160
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
      Left            =   6690
      TabIndex        =   36
      Top             =   1410
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
      Left            =   6690
      TabIndex        =   35
      Top             =   1740
      Width           =   825
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
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
      TabIndex        =   34
      Top             =   3120
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Occupation"
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
      TabIndex        =   33
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
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
      Left            =   6690
      TabIndex        =   32
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Weight"
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
      Left            =   6690
      TabIndex        =   31
      Top             =   2370
      Width           =   735
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
      Left            =   390
      TabIndex        =   30
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Status"
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
      Left            =   7020
      TabIndex        =   29
      Top             =   3210
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Nationality"
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
      TabIndex        =   28
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Guardian"
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
      Left            =   390
      TabIndex        =   27
      Top             =   4050
      Width           =   1095
   End
   Begin VB.Label Label17 
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
      Left            =   390
      TabIndex        =   26
      Top             =   4350
      Width           =   1065
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Relationship"
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
      Left            =   390
      TabIndex        =   25
      Top             =   4680
      Width           =   1005
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
      Left            =   7020
      TabIndex        =   24
      Top             =   2910
      Width           =   1065
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
      Left            =   210
      TabIndex        =   23
      Top             =   3480
      Width           =   885
   End
End
Attribute VB_Name = "frmSearchPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Conn As New ADODB.Connection
Dim adoPatients As New ADODB.Recordset

Private Sub cmdClose_Click()
  nPatNo = Val(Trim(nHospNo))
  Unload Me
End Sub


Private Sub Form_Unload(Cancel As Integer)
  Set adoPatients = Nothing
  Set Conn = Nothing
End Sub

Private Sub Form_Load()
  Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\InfoSys.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password="
  adoPatients.CursorLocation = adUseClient
  
  txtLastname_LostFocus
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  If Not adoPatients.BOF And Not adoPatients.EOF Then
    nAge = GetAge(dBirthDate)
  End If
End Sub

Private Sub cmdPrevious_Click()
  If Not adoPatients.BOF Then
    adoPatients.MovePrevious
    nAge = GetAge(dBirthDate)
  Else
    adoPatients.MoveFirst
  End If
End Sub

Private Sub cmdNext_Click()
  If Not adoPatients.EOF Then
    adoPatients.MoveNext
    nAge = GetAge(dBirthDate)
  Else
    adoPatients.MoveLast
  End If
End Sub

Function Search(Optional lHospNo As Boolean)
  Dim strSQL As String
  
  If adoPatients.State = 1 Then Set adoPatients = Nothing
    
  strSQL = "SELECT * FROM Patients "
  
  If lHospNo Then
    nHospNo.DataField = ""
    strSQL = strSQL & " Where " & nHospNo & " = HospNo"
    Debug.Print strSQL
  Else
    nHospNo.DataField = "HospNo"
    If Trim(txtLastname) <> "" Then
        strSQL = strSQL & "Where [Lastname] Like '%" & Trim(UCase(txtLastname)) & "%'"
        'strSQL = strSQL & "Where [Lastname] Like '" & Trim(UCase(txtLastname)) & "%'" 'removed as per panelist 2.21.5
    End If
    If Trim(txtLastname) <> "" And _
        Trim(txtFirstname) <> "" Then
        strSQL = strSQL & " and "
    End If
    If Trim(txtLastname) = "" And _
        Trim(txtFirstname) <> "" Then
        strSQL = strSQL & " Where "
    End If
    If Trim(txtFirstname) <> "" Then
        'strSQL = strSQL & " [FirstName] Like '%" & Trim(UCase(txtFirstname)) & "%'" 'remove as per panelist 2.21.5
        strSQL = strSQL & " [FirstName] Like '" & Trim(UCase(txtFirstname)) & "%'"
    End If
    strSQL = strSQL & " Order by [LastName], [Firstname]"
  End If
  
  adoPatients.CursorLocation = adUseClient
  adoPatients.Open strSQL, Conn, adOpenKeyset, adLockOptimistic 'adOpenDynamic, adLockPessimistic"
  
  Set DataGrid1.DataSource = adoPatients
  
  Set txtMiddlename.DataSource = adoPatients
  Set txtAddress1.DataSource = adoPatients
  Set txtAddress2.DataSource = adoPatients
  Set nTelNo.DataSource = adoPatients
  Set txtOccupation.DataSource = adoPatients
  Set txtNationality.DataSource = adoPatients
  Set cboSex.DataSource = adoPatients
  Set cboCivilStatus.DataSource = adoPatients
  Set dBirthDate.DataSource = adoPatients
  Set txtHeight.DataSource = adoPatients
  Set nWeight.DataSource = adoPatients
  Set cboSocialClass.DataSource = adoPatients
  Set cboPatientStatus.DataSource = adoPatients
  Set txtGuardian.DataSource = adoPatients
  Set txtGAddress.DataSource = adoPatients
  Set txtRelationship.DataSource = adoPatients
  Set txtFamilyDoctor.DataSource = adoPatients
  
  If Not lHospNo Then
    nHospNo.DataField = "HospNo"
    Set nHospNo.DataSource = adoPatients
  End If

End Function

Private Sub nHospNo_GotFocus()
  Call FocusMe(nHospNo)
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
  If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
End Sub

Private Sub nHospNo_LostFocus()
  Search (True)
  DataGrid1.SetFocus
End Sub

Private Sub txtLastname_LostFocus()
  Search (False)
End Sub

Private Sub txtFirstname_LostFocus()
  Search (False)
  DataGrid1.SetFocus
End Sub

'- UDF
Private Sub txtFirstname_Gotfocus()
  Call FocusMe(txtFirstname)
End Sub

Private Sub txtLastname_Gotfocus()
  Call FocusMe(txtLastname)
End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: UDF


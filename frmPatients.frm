VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPatients 
   BackColor       =   &H00DBAD8E&
   Caption         =   "Patient's information"
   ClientHeight    =   8070
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10125
   Icon            =   "frmPatients.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmPatients.frx":0E42
   ScaleHeight     =   8070
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox nHospNo 
      Height          =   285
      Left            =   1350
      MaxLength       =   7
      TabIndex        =   0
      Top             =   870
      Width           =   975
   End
   Begin VB.ComboBox cboCivilStatus 
      Height          =   315
      ItemData        =   "frmPatients.frx":AAAD
      Left            =   7200
      List            =   "frmPatients.frx":AABD
      TabIndex        =   13
      Top             =   2490
      Width           =   1425
   End
   Begin VB.ComboBox cboPatientStatus 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "frmPatients.frx":AAE2
      Left            =   8010
      List            =   "frmPatients.frx":AAEC
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ComboBox cboSocialClass 
      Height          =   315
      ItemData        =   "frmPatients.frx":AAF9
      Left            =   8010
      List            =   "frmPatients.frx":AB12
      TabIndex        =   14
      Top             =   3000
      Width           =   1125
   End
   Begin VB.ComboBox cboSex 
      Height          =   315
      ItemData        =   "frmPatients.frx":AB2F
      Left            =   7200
      List            =   "frmPatients.frx":AB39
      TabIndex        =   12
      Top             =   2130
      Width           =   1185
   End
   Begin VB.TextBox nWeight 
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
      Left            =   7170
      MaxLength       =   3
      TabIndex        =   11
      ToolTipText     =   "Weight in lbs"
      Top             =   1830
      Width           =   705
   End
   Begin VB.TextBox txtHeight 
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
      Left            =   7170
      MaxLength       =   20
      TabIndex        =   10
      ToolTipText     =   "ex.: 5' 11"""
      Top             =   1560
      Width           =   825
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
      Left            =   7170
      MaxLength       =   20
      TabIndex        =   41
      Top             =   1260
      Width           =   645
   End
   Begin VB.TextBox txtFamilyDoctor 
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
      MaxLength       =   35
      TabIndex        =   19
      Top             =   5010
      Width           =   3225
   End
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
      Left            =   1350
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3480
      Width           =   2445
   End
   Begin VB.TextBox txtOccupation 
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
      TabIndex        =   7
      Top             =   3150
      Width           =   2445
   End
   Begin VB.TextBox txtRelationship 
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
      TabIndex        =   18
      Top             =   4650
      Width           =   2445
   End
   Begin VB.TextBox txtGAddress 
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
      TabIndex        =   17
      Top             =   4320
      Width           =   4455
   End
   Begin VB.TextBox txtGuardian 
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
      MaxLength       =   35
      TabIndex        =   16
      Top             =   3990
      Width           =   3195
   End
   Begin VB.TextBox txtAddress2 
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
      TabIndex        =   5
      Top             =   2490
      Width           =   4455
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
      Left            =   1350
      MaxLength       =   50
      TabIndex        =   4
      Top             =   2160
      Width           =   4455
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
      Left            =   1350
      MaxLength       =   35
      TabIndex        =   3
      Top             =   1830
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
      Left            =   1350
      MaxLength       =   35
      TabIndex        =   2
      Top             =   1500
      Width           =   3195
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
      Left            =   1350
      MaxLength       =   35
      TabIndex        =   1
      Top             =   1170
      Width           =   3195
   End
   Begin MSAdodcLib.Adodc adoPatients 
      Height          =   405
      Left            =   330
      Top             =   6900
      Visible         =   0   'False
      Width           =   9525
      _ExtentX        =   16801
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\InfoSys.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\InfoSys.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Patients"
      Caption         =   "Patients"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmPatients.frx":AB4B
      Height          =   2025
      Left            =   180
      TabIndex        =   39
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
   Begin MSComCtl2.DTPicker dBirthDate 
      Height          =   315
      Left            =   7170
      TabIndex        =   9
      Top             =   900
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      _Version        =   393216
      Format          =   55771137
      CurrentDate     =   38382
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000003&
      Caption         =   "Save"
      Height          =   375
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7620
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "Close"
      Height          =   375
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7620
      Width           =   1095
   End
   Begin MSMask.MaskEdBox nTelNo 
      Height          =   285
      Left            =   1350
      TabIndex        =   6
      Top             =   2820
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "(###) ###-####"
      PromptChar      =   " "
   End
   Begin VB.Label Label22 
      BackStyle       =   0  'Transparent
      Caption         =   "e.g.: 5' 10"" = 5ft 10 inches"
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
      Left            =   8010
      TabIndex        =   45
      Top             =   1590
      Width           =   1995
   End
   Begin VB.Label Label21 
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
      Left            =   7920
      TabIndex        =   44
      Top             =   1860
      Width           =   735
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
      Left            =   6300
      TabIndex        =   43
      Top             =   2550
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
      Left            =   6660
      TabIndex        =   42
      Top             =   3060
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
      Left            =   210
      TabIndex        =   40
      Top             =   4680
      Width           =   1005
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
      Left            =   210
      TabIndex        =   38
      Top             =   4350
      Width           =   1065
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
      Left            =   210
      TabIndex        =   37
      Top             =   4050
      Width           =   1095
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
      TabIndex        =   36
      Top             =   3480
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
      Left            =   6660
      TabIndex        =   35
      Top             =   3360
      Visible         =   0   'False
      Width           =   1215
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
      TabIndex        =   34
      Top             =   5040
      Width           =   1095
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
      Left            =   6300
      TabIndex        =   33
      Top             =   1890
      Width           =   735
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
      Left            =   6300
      TabIndex        =   32
      Top             =   1590
      Width           =   525
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
      TabIndex        =   31
      Top             =   3180
      Width           =   1095
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
      Left            =   6300
      TabIndex        =   30
      Top             =   2190
      Width           =   705
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
      Left            =   6300
      TabIndex        =   29
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
      Left            =   6300
      TabIndex        =   28
      Top             =   960
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
      TabIndex        =   27
      Top             =   2880
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
      TabIndex        =   26
      Top             =   2220
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
      Left            =   210
      TabIndex        =   25
      Top             =   1890
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
      Left            =   210
      TabIndex        =   24
      Top             =   1560
      Width           =   1095
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
      Left            =   210
      TabIndex        =   23
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
      Left            =   210
      TabIndex        =   22
      Top             =   900
      Width           =   1125
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "Search"
      Begin VB.Menu mnuSearchPatient 
         Caption         =   "Patient"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmPatients"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Function CheckHospNo() As Boolean
  CheckHospNo = False
  rstPatients.Seek "=", Val(nHospNo)
  If Not rstPatients.NoMatch Then
    CheckHospNo = True
    cmdSave.Enabled = nLevel > 1
  End If
End Function

Private Sub cmdSave_Click()
  If ValidateEntries Then
    ' Confirm Saving..
    If MsgBox("Click Ok to Save", vbInformation + vbOKCancel) = vbOK Then
      If Not CheckHospNo Then
        rstPatients.AddNew 'New record
        rstPatients("DateRegistered") = Date
      Else
        rstPatients.Edit 'Existing
      End If
      rstPatients("HospNo") = nHospNo
      rstPatients("Lastname") = UCase(txtLastname)
      rstPatients("Firstname") = UCase(txtFirstname)
      rstPatients("Middlename") = UCase(txtMiddlename)
      rstPatients("Address1") = UCase(txtAddress1)
      rstPatients("Address2") = UCase(txtAddress2)
      rstPatients("TelNo") = nTelNo
      rstPatients("Occupation") = UCase(txtOccupation)
      rstPatients("Nationality") = UCase(txtNationality)
      rstPatients("CivilStatus") = cboCivilStatus
      rstPatients("Birthdate") = dBirthDate
      rstPatients("Age") = nAge
      rstPatients("Height") = UCase(txtHeight)
      rstPatients("Weight") = nWeight
      rstPatients("Sex") = cboSex
      rstPatients("SocialClass") = cboSocialClass
      rstPatients("Guardian") = UCase(txtGuardian)
      rstPatients("GAddress") = UCase(txtGAddress)
      rstPatients("Relationship") = UCase(txtRelationship)
      rstPatients("FamilyDoctor") = UCase(txtFamilyDoctor)
      rstPatients("PatientStatus") = cboPatientStatus
      rstPatients.Update
      adoPatients.Refresh
      UpdateLastNo (nHospNo)
      Init
      
      'Get Lastno
      'nHospNo.Mask = ""
      nHospNo = GetLastNo
      txtLastname.SetFocus
    End If
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub


Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  'nHospNo.Mask = ""
  nHospNo = Str(adoPatients.Recordset("HospNo"))
  nHospNo_LostFocus
End Sub

Private Sub Form_Activate()
  txtLastname.SetFocus
End Sub

Private Sub Form_Load()
  Dim nLastNo As Integer
  
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
  rstPatients.Index = "HospNo"

  Init 'Clear all entry fields
  
  'Get Lastno
  ''nLastNo = GetLastNo
  ''nHospNo.Mask = ""
  nHospNo.Text = Trim(Str(GetLastNo))
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstPatients.Close
  dbsInfoSys.Close
End Sub

'- clear all entry fields
Function Init()
  cmdSave.Caption = "Save"
  cmdSave.Enabled = True
  'nHospNo.Mask = "#######"
  txtLastname = ""
  txtFirstname = ""
  txtMiddlename = ""
  txtAddress1 = ""
  txtAddress2 = ""
  nTelNo.Mask = ""
  nTelNo = ""
  nTelNo.Mask = "(###) ###-####"
  txtOccupation = ""
  txtNationality = ""
  nAge = ""
  txtHeight = ""
  nWeight = ""
  txtGuardian = ""
  txtGAddress = ""
  txtRelationship = ""
  txtFamilyDoctor = ""
  cboPatientStatus = "OUT"
End Function

Function ValidateEntries() As Boolean
  If Val(nHospNo) < 1 Then
    ValidateEntries = False
    MsgBox "Invalid Hospital Number", vbCritical
    nHospNo.SetFocus
    Exit Function
  End If
  
  If IsEmpty(txtLastname) Then
    ValidateEntries = False
    MsgBox "Lastname required", vbCritical
    txtLastname.SetFocus
    Exit Function
  End If
  If IsEmpty(txtFirstname) Then
    ValidateEntries = False
    MsgBox "Firstname required", vbCritical
    txtFirstname.SetFocus
    Exit Function
  End If
  If IsEmpty(txtAddress1) And _
     IsEmpty(txtAddress2) Then
    ValidateEntries = False
    MsgBox "Address required", vbCritical
    txtAddress1.SetFocus
    Exit Function
  End If
  If dBirthDate > Date Then
    ValidateEntries = False
    MsgBox "Invalid birthdate", vbCritical
    dBirthDate.SetFocus
    Exit Function
  End If
  If IsEmpty(cboSex) Then
    ValidateEntries = False
    MsgBox "Sex required", vbCritical
    cboSex.SetFocus
    Exit Function
  End If
  If Val(nWeight) < 1 Then
    ValidateEntries = False
    MsgBox "Invalid weight", vbCritical
    nWeight.SetFocus
    Exit Function
  End If
  ValidateEntries = True
End Function

Private Sub mnuSearchPatient_Click()
  frmSearchPatient.Show 1
  
  '-Actual App
  '- Runtime Error: Invalid Procedure Call
  'nHospNo.Mask = ""
  nHospNo = Str(nPatNo) 'invalid property value
  'nHospNo.Mask = "#######"
  'MsgBox "Auto-retrieve record of patient number " & Trim(Str(nPatNo)), vbInformation
  nHospNo_LostFocus 'Invalid procedure call
  '- eo: Runtime Error:
End Sub

Private Sub nHospNo_LostFocus()
  cmdSave.Caption = "Save"
  If CheckHospNo Then
    cmdSave.Caption = "Edit"
    txtLastname = rstPatients("Lastname")
    txtFirstname = rstPatients("Firstname")
    txtMiddlename = rstPatients("Middlename")
    txtAddress1 = rstPatients("Address1")
    txtAddress2 = rstPatients("Address2")
    nTelNo.Mask = ""
    nTelNo = rstPatients("TelNo")
    nTelNo.Mask = "(###) ###-####"
    txtOccupation = rstPatients("Occupation")
    txtNationality = rstPatients("Nationality")
    cboSex = rstPatients("Sex")
    cboCivilStatus = rstPatients("CivilStatus")
    cboSocialClass = rstPatients("SocialClass")
    cboPatientStatus = rstPatients("PatientStatus")
    dBirthDate = rstPatients("Birthdate")
    nAge = GetAge(dBirthDate)
    txtHeight = rstPatients("Height")
    nWeight = rstPatients("Weight")
    txtGuardian = rstPatients("Guardian")
    txtGAddress = rstPatients("GAddress")
    txtRelationship = rstPatients("Relationship")
    txtFamilyDoctor = rstPatients("FamilyDoctor")
  Else
    Init
  End If
End Sub

Private Sub dBirthDate_LostFocus()
  nAge = GetAge(dBirthDate)
End Sub

'-- UDF
Private Sub nHospNo_GotFocus()
  'nHospNo.Mask = "#######"
  Call FocusMe(nHospNo)
End Sub

Private Sub nTelNo_GotFocus()
  Call FocusMe(nTelNo)
End Sub

Private Sub nWeight_GotFocus()
  Call FocusMe(nWeight)
End Sub

Private Sub txtAddress1_GotFocus()
  Call FocusMe(txtAddress1)
End Sub

Private Sub txtAddress2_GotFocus()
  Call FocusMe(txtAddress2)
End Sub

Private Sub txtFamilyDoctor_GotFocus()
  Call FocusMe(txtFamilyDoctor)
End Sub

Private Sub txtFirstname_Gotfocus()
  Call FocusMe(txtFirstname)
End Sub

Private Sub txtGAddress_GotFocus()
  Call FocusMe(txtGAddress)
End Sub

Private Sub txtGuardian_GotFocus()
  Call FocusMe(txtGuardian)
End Sub

Private Sub txtHeight_GotFocus()
  Call FocusMe(txtHeight)
End Sub

Private Sub txtLastname_Gotfocus()
  Call FocusMe(txtLastname)
End Sub

Private Sub txtMiddlename_GotFocus()
  Call FocusMe(txtMiddlename)
End Sub

Private Sub txtNationality_GotFocus()
  Call FocusMe(txtNationality)
End Sub

Private Sub txtOccupation_GotFocus()
  Call FocusMe(txtOccupation)
End Sub

Private Sub txtRelationship_GotFocus()
  Call FocusMe(txtRelationship)
End Sub

Private Sub dBirthDate_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub


Private Sub cboCivilStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub dBirthDay_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboPatientStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboSex_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboSocialClass_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub nTelNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtAddress2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub nWeight_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtAddress1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFamilyDoctor_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtFirstname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtGuardian_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtGAddress_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtLastname_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtMiddlename_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtNationality_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtRelationship_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'-- eo: UDF

VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMedRec 
   BackColor       =   &H00DBAD8E&
   Caption         =   "Medical Record"
   ClientHeight    =   8010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10125
   Icon            =   "frmMedRec.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMedRec.frx":2870A
   ScaleHeight     =   8010
   ScaleWidth      =   10125
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SStab1 
      Height          =   6735
      Left            =   150
      TabIndex        =   0
      Top             =   780
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   11880
      _Version        =   393216
      TabOrientation  =   2
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BackColor       =   14396814
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Personal Information"
      TabPicture(0)   =   "frmMedRec.frx":2FF42
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "nHospNo"
      Tab(0).Control(1)=   "txtLastname"
      Tab(0).Control(2)=   "txtFirstname"
      Tab(0).Control(3)=   "txtMiddlename"
      Tab(0).Control(4)=   "txtAddress1"
      Tab(0).Control(5)=   "txtAddress2"
      Tab(0).Control(6)=   "txtGuardian"
      Tab(0).Control(7)=   "txtGAddress"
      Tab(0).Control(8)=   "txtRelationship"
      Tab(0).Control(9)=   "txtOccupation"
      Tab(0).Control(10)=   "txtNationality"
      Tab(0).Control(11)=   "txtFamilyDoctor"
      Tab(0).Control(12)=   "nAge"
      Tab(0).Control(13)=   "txtHeight"
      Tab(0).Control(14)=   "nWeight"
      Tab(0).Control(15)=   "cboSex"
      Tab(0).Control(16)=   "cboSocialClass"
      Tab(0).Control(17)=   "cboCivilStatus"
      Tab(0).Control(18)=   "adoPatients"
      Tab(0).Control(19)=   "DataGrid1"
      Tab(0).Control(20)=   "dBirthDate"
      Tab(0).Control(21)=   "nTelNo"
      Tab(0).Control(22)=   "Label1"
      Tab(0).Control(23)=   "Label2"
      Tab(0).Control(24)=   "Label3"
      Tab(0).Control(25)=   "Label4"
      Tab(0).Control(26)=   "Label5"
      Tab(0).Control(27)=   "Label6"
      Tab(0).Control(28)=   "Label7"
      Tab(0).Control(29)=   "Label8"
      Tab(0).Control(30)=   "Label9"
      Tab(0).Control(31)=   "Label10"
      Tab(0).Control(32)=   "Label11"
      Tab(0).Control(33)=   "Label12"
      Tab(0).Control(34)=   "Label13"
      Tab(0).Control(35)=   "Label15"
      Tab(0).Control(36)=   "Label16"
      Tab(0).Control(37)=   "Label17"
      Tab(0).Control(38)=   "Label18"
      Tab(0).Control(39)=   "Label19"
      Tab(0).Control(40)=   "Label20"
      Tab(0).ControlCount=   41
      TabCaption(1)   =   "Medical Records"
      TabPicture(1)   =   "frmMedRec.frx":2FF5E
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Line1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label25"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label26"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label29"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label30"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "rtbRemarks"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "rtbFinalDiagnosis"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdSave"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdClose"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdRemove"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "frmePatientStatus"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Frame1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cboDisease"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtAllergies"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).ControlCount=   14
      Begin VB.TextBox nHospNo 
         Height          =   285
         Left            =   -73440
         MaxLength       =   7
         TabIndex        =   1
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox txtAllergies 
         Height          =   345
         Left            =   1680
         TabIndex        =   56
         Top             =   600
         Width           =   2985
      End
      Begin VB.ComboBox cboDisease 
         Height          =   315
         Left            =   1680
         TabIndex        =   54
         Top             =   210
         Width           =   3015
      End
      Begin VB.Frame Frame1 
         Caption         =   " Doctor's Orders "
         Height          =   2055
         Left            =   780
         TabIndex        =   43
         Top             =   4140
         Width           =   8925
         Begin MSComCtl2.DTPicker dOrderDate 
            Height          =   345
            Left            =   1440
            TabIndex        =   64
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   609
            _Version        =   393216
            Format          =   20578305
            CurrentDate     =   38385
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
            Left            =   1440
            MaxLength       =   35
            TabIndex        =   63
            Top             =   240
            Width           =   3225
         End
         Begin VB.TextBox txtNurseRemarks 
            Height          =   345
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   67
            Top             =   1560
            Width           =   6645
         End
         Begin VB.TextBox txtDocOrders 
            Height          =   345
            Left            =   1440
            MaxLength       =   60
            TabIndex        =   66
            Top             =   1200
            Width           =   7335
         End
         Begin MSMask.MaskEdBox txtOrderTime 
            Height          =   315
            Left            =   1440
            TabIndex        =   65
            Top             =   870
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   5
            Format          =   "hh:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label34 
            BackStyle       =   0  'Transparent
            Caption         =   "Attending Doctor"
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
            Left            =   150
            TabIndex        =   79
            Top             =   270
            Width           =   1365
         End
         Begin VB.Label Label24 
            Caption         =   "Nurse Remarks"
            Height          =   285
            Left            =   150
            TabIndex        =   47
            Top             =   1590
            Width           =   1185
         End
         Begin VB.Label Label23 
            Caption         =   "Doc Orders"
            Height          =   255
            Left            =   150
            TabIndex        =   46
            Top             =   1260
            Width           =   1005
         End
         Begin VB.Label Label22 
            Caption         =   "Order Time"
            Height          =   255
            Left            =   150
            TabIndex        =   45
            Top             =   900
            Width           =   915
         End
         Begin VB.Label Label21 
            Caption         =   "Order Date"
            Height          =   255
            Left            =   150
            TabIndex        =   44
            Top             =   570
            Width           =   975
         End
      End
      Begin VB.Frame frmePatientStatus 
         Caption         =   " Current Medical Status "
         Height          =   3735
         Left            =   5730
         TabIndex        =   40
         Top             =   180
         Width           =   3945
         Begin VB.TextBox cboRoomType 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1530
            TabIndex        =   78
            Top             =   1860
            Width           =   1335
         End
         Begin VB.TextBox cboRoomNo 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1530
            TabIndex        =   77
            Top             =   1470
            Width           =   1065
         End
         Begin VB.TextBox dDateOfArrival 
            Enabled         =   0   'False
            Height          =   345
            Left            =   1530
            TabIndex        =   76
            Top             =   1080
            Width           =   1395
         End
         Begin VB.OptionButton optCritical 
            Caption         =   "&Critical"
            Enabled         =   0   'False
            Height          =   315
            Left            =   3090
            TabIndex        =   62
            Top             =   3300
            Width           =   765
         End
         Begin VB.OptionButton optSerious 
            Caption         =   "Serio&us"
            Enabled         =   0   'False
            Height          =   315
            Left            =   2250
            TabIndex        =   61
            Top             =   3300
            Width           =   885
         End
         Begin VB.OptionButton optSatisfactory 
            Caption         =   "Satis&factory"
            Enabled         =   0   'False
            Height          =   315
            Left            =   1020
            TabIndex        =   60
            Top             =   3300
            Width           =   1185
         End
         Begin VB.OptionButton optGood 
            Caption         =   "&Good"
            Enabled         =   0   'False
            Height          =   315
            Left            =   240
            TabIndex        =   59
            Top             =   3300
            Width           =   735
         End
         Begin VB.OptionButton optTransferred 
            Caption         =   "Tran&sferred"
            Enabled         =   0   'False
            Height          =   285
            Left            =   2370
            TabIndex        =   52
            Top             =   2640
            Width           =   1305
         End
         Begin VB.OptionButton optTreated 
            Caption         =   "&Treated and Discharge"
            Enabled         =   0   'False
            Height          =   285
            Left            =   270
            TabIndex        =   51
            Top             =   2640
            Width           =   2025
         End
         Begin VB.ComboBox cboPatientStatus 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "frmMedRec.frx":2FF7A
            Left            =   1530
            List            =   "frmMedRec.frx":2FF84
            Style           =   1  'Simple Combo
            TabIndex        =   41
            Top             =   390
            Width           =   855
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00808080&
            BorderStyle     =   6  'Inside Solid
            Index           =   0
            X1              =   120
            X2              =   3810
            Y1              =   900
            Y2              =   900
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
            Left            =   270
            TabIndex        =   74
            Top             =   1950
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
            Left            =   270
            TabIndex        =   72
            Top             =   1530
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
            Left            =   270
            TabIndex        =   70
            Top             =   1140
            Width           =   1185
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
            Left            =   120
            TabIndex        =   53
            Top             =   3000
            Width           =   2205
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
            Left            =   120
            TabIndex        =   50
            Top             =   2370
            Width           =   1185
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
            Left            =   300
            TabIndex        =   42
            Top             =   420
            Width           =   1185
         End
      End
      Begin VB.CommandButton cmdRemove 
         BackColor       =   &H80000003&
         Caption         =   "&Remove"
         Height          =   375
         Left            =   8730
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   6270
         Width           =   975
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000003&
         Caption         =   "&Close"
         Height          =   375
         Left            =   7740
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   6270
         Width           =   975
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000003&
         Caption         =   "&Save"
         Height          =   375
         Left            =   6780
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   6270
         Width           =   975
      End
      Begin VB.TextBox txtLastname 
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
         Left            =   -73440
         MaxLength       =   35
         TabIndex        =   18
         Top             =   480
         Width           =   3195
      End
      Begin VB.TextBox txtFirstname 
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
         Left            =   -73440
         MaxLength       =   35
         TabIndex        =   17
         Top             =   810
         Width           =   3195
      End
      Begin VB.TextBox txtMiddlename 
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
         Left            =   -73440
         MaxLength       =   35
         TabIndex        =   16
         Top             =   1140
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   15
         Top             =   1470
         Width           =   4455
      End
      Begin VB.TextBox txtAddress2 
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   14
         Top             =   1800
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
         Left            =   -73440
         MaxLength       =   35
         TabIndex        =   2
         Top             =   3210
         Width           =   3195
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
         Left            =   -73440
         MaxLength       =   50
         TabIndex        =   3
         Top             =   3540
         Width           =   4455
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
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   4
         Top             =   3870
         Width           =   2445
      End
      Begin VB.TextBox txtOccupation 
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
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2460
         Width           =   2445
      End
      Begin VB.TextBox txtNationality 
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
         Left            =   -73440
         MaxLength       =   20
         TabIndex        =   12
         Top             =   2790
         Width           =   2445
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
         Left            =   -73440
         MaxLength       =   35
         TabIndex        =   5
         Top             =   4230
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
         Left            =   -67830
         MaxLength       =   20
         TabIndex        =   11
         Top             =   510
         Width           =   645
      End
      Begin VB.TextBox txtHeight 
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
         Left            =   -67830
         MaxLength       =   20
         TabIndex        =   10
         ToolTipText     =   "ex.: 5' 11"""
         Top             =   810
         Width           =   915
      End
      Begin VB.TextBox nWeight 
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
         Left            =   -67830
         MaxLength       =   20
         TabIndex        =   9
         ToolTipText     =   "Weight in lbs"
         Top             =   1110
         Width           =   705
      End
      Begin VB.ComboBox cboSex 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMedRec.frx":2FF91
         Left            =   -67830
         List            =   "frmMedRec.frx":2FF9B
         Style           =   1  'Simple Combo
         TabIndex        =   8
         Top             =   1410
         Width           =   1185
      End
      Begin VB.ComboBox cboSocialClass 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMedRec.frx":2FFAD
         Left            =   -67830
         List            =   "frmMedRec.frx":2FFC6
         Style           =   1  'Simple Combo
         TabIndex        =   7
         Top             =   2130
         Width           =   1125
      End
      Begin VB.ComboBox cboCivilStatus 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmMedRec.frx":2FFE3
         Left            =   -67830
         List            =   "frmMedRec.frx":2FFF3
         Style           =   1  'Simple Combo
         TabIndex        =   6
         Top             =   1770
         Width           =   1425
      End
      Begin MSAdodcLib.Adodc adoPatients 
         Height          =   405
         Left            =   -74460
         Top             =   6030
         Visible         =   0   'False
         Width           =   9105
         _ExtentX        =   16060
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
         Bindings        =   "frmMedRec.frx":30018
         Height          =   1935
         Left            =   -74520
         TabIndex        =   75
         Top             =   4650
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3413
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
         Left            =   -67830
         TabIndex        =   19
         Top             =   150
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   20578305
         CurrentDate     =   38382
      End
      Begin MSMask.MaskEdBox nTelNo 
         Height          =   285
         Left            =   -73440
         TabIndex        =   20
         Top             =   2130
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
      Begin RichTextLib.RichTextBox rtbFinalDiagnosis 
         Height          =   1515
         Left            =   2130
         TabIndex        =   57
         Top             =   990
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   2672
         _Version        =   393217
         ScrollBars      =   1
         TextRTF         =   $"frmMedRec.frx":30032
      End
      Begin RichTextLib.RichTextBox rtbRemarks 
         Height          =   1515
         Left            =   2130
         TabIndex        =   58
         Top             =   2520
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   2672
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmMedRec.frx":300B4
      End
      Begin VB.Label Label30 
         Caption         =   "Remarks"
         Height          =   255
         Left            =   810
         TabIndex        =   68
         Top             =   2520
         Width           =   1305
      End
      Begin VB.Label Label29 
         Caption         =   "Allergies"
         Height          =   255
         Left            =   810
         TabIndex        =   55
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label26 
         Caption         =   "Final Findings"
         Height          =   255
         Left            =   810
         TabIndex        =   49
         Top             =   990
         Width           =   1335
      End
      Begin VB.Label Label25 
         Caption         =   "Diagnosis"
         Height          =   255
         Left            =   810
         TabIndex        =   48
         Top             =   240
         Width           =   795
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   780
         X2              =   9660
         Y1              =   4080
         Y2              =   4080
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
         Left            =   -74520
         TabIndex        =   39
         Top             =   210
         Width           =   1125
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
         Left            =   -74520
         TabIndex        =   38
         Top             =   540
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
         Left            =   -74520
         TabIndex        =   37
         Top             =   870
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
         Left            =   -74520
         TabIndex        =   36
         Top             =   1200
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
         Left            =   -74520
         TabIndex        =   35
         Top             =   1530
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
         Left            =   -74520
         TabIndex        =   34
         Top             =   2190
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
         Left            =   -68730
         TabIndex        =   33
         Top             =   210
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
         Left            =   -68730
         TabIndex        =   32
         Top             =   540
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
         Left            =   -68730
         TabIndex        =   31
         Top             =   1470
         Width           =   375
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
         Left            =   -74520
         TabIndex        =   30
         Top             =   2490
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
         Left            =   -68730
         TabIndex        =   29
         Top             =   840
         Width           =   525
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
         Left            =   -68730
         TabIndex        =   28
         Top             =   1140
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
         Left            =   -74520
         TabIndex        =   27
         Top             =   4260
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
         Left            =   -74520
         TabIndex        =   26
         Top             =   2790
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
         Left            =   -74520
         TabIndex        =   25
         Top             =   3270
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
         Left            =   -74520
         TabIndex        =   24
         Top             =   3570
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
         Left            =   -74520
         TabIndex        =   23
         Top             =   3900
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
         Left            =   -68730
         TabIndex        =   22
         Top             =   2190
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
         Left            =   -68730
         TabIndex        =   21
         Top             =   1830
         Width           =   885
      End
   End
   Begin VB.Menu mnuSearch 
      Caption         =   "&Search"
      Begin VB.Menu mnuSearchPatient 
         Caption         =   "Patient"
         Shortcut        =   {F2}
      End
   End
End
Attribute VB_Name = "frmMedRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdRemove_Click()
  If ValidateEntries Then
    If CheckHospNo Then
      ' Confirm Saving..
      If MsgBox("Pls click Ok to confirm Medical Record DELETION.", vbExclamation + vbOKCancel) = vbOK Then
        'Medical Record
        If rstMedRec.RecordCount > 0 Then
          rstMedRec.MoveFirst
          Do Until rstMedRec.EOF
            If Trim(rstMedRec("HospNo")) = Trim(nHospNo) And _
              rstMedRec("Date") = Date Then 'ok to delete
              rstMedRec.Edit
              rstMedRec.Delete
              Exit Do
            End If
            rstMedRec.MoveNext
           Loop
        End If
        
        'Clear all fields
        Init
        nHospNo = ""
        nHospNo.SetFocus
      End If
    Else
      MsgBox "Unable to update medical record of unregistered patient. Pls create patient's info first.", vbCritical
    End If
  End If

End Sub

Private Sub cmdSave_Click()
  Dim strDisposition, strConditionOfDischarge As String
  Dim lMedRecNew As Boolean
  
  If ValidateEntries Then
    If CheckHospNo Then
      ' Confirm Saving..
      If MsgBox("Pls click Ok to confirm Medical Record update.", vbInformation + vbOKCancel) = vbOK Then
        'Medical Record
        lMedRecNew = True
        If rstMedRec.RecordCount > 0 Then
          rstMedRec.MoveFirst
          Do Until rstMedRec.EOF
          If Trim(rstMedRec("HospNo")) = Trim(nHospNo) And _
            rstMedRec("Date") = Date Then 'edit
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
        rstMedRec("Lastname") = txtLastname
        rstMedRec("Firstname") = txtFirstname
        rstMedRec("Middlename") = txtMiddlename
        rstMedRec("Address") = txtAddress1
        rstMedRec("TelNo") = nTelNo
        rstMedRec("Birthdate") = dBirthDate
        rstMedRec("Age") = nAge
        rstMedRec("Nationality") = txtNationality
        rstMedRec("CivilStatus") = cboCivilStatus
        rstMedRec("Sex") = cboSex
        rstMedRec("FamilyDoctor") = txtFamilyDoctor
        rstMedRec("Diagnosis") = cboDisease
        rstMedRec("DateOfArrival") = dDateOfArrival
        rstMedRec("Date") = Date
        rstMedRec("Guardian") = UCase(txtGuardian)
        rstMedRec("GAddress") = UCase(txtGAddress)
        rstMedRec("Relationship") = UCase(txtRelationship)
        rstMedRec("Allergies") = UCase(txtAllergies)
        rstMedRec("FinalDiagnosis") = UCase(rtbFinalDiagnosis.Text)
        rstMedRec("Remarks") = UCase(rtbRemarks.Text)
        rstMedRec("AttendingPhysician") = txtAttendingPhysician
        rstMedRec("OrderDate") = dOrderDate
        rstMedRec("OrderTime") = txtOrderTime
        rstMedRec("DocOrders") = UCase(txtDocOrders)
        rstMedRec("NurseRemarks") = UCase(txtNurseRemarks)
        
        'discharge info
        strDisposition = "T"
        If optTransferred = True Then
          strDisposition = "S"
        End If
        strConditionOfDischarge = "G"
        If optSatisfactory Then
          strConditionOfDischarge = "F"
        ElseIf optSerious Then
          strConditionOfDischarge = "U"
        ElseIf optCritical Then
          strConditionOfDischarge = "C"
        End If
        rstMedRec("Disposition") = strDisposition
        rstMedRec("ConditionOfDischarge") = strConditionOfDischarge
          
        rstMedRec.Update
        
        'Clear all fields
        Init
        nHospNo = ""
        nHospNo.SetFocus
      End If
    Else
      MsgBox "Unable to update medical record of unregistered patient. Pls create patient's info first.", vbCritical
    End If
  End If
End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
  'nHospNo.Mask = ""
  nHospNo = Str(adoPatients.Recordset("HospNo"))
  nHospNo_LostFocus
End Sub

Private Sub Form_Activate()
  nHospNo.SetFocus
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstPatients = dbsInfoSys.OpenRecordset("Patients")
  Set rstMedRec = dbsInfoSys.OpenRecordset("MedRec")
  Set rstDisease = dbsInfoSys.OpenRecordset("Disease")
  Set rstAdmitStatus = dbsInfoSys.OpenRecordset("AdmitStatus")
  rstPatients.Index = "HospNo"
  rstMedRec.Index = "HospNo"
  rstAdmitStatus.Index = "HospNo"

  Init 'Clear all entry fields
  
  GetDisease
  
  SStab1.Tab = 0
  cmdRemove.Visible = nLevel > 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstAdmitStatus.Close
  rstDisease.Close
  rstMedRec.Close
  rstPatients.Close
  dbsInfoSys.Close
End Sub

'- clear all entry fields
Function Init()
  SStab1.Tab = 0
  cmdSave.Caption = "Save"
  cmdSave.Enabled = True
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
  cboSex = ""
  cboCivilStatus = ""
  cboSocialClass = ""
  txtGuardian = ""
  txtGAddress = ""
  txtRelationship = ""
  txtFamilyDoctor = ""
  cboPatientStatus = ""
  dDateOfArrival = ""
  cboRoomNo = ""
  cboRoomType = ""
  txtAllergies = ""
  rtbFinalDiagnosis = ""
  rtbRemarks = ""
  txtAttendingPhysician = ""
  txtOrderTime.Mask = ""
  txtOrderTime = ""
  txtOrderTime.Mask = "##:##"
  txtDocOrders = ""
  txtNurseRemarks = ""
End Function

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

Function ValidateEntries() As Boolean
  If Val(nHospNo) < 1 Then
    ValidateEntries = False
    MsgBox "Invalid Hospital Number", vbCritical
    nHospNo.SetFocus
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

Private Sub mnuSearchPatient_Click()
  frmSearchPatient.Show 1
  nHospNo = Str(nPatNo)
  nHospNo_LostFocus
  txtGuardian.SetFocus
End Sub

Private Sub nHospNo_LostFocus()
  cmdSave.Caption = "Save"
  If CheckHospNo Then
    'cmdSave.Caption = "Edit"
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
    
    GetDataFromAdmitStatus
    GetDataFromMedRec
    
  Else
    Init
  End If
End Sub

Function CheckAdmitStatus() As Boolean
  CheckAdmitStatus = False
  rstAdmitStatus.Seek "=", Val(nHospNo)
  If Not rstAdmitStatus.NoMatch Then
    CheckAdmitStatus = True
  End If
End Function

Function GetDataFromAdmitStatus()
  If CheckAdmitStatus Then
    'If Trim(rstAdmitStatus("DateDischarge")) <> "" Then 'not yet discharge
      cmdSave.Caption = "Edit"
      cmdSave.Enabled = nLevel > 1
      dDateOfArrival = rstAdmitStatus("Date")
      cboDisease = rstAdmitStatus("Disease")
      cboRoomNo = rstAdmitStatus("RoomNo")
      cboRoomType = rstAdmitStatus("RoomType")
      'txtAttendingPhysician = rstAdmitStatus("AttendingPhysician")
    'End If
  End If
End Function

Function GetDataFromMedRec()
  If rstMedRec.RecordCount > 0 Then
    rstMedRec.MoveFirst
    Do Until rstMedRec.EOF
      If Trim(rstMedRec("HospNo")) = Trim(nHospNo) And _
         rstMedRec("Date") = Date And _
         rstMedRec("Discharge") = False Then 'edit
         
        cmdSave.Caption = "Edit"
        'cboDisease = rstMedRec("Disease")
        txtAllergies = rstMedRec("Allergies")
        rtbFinalDiagnosis = rstMedRec("FinalDiagnosis")
        rtbRemarks = rstMedRec("Remarks")
        txtAttendingPhysician = rstMedRec("AttendingPhysician")
        'Disposition
        optTreated = (rstMedRec("Disposition") = "T")
        optTransferred = (rstMedRec("Disposition") = "S")
        'Condition of Discharged
        optGood = (rstMedRec("ConditionOfDischarge") = "G")
        optSatisfactory = (rstMedRec("ConditionOfDischarge") = "F")
        optSerious = (rstMedRec("ConditionOfDischarge") = "U")
        optCritical = (rstMedRec("ConditionOfDischarge") = "C")
        
        dOrderDate = rstMedRec("OrderDate")
        txtOrderTime = rstMedRec("OrderTime")
        txtDocOrders = rstMedRec("DocOrders")
        txtNurseRemarks = rstMedRec("NurseRemarks")
        'Exit Do
      End If
      rstMedRec.MoveNext
    Loop
  End If
End Function

Private Sub SSTab1_Click(PreviousTab As Integer)
  If PreviousTab = 1 Then
    'nHospNo.SetFocus
  End If
End Sub

'-- UDF
Private Sub nHospNo_GotFocus()
  Call FocusMe(nHospNo)
End Sub

Private Sub txtAllergies_GotFocus()
  Call FocusMe(txtAllergies)
End Sub

Private Sub txtAttendingPhysician_Gotfocus()
  Call FocusMe(txtAttendingPhysician)
End Sub

Private Sub txtDocOrders_GotFocus()
  Call FocusMe(txtDocOrders)
End Sub

Private Sub txtFamilyDoctor_GotFocus()
  Call FocusMe(txtFamilyDoctor)
End Sub

Private Sub txtGAddress_GotFocus()
  Call FocusMe(txtGAddress)
End Sub

Private Sub txtGuardian_GotFocus()
  Call FocusMe(txtGuardian)
End Sub

Private Sub txtNurseRemarks_GotFocus()
  Call FocusMe(txtNurseRemarks)
End Sub

Private Sub txtOrderTime_GotFocus()
  Call FocusMe(txtOrderTime)
End Sub

Private Sub txtOrderTime_LostFocus()
  txtOrderTime = ValidTime(txtOrderTime)
End Sub

Private Sub txtRelationship_GotFocus()
  Call FocusMe(txtRelationship)
End Sub

Private Sub cboDisease_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtAllergies_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtAttendingPhysician_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtDocOrders_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtOrderTime_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtNurseRemarks_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub cboPatientStatus_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub nHospNo_KeyPress(KeyAscii As Integer)
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

Private Sub txtGAddress_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub

Private Sub txtRelationship_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'-- eo: UDF

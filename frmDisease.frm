VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmDisease 
   Caption         =   "Diseases table maintenance"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4020
   Icon            =   "frmDisease.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmDisease.frx":0E42
   ScaleHeight     =   5610
   ScaleWidth      =   4020
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmDisease.frx":5D48
      Height          =   3225
      Left            =   270
      TabIndex        =   5
      Top             =   1680
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   5689
      _Version        =   393216
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
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
   Begin MSAdodcLib.Adodc adoDisease 
      Height          =   345
      Left            =   510
      Top             =   3900
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   609
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
      RecordSource    =   "Disease"
      Caption         =   "Disease"
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
   Begin VB.CommandButton cmdRemove 
      BackColor       =   &H80000003&
      Caption         =   "Remove"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H80000003&
      Caption         =   "Close"
      Height          =   375
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5190
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      BackColor       =   &H80000003&
      Caption         =   "Save"
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5190
      Width           =   1095
   End
   Begin VB.TextBox txtDisease 
      Height          =   345
      Left            =   1200
      MaxLength       =   25
      TabIndex        =   1
      Top             =   990
      Width           =   2325
   End
   Begin VB.Label Disease 
      BackStyle       =   0  'Transparent
      Caption         =   "Disease"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   270
      TabIndex        =   0
      Top             =   1020
      Width           =   945
   End
End
Attribute VB_Name = "frmDisease"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Function CheckDisease(strDisease As String) As Boolean
  CheckDisease = False
  rstDisease.Seek "=", txtDisease
  If Not rstDisease.NoMatch Then
    CheckDisease = True
    cmdSave.Caption = "Edit"
  End If
End Function

Private Sub cmdRemove_Click()
  If Not IsEmpty(txtDisease) Then
    If CheckDisease(txtDisease) Then
      If MsgBox("Click Ok to Remove this record.", vbExclamation + vbOKCancel) = vbOK Then
        rstDisease.Edit
        rstDisease.Delete
        adoDisease.Refresh
        cmdSave.Caption = "Save"
        txtDisease = ""
        txtDisease.SetFocus
      End If
    Else
      MsgBox txtDisease & " not found ", vbCritical
    End If
  Else
    MsgBox "Nothing to remove.", vbInformation
  End If
End Sub

Private Sub cmdSave_Click()
  If Not IsEmpty(txtDisease) Then
    ' Confirm Saving..
    If MsgBox("Click Ok to Save", vbInformation + vbOKCancel) = vbOK Then
      If Not CheckDisease(txtDisease) Then
        rstDisease.AddNew 'New record
      Else
        rstDisease.Edit 'Existing
      End If
      rstDisease("Disease") = UCase(txtDisease)
      rstDisease.Update
      adoDisease.Refresh
      
      cmdSave.Caption = "Save"
      txtDisease = ""
      txtDisease.SetFocus
    End If
  Else
    MsgBox "Nothing to save", vbInformation
  End If
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstDisease = dbsInfoSys.OpenRecordset("Disease")
  rstDisease.Index = "Disease"
  cmdRemove.Visible = nLevel > 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
  rstDisease.Close
  dbsInfoSys.Close
End Sub

'- UDF
Private Sub txtDisease_GotFocus()
  Call FocusMe(txtDisease)
End Sub

Private Sub txtDisease_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then SendKeys Chr(9) 'Tab emulation
End Sub
'- eo: UDF


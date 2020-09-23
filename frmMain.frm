VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmMain 
   Caption         =   "Patient's Information System"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   882.558
   ScaleMode       =   0  'User
   ScaleWidth      =   1202.559
   StartUpPosition =   1  'CenterOwner
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   615
      Left            =   1080
      TabIndex        =   12
      Top             =   6840
      Width           =   5235
      _cx             =   9234
      _cy             =   1085
      FlashVars       =   ""
      Movie           =   ""
      Src             =   ""
      WMode           =   "Window"
      Play            =   -1  'True
      Loop            =   -1  'True
      Quality         =   "High"
      SAlign          =   ""
      Menu            =   -1  'True
      Base            =   ""
      AllowScriptAccess=   "always"
      Scale           =   "NoBorder"
      DeviceFont      =   0   'False
      EmbedMovie      =   0   'False
      BGColor         =   ""
      SWRemote        =   ""
      MovieData       =   ""
      SeamlessTabbing =   -1  'True
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Discharge"
      Height          =   1395
      Left            =   8190
      Picture         =   "frmMain.frx":2B9F8
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Click this button to load sub-form3"
      Top             =   2640
      Width           =   2085
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Admit"
      Height          =   1425
      Left            =   5670
      Picture         =   "frmMain.frx":2D574
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Click this button to load sub-form2"
      Top             =   2640
      Width           =   2085
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00DBAD8E&
      Caption         =   "&Patient"
      Height          =   1425
      Left            =   3180
      Picture         =   "frmMain.frx":2F6E2
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Click this button to load sub-form1"
      Top             =   2640
      Width           =   2085
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H80000003&
      Caption         =   "Exit"
      Height          =   435
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Quit to Windows"
      Top             =   7740
      Width           =   1305
   End
   Begin VB.CommandButton cmdLogoff 
      BackColor       =   &H80000003&
      Caption         =   "Log-Off"
      Enabled         =   0   'False
      Height          =   435
      Left            =   8250
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Return to Welcome Screen"
      Top             =   7740
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.PictureBox picSetup 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   420
      Picture         =   "frmMain.frx":30B7C
      ScaleHeight     =   1635
      ScaleWidth      =   1845
      TabIndex        =   2
      ToolTipText     =   "Setup"
      Top             =   1170
      Width           =   1845
   End
   Begin VB.PictureBox picMaintenance 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   420
      Picture         =   "frmMain.frx":32F8F
      ScaleHeight     =   1635
      ScaleWidth      =   1815
      TabIndex        =   1
      ToolTipText     =   "Maintenance"
      Top             =   4830
      Width           =   1815
   End
   Begin VB.PictureBox picReport 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   420
      Picture         =   "frmMain.frx":3406B
      ScaleHeight     =   1635
      ScaleWidth      =   1845
      TabIndex        =   0
      ToolTipText     =   "Reports"
      Top             =   2970
      Width           =   1845
   End
   Begin VB.Image imgSearch 
      Height          =   765
      Left            =   7440
      ToolTipText     =   "Search patient"
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Image imgMedRec 
      Height          =   765
      Left            =   8610
      ToolTipText     =   "Update medical records"
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Image imgAuthors 
      Height          =   765
      Left            =   9780
      ToolTipText     =   "About the authors"
      Top             =   6720
      Width           =   1035
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Patient's Information System v1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   11
      Top             =   7680
      Width           =   4395
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2005 by PC Land Computers and Cellphones for Region I Medical Center"
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   10
      Top             =   7980
      Width           =   6465
   End
   Begin VB.Label lblCurrentMenu 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   330
      TabIndex        =   9
      Top             =   330
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Label lblMenuTips 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Reference Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2940
      TabIndex        =   5
      Top             =   4980
      Width           =   7545
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------
'Purpose.....: Main Form
'Programmer..: Jeffrey Lim
'Email.......: jil821@yahoo.com
'Cell#.......: 0921-825-9455
'Url.........: http://www.pcland.cjb.net
'--------------------------------------------

Private Sub cmdExit_Click()
  End
End Sub

'Private Sub cmdLogOff_Click()
'  Unload Me
'  frmLogin.Show
'End Sub

Function GetMedSurvey()
  Dim nTreated, nAdmitted, nTransferred As Long
  Dim strPrevDisease As String
  
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
  Set rstTemp = dbsInfoSys.OpenRecordset("Temp")
  rstDiagnos.Index = "Disease"
  rstTemp.Index = "Disease"
  
  If rstTemp.RecordCount > 0 Then
    rstTemp.MoveFirst
    Do Until rstTemp.EOF
      rstTemp.Delete
      rstTemp.MoveNext
    Loop
  End If
  
  nTreated = 0
  nTransferred = 0
  nAdmitted = 0
  If rstDiagnos.RecordCount > 0 Then
    rstDiagnos.MoveFirst
    strPrevDisease = Trim(rstDiagnos("Disease"))
    Do Until rstDiagnos.EOF
      If Trim(UCase(rstDiagnos("Disease"))) = Trim(UCase(strPrevDisease)) Then
        If rstDiagnos("Disposition") = "T" Then
          nTreated = nTreated + 1
        ElseIf rstDiagnos("Disposition") = "S" Then
          nTransferred = nTransferred + 1
        Else 'A-Still Recovering
          nAdmitted = nAdmitted + 1
        End If
        
        rstTemp.Seek "=", strPrevDisease
        If Not rstTemp.NoMatch Then
          rstTemp.Edit
        Else
          rstTemp.AddNew
        End If
        rstTemp("Disease") = strPrevDisease
        rstTemp("Treated") = nTreated
        rstTemp("Admitted") = nAdmitted
        rstTemp("Transferred") = nTransferred
        rstTemp.Update
      
        rstDiagnos.MoveNext
      Else 'New disease
        nTreated = 0
        nTransferred = 0
        nAdmitted = 0
        strPrevDisease = rstDiagnos("Disease")
      End If
    Loop
  End If
End Function

Private Sub Command1_Click()
  Select Case Command1.Caption
    Case "&Patient"
      frmPatients.Show 1
    Case "&Medical Survey"
      GetMedSurvey
      rptMedSurvey.DataMember = "cmdMedSurvey"
      rptMedSurvey.Show 1
      Unload rptMedSurvey
      Unload denvMedSurvey
    Case Else '&Users
      frmUsers.Show 1
  End Select
End Sub

Private Sub Command2_Click()
  Select Case Command2.Caption
    Case "&Admit"
      frmAdmit.Show 1
    Case "&Medical Record"
      frmMedRecRep.Show 1
    Case Else '&Diseases
      frmDisease.Show 1
  End Select
End Sub

Private Sub Command3_Click()
  Select Case Command3.Caption
    Case "&Discharge"
      frmDischarge.Show 1
    Case "&Clear Tables"
      If nLevel > 2 Then
        frmInitialize.Show 1
      Else
        MsgBox "Only the administrator can access this menu.", vbCritical
      End If
    Case Else 'Screen Survey
      frmScreenSurvey.Show 1
  End Select
  
End Sub

Private Sub Form_Load()
  ShockwaveFlash1.Movie = App.Path & "\region1.swf"
  
  Set dbsInfoSys = OpenDatabase(App.Path & "\InfoSys.mdb")
  Set rstDiagnos = dbsInfoSys.OpenRecordset("Diagnos")
  Set rstTemp = dbsInfoSys.OpenRecordset("Temp")
  rstDiagnos.Index = "Disease"
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'rstDiagnos.Close
  'rstTemp.Close
  'dbsInfoSys.Close
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMenuTips.Caption = "                               Medical Center" & Chr(13) & Chr(13) & "                       Patient's Information System"
  picSetup.Picture = LoadPicture(App.Path & "\SetFlat.jpg")
  picReport.Picture = LoadPicture(App.Path & "\RepFlat.jpg")
  picMaintenance.Picture = LoadPicture(App.Path & "\MaintFlat.jpg")
  
  '- Lower buttons
  imgSearch.Picture = LoadPicture(App.Path & "\SearchFlat.jpg")
  imgMedRec.Picture = LoadPicture(App.Path & "\MedRecFlat.jpg")
  imgAuthors.Picture = LoadPicture(App.Path & "\AuthorsFlat.jpg")
End Sub



'- Main Menu
Private Sub picSetup_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  picSetup.Picture = LoadPicture(App.Path & "\SetDown.jpg")
  Command1.Picture = LoadPicture(App.Path & "\Patient.jpg")
  Command2.Picture = LoadPicture(App.Path & "\Admit.jpg")
  Command3.Picture = LoadPicture(App.Path & "\Dischrg.jpg")
  Command1.Caption = "&Patient"
  Command2.Caption = "&Admit"
  Command3.Caption = "&Discharge"
End Sub

Private Sub picReport_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  picReport.Picture = LoadPicture(App.Path & "\RepDown.jpg")
  Command1.Picture = LoadPicture(App.Path & "\MedSur.jpg")
  Command2.Picture = LoadPicture(App.Path & "\MedRec.jpg")
  Command3.Picture = LoadPicture(App.Path & "\ScrSur.jpg")
  Command1.Caption = "&Medical Survey"
  Command2.Caption = "&Medical Record"
  Command3.Caption = "&Screen Survey"
End Sub

Private Sub picMaintenance_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  picMaintenance.Picture = LoadPicture(App.Path & "\MaintDown.jpg")
  Command1.Picture = LoadPicture(App.Path & "\Users.jpg")
  Command2.Picture = LoadPicture(App.Path & "\Disease.jpg")
  Command3.Picture = LoadPicture(App.Path & "\ClrTab.jpg")
  Command1.Caption = "&Users"
  Command2.Caption = "&Diseases"
  Command3.Caption = "&Clear Tables"
  
  'Command3.Enabled = nLevel > 2
End Sub

Private Sub picMaintenance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMenuTips.Caption = "Maintenance: Click this button if you want to create users, update table of diseases or clear all tables for initial system installation."
  picMaintenance.Picture = LoadPicture(App.Path & "\MaintUp.jpg")
End Sub

Private Sub picReport_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMenuTips.Caption = "Reports: Click this button if you want to generate or print a medical survey report or patient's medical records."
  picReport.Picture = LoadPicture(App.Path & "\RepUp.jpg")
End Sub

Private Sub picSetup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  lblMenuTips.Caption = "Setup: Click this button if you want to register patient, confine or discharge previously admitted patient."
  picSetup.Picture = LoadPicture(App.Path & "\SetUp.jpg")
End Sub
'- eo: Main Menu Buttons

'- Lower Buttons
Private Sub imgSearch_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgSearch.Picture = LoadPicture(App.Path & "\SearchDown.jpg")
  frmSearchPatient.Show 1
End Sub

Private Sub imgMedRec_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMedRec.Picture = LoadPicture(App.Path & "\MedRecDown.jpg")
  frmMedRec.Show 1
End Sub

Private Sub imgAuthors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgAuthors.Picture = LoadPicture(App.Path & "\AuthorsDown.jpg")
  frmAuthors.Show 1
End Sub

Private Sub imgSearch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgSearch.Picture = LoadPicture(App.Path & "\SearchUp.jpg")
End Sub

Private Sub imgMedRec_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgMedRec.Picture = LoadPicture(App.Path & "\MedRecUp.jpg")
End Sub

Private Sub imgAuthors_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  imgAuthors.Picture = LoadPicture(App.Path & "\AuthorsUp.jpg")
End Sub
'- eo: Lower Buttons

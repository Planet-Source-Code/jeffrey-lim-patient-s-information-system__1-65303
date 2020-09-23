VERSION 5.00
Object = "{D27CDB6B-AE6D-11CF-96B8-444553540000}#1.0#0"; "Flash.ocx"
Begin VB.Form frmSplash 
   BackColor       =   &H00DBAD8E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4905
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1185
      Left            =   0
      TabIndex        =   1
      Top             =   3750
      Width           =   7380
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Author: Jeffrey Lim"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   255
         Left            =   4050
         TabIndex        =   4
         Top             =   870
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   60
         Picture         =   "frmSplash.frx":0000
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright © 2005 by PC Land Computers and Cellphones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   420
         Width           =   6465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Patient's Information System v1.0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   810
         TabIndex        =   2
         Top             =   90
         Width           =   4395
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   6510
      Top             =   2580
   End
   Begin ShockwaveFlashObjectsCtl.ShockwaveFlash ShockwaveFlash1 
      Height          =   3765
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      _cx             =   12832
      _cy             =   6641
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
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
  Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  frmLogin.Show
End Sub

Private Sub Timer1_Timer()
  Unload Me
End Sub

Private Sub Form_Load()
  If App.PrevInstance = True Then
    MsgBox "System is Already in Run Mode ...", vbCritical, "System is Already Running ..."
    Unload Me
    Exit Sub
  End If
  ChDir (App.Path)
  ShockwaveFlash1.Movie = App.Path & "\splash.swf"
  
End Sub


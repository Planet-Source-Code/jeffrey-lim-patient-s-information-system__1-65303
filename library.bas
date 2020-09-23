Attribute VB_Name = "modLibrary"
'--------------------------------------------
'Purpose.....: Public Variables and UDF
'Programmer..: Jeffrey Lim
'Email.......: jil821@yahoo.com
'Cell#.......: 0921-825-9455
'Url.........: http://www.pcland.cjb.net
'--------------------------------------------

Option Explicit

'GLOBAL VARIABLES-->
Global dbsInfoSys As Database
Global nLevel As Integer
Global nPatNo As Long
Public rstAdmitStatus, rstDiagnos, rstDisease, rstLastNo As Recordset
Public rstMedRec, rstPatients, rstUsers, rstTemp As Recordset

Public Function GetLastNo() As String
  Set rstLastNo = dbsInfoSys.OpenRecordset("LastNo")
  'rstLastNo.MoveFirst
  GetLastNo = Str(rstLastNo("HospNo") + 1)
  rstLastNo.Close
End Function

Public Function UpdateLastNo(nNo As Long)
  Set rstLastNo = dbsInfoSys.OpenRecordset("LastNo")
  If nNo > rstLastNo("HospNo") Then
    rstLastNo.Edit
    rstLastNo("HospNo") = nNo
    rstLastNo.Update
  End If
  rstLastNo.Close
End Function

Public Function GetAge(datEmpDateOfBirth As Variant) As Integer
    GetAge = Int(DateDiff("y", CDate(datEmpDateOfBirth), Date) / 365.25)
End Function

Public Function ValidTime(strTime) As String
  Dim strHr, strMin As String
  
  strHr = Left(strTime, InStr(strTime, ":") - 1)
  strMin = Mid(strTime, InStr(strTime, ":") + 1)
  
  'Validate Hour
  If Val(strHr) < 1 Then strHr = "00"
  If Val(strHr) > 23 Then
    strHr = "23" 'default
    strMin = "59"
  End If
  'Validate Min
  If Val(strMin) < 1 Then strMin = "00"
  If Val(strMin) > 59 Then strMin = "59" 'default
  
ValidTime = strHr & ":" & strMin
End Function

'-- User Defined Functions
Function FocusMe(ctlName As Control) 'Automatically select text/value
  With ctlName
    .SelStart = 0
    .SelLength = Len(ctlName)
  End With
End Function
'-- eo: UDF

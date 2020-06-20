Attribute VB_Name = "Module1"
Option Explicit

Public strCon226 As String
Public strCon230 As String
Public strConSAP As String
Public strCon227 As String

Public ConLocal As ADODB.Connection
Public ConXls As ADODB.Connection
Public Con227 As ADODB.Connection
Public ConSAP As ADODB.Connection
Public ConMDB As ADODB.Connection
Public Con226 As ADODB.Connection
Public FlagErr As Integer
Public Fx As Integer


Public Sub ConnectionDB()

On Error GoTo ErrMsg
FlagErr = 0


Set ConLocal = New ADODB.Connection
ConLocal.ConnectionString = strCon230 '"Provider=MSDASQL.1;Persist Security Info=False;Data Source=TE;Initial Catalog=minimal_internal"
ConLocal.Open
frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CON230  : CONNECTED " & vbCrLf

'--untuk server, langsung arahin ke 226, soalnya hanya sendiri, cek di ODBC BA
'--untuk store, ke 227 ambil dari db offline di ODBC NE

Set Con226 = New ADODB.Connection
Con226.ConnectionString = strCon226 ' "Provider=MSDASQL.1;Persist Security Info=False;Data Source=BA;Initial Catalog=minimal-test"
Con226.Open
frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CON226  : CONNECTED " & vbCrLf


Set Con227 = New ADODB.Connection
Con227.ConnectionString = strCon227 '"Provider=MSDASQL.1;Persist Security Info=False;Data Source=NE"
Con227.Open
frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CON227  : CONNECTED " & vbCrLf

'SAP

Set ConSAP = New ADODB.Connection
ConSAP.ConnectionString = strConSAP '"Provider=MSDASQL.1;Password=minimalUOB23;Persist Security Info=True;User ID=sa;Data Source=uob"
ConSAP.Open
frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CONSAP  : CONNECTED " & vbCrLf



'Set ConMDB = New ADODB.Connection
'ConMDB.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\offline.mdb;User ID=Admin;Persist Security Info=False;JET OLEDB:Database Password=minimalUOB23"
'
'ConMDB.Open

frmGenData.txtLog.Text = frmGenData.txtLog.Text & "CONMDB  : CONNECTED " & vbCrLf
Exit Sub

ErrMsg:
    FlagErr = 1



End Sub




Attribute VB_Name = "Module1"
'===========================================================
'===========================================================
'===========================================================
'====   Programmed by:      Macku
'====   Email:              mackusoft@yahoo.com
'====   Web:                http://www.mackusoft.tk
'====   Credits:            Philip V. Naparan and
'====                       other pscode contributor
'====   Note:               Send me feedback regarding
'====                       this application
'====
'====   MABUHAY MY FELLOW FILIPINOS! ITO'Y GAWANG PINOY!!!
'====
'===========================================================
'===========================================================
'===========================================================

Option Explicit

Public strFind As String
Public dbconn As New ADODB.Connection

Sub main()
dbconn.Open "Provider=MSDASQL.1;Data Source=Maindata"
frmSplash.Show
End Sub

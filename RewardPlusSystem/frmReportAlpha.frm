VERSION 5.00
Begin VB.Form frmReportAlpha 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mackusoft"
   ClientHeight    =   2745
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportAlpha.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.ComboBox cboTo 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1560
      Width           =   4095
   End
   Begin VB.ComboBox cboFrom 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   4095
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4200
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "To"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "From"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "frmReportAlpha.frx":058A
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -240
      Picture         =   "frmReportAlpha.frx":0E54
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmReportAlpha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private rs As New ADODB.Recordset

Private Sub loadAlpha()
With cboFrom
    .clear
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
    .AddItem "F"
    .AddItem "G"
    .AddItem "H"
    .AddItem "I"
    .AddItem "J"
    .AddItem "K"
    .AddItem "L"
    .AddItem "M"
    .AddItem "N"
    .AddItem "O"
    .AddItem "P"
    .AddItem "Q"
    .AddItem "R"
    .AddItem "S"
    .AddItem "T"
    .AddItem "U"
    .AddItem "V"
    .AddItem "W"
    .AddItem "X"
    .AddItem "Y"
    .AddItem "Z"
End With

With cboTo
    .clear
    .AddItem "A"
    .AddItem "B"
    .AddItem "C"
    .AddItem "D"
    .AddItem "E"
    .AddItem "F"
    .AddItem "G"
    .AddItem "H"
    .AddItem "I"
    .AddItem "J"
    .AddItem "K"
    .AddItem "L"
    .AddItem "M"
    .AddItem "N"
    .AddItem "O"
    .AddItem "P"
    .AddItem "Q"
    .AddItem "R"
    .AddItem "S"
    .AddItem "T"
    .AddItem "U"
    .AddItem "V"
    .AddItem "W"
    .AddItem "X"
    .AddItem "Y"
    .AddItem "Z"
End With

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
rs.Open "select * from custinfo where left(name, 1) between '" & cboFrom & "' and '" & cboTo & "'", dbconn, 1, 3

If rs.RecordCount = 0 Then
    MsgBox "No record found on query.", vbCritical, "Mackusoft"
    cboFrom.SetFocus
Else
    Set DTRCustInfo.DataSource = rs
    Unload Me
    DTRCustInfo.Show
End If

Set rs = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        cmdCancel.Value = True
End Select
End Sub

Private Sub Form_Load()
Call loadAlpha
MDIFrmMain.StatusBar1.Panels(2).Text = "Reports..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

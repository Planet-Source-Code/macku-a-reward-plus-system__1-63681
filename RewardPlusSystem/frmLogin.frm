VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login: Mackusoft"
   ClientHeight    =   1725
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3960
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1019.187
   ScaleMode       =   0  'User
   ScaleWidth      =   3718.226
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   322
      Left            =   1410
      TabIndex        =   1
      Top             =   255
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   615
      TabIndex        =   4
      Top             =   1140
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2220
      TabIndex        =   5
      Top             =   1140
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   322
      IMEMode         =   3  'DISABLE
      Left            =   1410
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   645
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   0
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Index           =   1
      Left            =   225
      TabIndex        =   2
      Top             =   660
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
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
Private Sub cmdCancel_Click()
End
End Sub

Private Sub cmdOK_Click()
rs.Open "select * from userpass where username = '" & txtUserName & "' and password = '" & txtPassword & "'", dbconn, 1, 3

If rs.RecordCount = 0 Then
    MsgBox "Invalid username or password.", vbCritical, "Mackusoft"
    txtUserName.SetFocus
    SendKeys "{home}+{end}"
Else
    If rs!UserClass = "USER" Then
        MDIFrmMain.mnuCategories.Enabled = False
        MDIFrmMain.mnuRewards.Enabled = False
        MDIFrmMain.mnuUserPass.Enabled = False
    End If
    
    MDIFrmMain.StatusBar1.Panels(4).Text = txtUserName
    
    Unload Me
    MDIFrmMain.Show
End If

Set rs = Nothing
        
End Sub

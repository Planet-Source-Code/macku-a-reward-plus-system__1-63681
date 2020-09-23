VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUserPassSettings 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Username and Password Settings"
   ClientHeight    =   3840
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7770
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUserPassSettings.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   375
      Left            =   5040
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtConfirmPass 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   2760
      Width           =   3375
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2160
      Width           =   3375
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   3375
   End
   Begin VB.ComboBox cboUserClass 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3375
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   2330
      Left            =   3600
      TabIndex        =   8
      Top             =   720
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   4101
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "IL1"
      SmallIcons      =   "IL1"
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Username"
         Object.Width           =   7215
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   3600
      Top             =   2400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":19A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmUserPassSettings.frx":1D42
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7800
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Confirm Password"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Password"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Name"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Classification"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1575
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmUserPassSettings.frx":22DC
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -120
      Picture         =   "frmUserPassSettings.frx":2BA6
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmUserPassSettings"
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

Private Sub ShowUser()
On Error Resume Next
rs.Open "select * from userpass where username = '" & LV1.SelectedItem & "'", dbconn, 1, 3

If Not rs.RecordCount = 0 Then
    cboUserClass.Text = rs!UserClass
    txtUserName.Text = rs!UserName
End If

Set rs = Nothing
End Sub

Private Sub clear()
Call UserClass
txtUserName.Text = ""
txtPassword.Text = ""
txtConfirmPass.Text = ""
End Sub

Private Sub UserClass()
With cboUserClass
    .clear
    .AddItem "ADMINISTRATOR"
    .AddItem "USER"
End With
End Sub

Private Sub LoadUsers()
Me.MousePointer = vbHourglass
rs.Open "UserPass order by username", dbconn, 1, 3
Call WriteListView(LV1, rs, 7, 1)
Set rs = Nothing
Me.MousePointer = vbDefault
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim ans As String

rs.Open "select * from userpass where username = '" & LV1.SelectedItem & "'", dbconn, 1, 3

If Not rs.RecordCount = 0 Then
    ans = MsgBox("Delete username " & UCase(LV1.SelectedItem) & ". Are you sure?", vbYesNo + vbQuestion, "Mackusoft")
    
    If ans = vbYes Then
        rs.Delete
        
        Set rs = Nothing
        Call LoadUsers
        Call clear
    End If
End If

Set rs = Nothing
End Sub

Private Sub cmdSave_Click()
If cboUserClass.Text = "" Then
    MsgBox "User classification not found.", vbCritical, "Mackusoft"
    cboUserClass.SetFocus
    Exit Sub
End If

If txtUserName.Text = "" Then
    MsgBox "Username not found.", vbCritical, "Mackusoft"
    txtUserName.SetFocus
    Exit Sub
End If

If txtPassword.Text = "" Then
    MsgBox "Password not found.", vbInformation, "Mackusoft"
    txtPassword.SetFocus
    Exit Sub
End If

If Not txtConfirmPass.Text = txtPassword.Text Then
    MsgBox "Password value should be the same.", vbInformation, "Mackusoft"
    txtConfirmPass.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

rs.Open "select * from userpass where username = '" & txtUserName & "'", dbconn, 1, 3

If rs.RecordCount = 0 Then
    With rs
        .AddNew
        !UserClass = cboUserClass
        !UserName = txtUserName
        !Password = txtPassword
        .Update
    End With
Else
    With rs
        !UserClass = cboUserClass
        !UserName = txtUserName
        !Password = txtPassword
        .Update
    End With
End If

Set rs = Nothing
Call clear
Call LoadUsers
cboUserClass.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        cmdClose.Value = True
End Select
End Sub

Private Sub Form_Load()
Call LoadUsers
Call UserClass
MDIFrmMain.StatusBar1.Panels(2).Text = "Username and Password Settings..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub LV1_Click()
Call ShowUser
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
        cmdDelete.Value = True
End Select
End Sub

Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
Call ShowUser
End Sub

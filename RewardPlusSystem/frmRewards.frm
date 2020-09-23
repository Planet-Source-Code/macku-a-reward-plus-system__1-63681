VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRewards 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rewards"
   ClientHeight    =   7110
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRewards.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPointsToEarn 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtReward 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   5775
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   10186
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "IL1"
      SmallIcons      =   "IL1"
      ForeColor       =   16576
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reward"
         Object.Width           =   5380
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Points to Earn"
         Object.Width           =   5380
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   120
      Top             =   6240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRewards.frx":19A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Points to Earn"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reward"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "frmRewards.frx":1D42
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -240
      Picture         =   "frmRewards.frx":260C
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmRewards"
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

Private Sub loadRewards()
Me.MousePointer = vbHourglass
rs.Open "tblrewards order by reward", dbconn, 1, 3
Call WriteListView(LV1, rs, 1, 2)
Set rs = Nothing
Me.MousePointer = vbDefault
End Sub

Private Sub clear()
txtReward.Text = ""
txtPointsToEarn.Text = ""
End Sub

Private Sub SaveNow()
rs.Open "select * from tblrewards where reward = '" & txtReward & "'", dbconn, 1, 3

If rs.RecordCount = 0 Then
    With rs
        .AddNew
        !reward = UCase(txtReward)
        !pointstoearn = txtPointsToEarn
        .Update
    End With
Else
    With rs
        !reward = UCase(txtReward)
        !pointstoearn = txtPointsToEarn
        .Update
    End With
End If

Set rs = Nothing
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub txtPointToEarn_Change()

End Sub

Private Sub Form_Load()
Call loadRewards
MDIFrmMain.StatusBar1.Panels(2).Text = "Rewards..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub LV1_Click()
Call FillUp
End Sub

Private Sub FillUp()
On Error Resume Next
txtReward.Text = LV1.SelectedItem
txtPointsToEarn.Text = LV1.SelectedItem.SubItems(1)
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Call DelRec
End Sub

Private Sub DelRec()
rs.Open "select * from tblrewards where reward = '" & LV1.SelectedItem & "'", dbconn, 1, 3

If Not rs.RecordCount = 0 Then
    rs.Delete
    txtReward.Text = ""
    txtPointsToEarn.Text = ""
End If

Set rs = Nothing

Call loadRewards
End Sub

Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
Call FillUp
End Sub

Private Sub txtPointsToEarn_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtReward.Text = "" Then
        MsgBox "Reward should not be empty.", vbCritical, "Mackusoft"
        txtReward.SetFocus
        Exit Sub
    End If
    
    If Not IsNumeric(txtPointsToEarn) Then
        MsgBox "Value is invalid.", vbCritical, "Mackusoft"
        txtPointsToEarn.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    Call SaveNow
    Call clear
    txtReward.SetFocus
    
    Call loadRewards
End If

End Sub

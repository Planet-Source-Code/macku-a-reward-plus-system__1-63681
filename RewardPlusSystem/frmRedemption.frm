VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRedemption 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Redemption"
   ClientHeight    =   6615
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRedemption.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F11 - &Delete"
      Height          =   375
      Left            =   5880
      TabIndex        =   11
      Top             =   6120
      Width           =   1335
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "F10-Browse Reward"
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F9 - F&ind"
      Height          =   375
      Left            =   2880
      TabIndex        =   9
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F3 - &Save"
      Height          =   375
      Left            =   1560
      TabIndex        =   8
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "F2 - &New"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   6120
      Width           =   1335
   End
   Begin VB.ComboBox cboReward 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   1560
      Width           =   4215
   End
   Begin VB.TextBox txtCelnum 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   3015
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   1920
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   6376
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "IL1"
      SmallIcons      =   "IL1"
      ForeColor       =   16711680
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Reward"
         Object.Width           =   9172
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Points"
         Object.Width           =   3704
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   0
      Top             =   4680
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
            Picture         =   "frmRedemption.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRedemption.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRedemption.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRedemption.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRedemption.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRedemption.frx":19A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblTotalPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5400
      TabIndex        =   15
      Top             =   5640
      Width           =   1935
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Points:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   4200
      TabIndex        =   14
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label lblCurrentPoints 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1320
      TabIndex        =   13
      Top             =   5640
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Current Points:"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   12
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label cboRewarda 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Reward"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   6
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -120
      Picture         =   "frmRedemption.frx":1D42
      Top             =   6000
      Width           =   15030
   End
   Begin VB.Label lblName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   285
      Left            =   3120
      TabIndex        =   1
      Top             =   960
      Width           =   4215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Celphone No."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   3
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3120
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frmRedemption.frx":3A1C
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -120
      Picture         =   "frmRedemption.frx":42E6
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmRedemption"
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
Private rs2 As New ADODB.Recordset
Private Sub countPoints()
Dim a, b

b = 0
For a = 1 To LV1.ListItems.Count
    b = Val(b) + Val(LV1.ListItems(a).SubItems(1))
Next a

lblTotalPoints.Caption = b
End Sub

Private Sub SaveNow()
Dim a

If lblName.Caption = "" Then
    MsgBox "Name not found.", vbCritical, "Mackusoft"
    txtCelnum.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
    
If txtCelnum.Text = "" Then
    MsgBox "Celphone number not found.", vbCritical, "Mackusoft"
    txtCelnum.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

For a = 1 To LV1.ListItems.Count
    rs.Open "Redemption", dbconn, 1, 3
    
    With rs
        .AddNew
        !celnum = txtCelnum
        !Name = lblName
        !reward = LV1.ListItems(a)
        !Points = LV1.ListItems(a).SubItems(1)
        !Date = Format(Date, "mm/dd/yyyy")
        !Time = Time
        .Update
    End With
    
    Set rs = Nothing
    
    rs.Open "select * from totalpoints where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If Not rs.RecordCount = 0 Then
        With rs
            !Points = Val(!Points) - Val(LV1.ListItems(a).SubItems(1))
            .Update
        End With
    End If
    
    Set rs = Nothing
Next a
End Sub

Private Sub clear()
txtCelnum.Text = ""
lblName.Caption = ""
Call loadRewards
LV1.ListItems.clear
lblTotalPoints.Caption = "0"
lblCurrentPoints.Caption = "0"
End Sub

Private Sub loadRewards()
rs.Open "tblrewards order by reward", dbconn, 1, 3

cboReward.clear
Do Until rs.EOF
    cboReward.AddItem rs!reward
    rs.MoveNext
Loop

Set rs = Nothing
End Sub

Private Sub cboCategories_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboReward.SetFocus
End If
End Sub

Private Sub cboReward_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim a As Variant
    Dim vPoints
    
    vPoints = 0
    
    If lblName.Caption = "" Then
        MsgBox "Name not found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    If txtCelnum.Text = "" Then
        MsgBox "Celphone number not found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
      
    If cboReward.Text = "" Then
        MsgBox "Reward not found.", vbCritical, "Mackusoft"
        cboReward.SetFocus
        Exit Sub
    End If
    
    rs.Open "select * from totalpoints where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        MsgBox "Currently need more points to earn on selected reward.", vbInformation, "Mackusoft"
        Set rs = Nothing
        cboReward.SetFocus
        Exit Sub
    Else
        rs2.Open "select * from tblrewards where reward = '" & cboReward & "'", dbconn, 1, 3
                    
        vPoints = Val(lblTotalPoints) + Val(rs2!pointstoearn)
                    
        If Val(vPoints) > Val(lblCurrentPoints) Then
            MsgBox "No more available points on selected reward.", vbInformation, "Mackusoft"
            Set rs2 = Nothing
            Set rs = Nothing
            cboReward.SetFocus
            
            Exit Sub
        Else
            If Val(rs!Points) >= Val(rs2!pointstoearn) Then
                Set a = LV1.ListItems.Add(, , cboReward, 1, 1)
                With a
                    .SubItems(1) = rs2!pointstoearn
                End With
                
                lblTotalPoints.Caption = Val(lblTotalPoints) + Val(rs2!pointstoearn)
            
                Set rs2 = Nothing
            Else
                MsgBox "Need more points to the selected reward.", vbInformation, "Mackusoft"
                Set rs2 = Nothing
                Set rs = Nothing
           
                cboReward.SetFocus
                Exit Sub
            End If
        End If
    End If
            
    Set rs2 = Nothing
    Set rs = Nothing
    
    Call loadRewards
    
    cboReward.SetFocus
End If
        
End Sub

Private Sub cmdBrowse_Click()
Me.MousePointer = vbHourglass
rs.Open "tblrewards order by reward", dbconn, 1, 3
Call WriteListView(frmFind.LV1, rs, 1, 2)
Set rs = Nothing
Me.MousePointer = vbDefault
strFind = "Browse"
frmFind.LV1.ColumnHeaders(1).Text = "Reward"
frmFind.LV1.ColumnHeaders(2).Text = "Points to Earn"

frmFind.Show vbModal
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next

LV1.ListItems.Remove (LV1.SelectedItem.Index)
Call countPoints
End Sub

Private Sub cmdFind_Click()
Dim ans As String

ans = InputBox("Enter Keyword:", "Mackusoft")

Me.MousePointer = vbHourglass
rs.Open "select * from custinfo where left(name, " & Len(ans) & ") = '" & ans & "'", dbconn, 1, 3
Call WriteListView(frmFind.LV1, rs, 1, 2)
Set rs = Nothing
Me.MousePointer = vbDefault
strFind = "Redemption"

frmFind.Show vbModal
End Sub

Private Sub cmdNew_Click()
Dim ans As String

If LV1.ListItems.Count > 0 Then
    ans = MsgBox("Do you want to save?", vbYesNoCancel + vbQuestion, "Mackusoft")
    
    If ans = vbYes Then
        Call SaveNow
    ElseIf ans = vbCancel Then
        Exit Sub
    End If
End If

Call clear
txtCelnum.SetFocus
End Sub

Private Sub cmdSave_Click()
Call SaveNow
Call clear
txtCelnum.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF2
        cmdNew.Value = True
    Case vbKeyF3
        cmdSave.Value = True
    Case vbKeyF9
        cmdFind.Value = True
    Case vbKeyF10
        cmdBrowse.Value = True
    Case vbKeyF11
        cmdDelete.Value = True
    Case vbKeyEscape
        cmdNew.Value = True
End Select
End Sub

Private Sub Form_Load()
Call CenterScreen(Me, Screen.Height - 1400, Screen.Width)
Call loadRewards
MDIFrmMain.StatusBar1.Panels(2).Text = "Redemption..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyDelete
        cmdDelete.Value = True
End Select
End Sub

Private Sub txtCelnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rs.Open "select * from custinfo where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        MsgBox "Record not found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
    Else
        lblName.Caption = rs!Name
        cboReward.SetFocus
    End If
    
    Set rs = Nothing
    
    rs.Open "select * from totalpoints where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If Not rs.RecordCount = 0 Then
        lblCurrentPoints.Caption = rs!Points
    End If
    
    Set rs = Nothing
End If
End Sub

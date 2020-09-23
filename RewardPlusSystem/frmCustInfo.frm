VERSION 5.00
Begin VB.Form frmCustInfo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customer Info"
   ClientHeight    =   3330
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5505
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCustInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "Esc-&Cancel"
      Height          =   375
      Left            =   4440
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   14
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "F11-&Delete"
      Height          =   375
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   13
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdFind 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "F9 - F&ind"
      Height          =   375
      Left            =   2280
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   12
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000E&
      Caption         =   "F3 - &Save"
      Height          =   375
      Left            =   1200
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   11
      Top             =   2880
      Width           =   975
   End
   Begin VB.CommandButton cmdNew 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "F2 - &New"
      Height          =   375
      Left            =   120
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   10
      Top             =   2880
      Width           =   975
   End
   Begin VB.TextBox txtTelnum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   1815
   End
   Begin VB.ComboBox cboGender 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox txtAddress 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   3495
   End
   Begin VB.TextBox txtCelnum 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5760
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Tel. No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   9
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Gender"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Address"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1920
      TabIndex        =   6
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Celphone No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmCustInfo.frx":058A
      Top             =   120
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -1320
      Picture         =   "frmCustInfo.frx":0E6C
      Top             =   0
      Width           =   7530
   End
End
Attribute VB_Name = "frmCustInfo"
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
Private Sub clear()
txtCelnum.Text = ""
txtName.Text = ""
txtAddress = ""
Call LoadGender
txtTelnum.Text = ""
End Sub

Private Sub LoadGender()
With cboGender
    .clear
    .AddItem "MALE"
    .AddItem "FEMALE"
End With
End Sub

Private Sub cmdCancel_Click()
cmdNew.Value = True
End Sub

Private Sub cmdDelete_Click()
Dim ans As String
rs.Open "select * from custinfo where celnum = '" & txtCelnum & "'", dbconn, 1, 3

If Not rs.RecordCount = 0 Then
    ans = MsgBox("Delete record of " & UCase(txtName) & "?", vbQuestion + vbYesNo, "Mackusoft")
    
    If ans = vbYes Then
        rs2.Open "CustinfoArchive", dbconn, 1, 3
        
        With rs2
            .AddNew
            !celnum = rs!celnum
            !Name = UCase(rs!Name)
            !Address = UCase(rs!Address)
            !Gender = rs!Gender
            !Telnum = rs!Telnum
            !Date = rs!Date
            !Time = rs!Time
            !Datedeleted = Format(Date, "mm/dd/yyyy")
            !Timedeleted = Time
            .Update
        End With
        
        Set rs2 = Nothing
        
        rs.Delete
    End If
End If

Set rs = Nothing

cmdNew.Value = True
End Sub

Private Sub cmdFind_Click()
Dim ans As String

ans = InputBox("Enter Keyword:", "Mackusoft")

Me.MousePointer = vbHourglass
rs.Open "select * from custinfo where left(name, " & Len(ans) & ") = '" & ans & "'", dbconn, 1, 3
Call WriteListView(frmFind.LV1, rs, 1, 2)
Set rs = Nothing
Me.MousePointer = vbDefault
strFind = "CustInfo"

frmFind.Show vbModal
End Sub

Private Sub cmdNew_Click()
Call clear
txtCelnum.SetFocus
End Sub

Private Sub cmdSave_Click()
If txtCelnum.Text = "" Then
    MsgBox "Celphone number is missing.", vbCritical, "Mackusoft"
    txtCelnum.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

If txtName.Text = "" Then
    MsgBox "Name not found.", vbCritical, "Mackusoft"
    txtName.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

If cboGender.Text = "" Then
    MsgBox "Gender not found.", vbCritical, "Mackusoft"
    cboGender.SetFocus
    Exit Sub
End If

If txtAddress.Text = "" Then
    txtAddress = " "
End If

If txtTelnum.Text = "" Then
    txtTelnum.Text = " "
End If

rs.Open "select * from custinfo where celnum = '" & txtCelnum & "'", dbconn, 1, 3

If rs.RecordCount = 0 Then
    With rs
        .AddNew
        !celnum = txtCelnum
        !Name = UCase(txtName)
        !Address = UCase(txtAddress)
        !Gender = cboGender
        !Telnum = txtTelnum
        !Date = Format(Date, "mm/dd/yyyy")
        !Time = Time
        .Update
    End With
Else
    With rs
        !celnum = txtCelnum
        !Name = UCase(txtName)
        !Address = UCase(txtAddress)
        !Gender = cboGender
        !Telnum = txtTelnum
        !Date = Format(Date, "mm/dd/yyyy")
        !Time = Time
        .Update
    End With
End If

Set rs = Nothing

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
    Case vbKeyF11
        cmdDelete.Value = True
    Case vbKeyEscape
        cmdCancel.Value = True
End Select
End Sub

Private Sub Form_Load()
Call CenterScreen(Me, Screen.Height - 1900, Screen.Width)
Call LoadGender
MDIFrmMain.StatusBar1.Panels(2).Text = "Customer Info..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub txtCelnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rs.Open "select * from custinfo where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If Not rs.RecordCount = 0 Then
        txtName = rs!Name
        txtAddress = rs!Address
        cboGender = rs!Gender
        txtTelnum = rs!Telnum
    End If
    
    Set rs = Nothing
    
    txtName.SetFocus
    SendKeys "{home}+{end}"
End If
End Sub

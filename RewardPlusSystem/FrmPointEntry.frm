VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form FrmPointEntry 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Entry"
   ClientHeight    =   6735
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   10485
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPointEntry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ESC - &Cancel"
      Height          =   375
      Left            =   8040
      TabIndex        =   17
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "F11 - &Delete"
      Height          =   375
      Left            =   6240
      TabIndex        =   16
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "F9 - F&ind"
      Height          =   375
      Left            =   4440
      TabIndex        =   15
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "F3 - &Save"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "F2 - &New"
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   6240
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDate 
      Height          =   285
      Left            =   3840
      TabIndex        =   11
      Top             =   1560
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   7680
      TabIndex        =   6
      Top             =   2160
      Width           =   2775
   End
   Begin VB.TextBox txtParticulars 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   5
      Top             =   2160
      Width           =   7575
   End
   Begin VB.ComboBox cboCategories 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1560
      Width           =   3735
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
      TabIndex        =   2
      Top             =   960
      Width           =   3735
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3615
      Left            =   0
      TabIndex        =   10
      Top             =   2520
      Width           =   10455
      _ExtentX        =   18441
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
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category"
         Object.Width           =   5733
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Particulars"
         Object.Width           =   7497
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Points"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   0
      Top             =   5160
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
            Picture         =   "FrmPointEntry.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPointEntry.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPointEntry.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPointEntry.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPointEntry.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPointEntry.frx":19A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Date"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   4
      Left            =   3840
      TabIndex        =   12
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amount"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   3
      Left            =   7680
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Particulars"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   2
      Left            =   0
      TabIndex        =   8
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Categories"
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   975
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
      Left            =   3840
      TabIndex        =   3
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3840
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Celphone No."
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "FrmPointEntry.frx":1D42
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -240
      Picture         =   "FrmPointEntry.frx":260C
      Top             =   0
      Width           =   15030
   End
   Begin VB.Image Image4 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -120
      Picture         =   "FrmPointEntry.frx":4EC4
      Top             =   6120
      Width           =   15030
   End
   Begin VB.Image Image3 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   480
      Left            =   -240
      Picture         =   "FrmPointEntry.frx":6B9E
      Top             =   6600
      Visible         =   0   'False
      Width           =   7530
   End
End
Attribute VB_Name = "FrmPointEntry"
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

Private Sub SaveNow()
Dim a, b

b = 0

If txtCelnum.Text = "" Then
    MsgBox "Celphone number not found.", vbCritical, "Mackusoft"
    txtCelnum.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If

If lblName.Caption = "" Then
    MsgBox "Name not found.", vbCritical, "Mackusoft"
    txtCelnum.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
    
If Not IsDate(mskDate) Then
    MsgBox "Invalid date format.", vbCritical, "Mackusoft"
    mskDate.SetFocus
    SendKeys "{home}+{end}"
    Exit Sub
End If
    
For a = 1 To LV1.ListItems.Count
    b = Val(b) + Val(LV1.ListItems.Item(a).SubItems(3))
    
    rs.Open "pointentry", dbconn, 1, 3
        With rs
            .AddNew
            !celnum = txtCelnum
            !Name = lblName
            !Category = LV1.ListItems.Item(a)
            !particular = LV1.ListItems.Item(a).SubItems(1)
            !amount = LV1.ListItems.Item(a).SubItems(2)
            !Points = LV1.ListItems.Item(a).SubItems(3)
            !Date = mskDate
            !Time = Time
            .Update
        End With
    Set rs = Nothing
    
    rs2.Open "select * from pointassessment where celnum = '" & txtCelnum & "' and category = '" & LV1.ListItems.Item(a) & "'", dbconn, 1, 3
    
    If rs2.RecordCount = 0 Then
        With rs2
            .AddNew
            !celnum = txtCelnum
            !Name = lblName
            !Category = LV1.ListItems.Item(a)
            !Points = LV1.ListItems.Item(a).SubItems(3)
            .Update
        End With
    Else
        With rs2
            !Points = Val(!Points) + Val(LV1.ListItems.Item(a).SubItems(3))
            .Update
        End With
    End If
        
    Set rs2 = Nothing
Next a

rs2.Open "select * from totalpoints where celnum = '" & txtCelnum & "'", dbconn, 1, 3

If rs2.RecordCount = 0 Then
    With rs2
        .AddNew
        !celnum = txtCelnum
        !Name = lblName
        !Points = b
        .Update
    End With
Else
    With rs2
        !Points = Val(!Points) + Val(b)
        .Update
    End With
End If

Set rs2 = Nothing
End Sub

Private Sub clear()
txtCelnum.Text = ""
lblName.Caption = ""
Call LoadCategories
mskDate.Text = Format(Date, "mm/dd/yyyy")
txtParticulars.Text = ""
txtAmount.Text = ""
LV1.ListItems.clear
End Sub
Private Sub LoadCategories()
rs.Open "categories order by categoryname", dbconn, 1, 3

cboCategories.clear
Do Until rs.EOF
    cboCategories.AddItem rs!categoryname
    rs.MoveNext
Loop

Set rs = Nothing
End Sub

Private Sub cboCategories_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtParticulars.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
cmdNew.Value = True
End Sub

Private Sub cmdDelete_Click()
On Error Resume Next
Dim a
a = LV1.SelectedItem.Index
LV1.ListItems.Remove (a)
End Sub

Private Sub cmdFind_Click()
Dim ans As String

ans = InputBox("Enter Keyword:", "Mackusoft")

Me.MousePointer = vbHourglass
rs.Open "select * from custinfo where left(name, " & Len(ans) & ") = '" & ans & "'", dbconn, 1, 3
Call WriteListView(frmFind.LV1, rs, 1, 2)
Set rs = Nothing
Me.MousePointer = vbDefault
strFind = "PointEntry"

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
    Case vbKeyF11
        cmdDelete.Value = True
    Case vbKeyEscape
        cmdCancel.Value = True
End Select
End Sub

Private Sub Form_Load()
Call CenterScreen(Me, Screen.Height - 1600, Screen.Width)
mskDate.Text = Format(Date, "mm/dd/yyyy")
Call LoadCategories
MDIFrmMain.StatusBar1.Panels(2).Text = "Point Entry..."
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

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
Dim a

If KeyAscii = 13 Then
    If txtCelnum.Text = "" Then
        MsgBox "Celphone number not found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    If cboCategories.Text = "" Then
        MsgBox "Categories not found.", vbCritical, "Mackusoft"
        cboCategories.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(mskDate) Then
        MsgBox "Invalid date format.", vbCritical, "Mackusoft"
        mskDate.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    If txtParticulars.Text = "" Then
        MsgBox "Particulars not defined.", vbCritical, "Mackusoft"
        txtParticulars.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    If Not IsNumeric(txtAmount) Then
        MsgBox "Value not valid.", vbCritical, "Mackusoft"
        txtAmount.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    rs.Open "select * from categories where categoryname = '" & cboCategories & "'", dbconn, 1, 3
        
    Set a = LV1.ListItems.Add(, , cboCategories, 1, 1)
    With a
        .SubItems(1) = txtParticulars
        .SubItems(2) = txtAmount
        .SubItems(3) = Format$(Int(txtAmount) / Int(rs!amount), "0.00")
    End With
    
    Set rs = Nothing
    
    txtParticulars.Text = ""
    txtAmount.Text = ""
    cboCategories.SetFocus
End If
End Sub

Private Sub txtCelnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rs.Open "select * from custinfo where celnum = '" & txtCelnum & "'", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        MsgBox "No celphone number found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
    Else
        lblName.Caption = rs!Name
        cboCategories.SetFocus
    End If
    
    Set rs = Nothing
End If
End Sub

Private Sub txtParticulars_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtAmount.SetFocus
End If
End Sub

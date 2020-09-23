VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCategories 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Categories"
   ClientHeight    =   6855
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCategories.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAmount 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtPoints 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.TextBox txtCategoryName 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   5535
      Left            =   0
      TabIndex        =   7
      Top             =   1320
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9763
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
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category Name"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Point(s)"
         Object.Width           =   3969
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Amount"
         Object.Width           =   3881
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   0
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
            Picture         =   "frmCategories.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCategories.frx":19A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Amount"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4800
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Point(s)"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Category Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "In Every"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4080
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   120
      Picture         =   "frmCategories.frx":1D42
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -360
      Picture         =   "frmCategories.frx":260C
      Top             =   0
      Width           =   7530
   End
End
Attribute VB_Name = "frmCategories"
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

Private Sub clear()
txtCategoryName.Text = ""
txtPoints.Text = ""
txtAmount.Text = ""
End Sub
Private Sub LoadCategories()
Me.MousePointer = vbHourglass
rs.Open "categories order by categoryname", dbconn, 1, 3
Call WriteListView(LV1, rs, 1, 3)
Set rs = Nothing
Me.MousePointer = vbDefault
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Call LoadCategories
MDIFrmMain.StatusBar1.Panels(2).Text = "Categories..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub LV1_Click()
Call ShowSelected
End Sub

Private Sub ShowSelected()
On Error Resume Next
txtCategoryName.Text = LV1.SelectedItem
txtPoints.Text = LV1.SelectedItem.SubItems(1)
txtAmount.Text = LV1.SelectedItem.SubItems(2)
End Sub

Private Sub LV1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
Select Case KeyCode
    Case vbKeyDelete
        rs.Open "categories where categoryname = '" & LV1.SelectedItem & "'", dbconn, 1, 3
        
        If Not rs.RecordCount = 0 Then
            rs.Delete
            txtCategoryName.Text = ""
            txtPoints.Text = ""
            txtAmount.Text = ""
        End If
        
        Set rs = Nothing
        Call LoadCategories
End Select
End Sub

Private Sub LV1_KeyUp(KeyCode As Integer, Shift As Integer)
Call ShowSelected
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Not IsNumeric(txtPoints) Then
        MsgBox "Invalid value.", vbCritical, "Mackusoft"
        txtPoints.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    If Not IsNumeric(txtAmount) Then
        MsgBox "Invalid value.", vbCritical, "Mackusoft"
        txtAmount.SetFocus
        SendKeys "{home}+{end}"
        Exit Sub
    End If
    
    rs.Open "select * from categories where categoryname = '" & txtCategoryName & "'", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        With rs
            .AddNew
            !categoryname = UCase(txtCategoryName)
            !Points = txtPoints
            !amount = txtAmount
            .Update
        End With
    Else
        With rs
            !Points = txtPoints
            !amount = txtAmount
            .Update
        End With
    End If
    
    Set rs = Nothing
    
    Call LoadCategories
    Call clear
    txtCategoryName.SetFocus
End If
End Sub

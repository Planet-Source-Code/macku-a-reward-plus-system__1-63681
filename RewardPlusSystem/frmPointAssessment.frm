VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPointAssessment 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Point Assessment"
   ClientHeight    =   2055
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPointAssessment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmPointAssessment.frx":058A
   ScaleHeight     =   2055
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
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
      Width           =   2295
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   4575
      Left            =   0
      TabIndex        =   4
      Top             =   2160
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "IL1"
      SmallIcons      =   "IL1"
      ForeColor       =   192
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Category"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Points"
         Object.Width           =   3616
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   0
      Top             =   5280
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
            Picture         =   "frmPointAssessment.frx":0E54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPointAssessment.frx":11EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPointAssessment.frx":1596
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPointAssessment.frx":193A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPointAssessment.frx":1CD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPointAssessment.frx":2272
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   6360
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label lblTotalPoints 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   1560
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Total Points:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
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
      Left            =   2400
      TabIndex        =   1
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Name"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Celphone No."
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "frmPointAssessment.frx":260C
      Top             =   75
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -240
      Picture         =   "frmPointAssessment.frx":2ED6
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmPointAssessment"
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

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyF9
        Dim ans As String

        ans = InputBox("Enter Keyword:", "Mackusoft")

        Me.MousePointer = vbHourglass
        rs.Open "select * from custinfo where left(name, " & Len(ans) & ") = '" & ans & "'", dbconn, 1, 3
        Call WriteListView(frmFind.LV1, rs, 1, 2)
        Set rs = Nothing
        Me.MousePointer = vbDefault
        strFind = "PointAssessment"

        frmFind.Show vbModal
    Case vbKeyEscape
        Unload Me
End Select
End Sub

Private Sub Form_Load()
Call CenterScreen(Me, Screen.Height - 1700, Screen.Width)
MDIFrmMain.StatusBar1.Panels(2).Text = "Point Assessment..."
End Sub

Private Sub Form_Unload(Cancel As Integer)
MDIFrmMain.StatusBar1.Panels(2).Text = ""
End Sub

Private Sub txtCelnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rs.Open "select custinfo.celnum, custinfo.name, totalpoints.points from custinfo, totalpoints " & _
            "where custinfo.celnum = '" & txtCelnum & "' and custinfo.celnum = totalpoints.celnum", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        MsgBox "No record found.", vbCritical, "Mackusoft"
        txtCelnum.SetFocus
        SendKeys "{home}+{end}"
    Else
        lblName.Caption = rs!Name
        lblTotalPoints.Caption = rs!Points
    End If
    
    Set rs = Nothing
    'Dim a As Variant
    'Dim vtotal
    
    'rs.Open "select custinfo.celnum, custinfo.name, pointassessment.category, pointassessment.points from" & _
            " custinfo, pointassessment where custinfo.celnum = '" & txtCelnum & "' and custinfo.celnum = pointassessment.celnum order by " & _
            "pointassessment.category", dbconn, 1, 3
    
    'If rs.RecordCount = 0 Then
    '    MsgBox "No record found.", vbCritical, "Mackusoft"
    '    txtCelnum.SetFocus
    '    SendKeys "{home}+{end}"
    'Else
    '    lblName.Caption = rs!Name
        
    '    vtotal = 0
    '    Do Until rs.EOF
    '        Set a = LV1.ListItems.Add(, , rs!Category, 1, 1)
    '        a.SubItems(1) = rs!Points
            
    '        vtotal = Val(vtotal) + Val(rs!Points)
    '        rs.MoveNext
    '    Loop
        
    '    lblTotalPoints.Caption = vtotal
    'End If
    
    'Set rs = Nothing
End If
        
End Sub

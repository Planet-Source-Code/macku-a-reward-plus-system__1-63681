VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTelecomPrefixes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Telecom Prefixes"
   ClientHeight    =   4950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTelecomPrefixes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   4935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPrefixes 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtTelecom 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin MSComctlLib.ListView LV1 
      Height          =   3615
      Left            =   0
      TabIndex        =   4
      Top             =   1320
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6376
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
         Text            =   "Telecom"
         Object.Width           =   4322
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Prefixes"
         Object.Width           =   4322
      EndProperty
   End
   Begin MSComctlLib.ImageList IL1 
      Left            =   0
      Top             =   4320
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
            Picture         =   "frmTelecomPrefixes.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelecomPrefixes.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelecomPrefixes.frx":0CCC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelecomPrefixes.frx":1070
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelecomPrefixes.frx":140C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTelecomPrefixes.frx":19A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Prefixes"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2520
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Telecom"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   240
      Picture         =   "frmTelecomPrefixes.frx":1D42
      Top             =   80
      Width           =   480
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   630
      Left            =   -120
      Picture         =   "frmTelecomPrefixes.frx":260C
      Top             =   0
      Width           =   7530
   End
End
Attribute VB_Name = "frmTelecomPrefixes"
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

Private Sub txtPrefixes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    rs.Open "select * from telecomprefixes where telecom = '" & txtTelecom & "'", dbconn, 1, 3
    
    If rs.RecordCount = 0 Then
        With rs
            .AddNew
            !telecom = UCase(txtTelecom)
            !prefixes = txtPrefixes
            .Update
        End With
    Else
        With rs
            !telecom = UCase(txtTelecom)
            !prefixes = txtPrefixes
            .Update
        End With
    End If
    
    Set rs = Nothing
End If
End Sub


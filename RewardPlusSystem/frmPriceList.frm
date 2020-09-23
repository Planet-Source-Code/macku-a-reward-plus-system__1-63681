VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPriceList 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Price List"
   ClientHeight    =   6975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPriceList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   12303
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      FocusRect       =   0
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
End
Attribute VB_Name = "frmPriceList"
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

Dim objExcel As Excel.Application
Dim objWorkbook As Excel.Workbook
Dim objWorksheet As Excel.Worksheet

Public Sub LoadWorksheet()
Dim i As Long
Dim n As Long
On Error Resume Next
Set objExcel = GetObject(, "Excel.Application")
If Err.Number Then
   Err.clear
   Set objExcel = CreateObject("Excel.Application")
   If Err.Number Then
      MsgBox "Can't open Excel."
   End If
End If
objExcel.Visible = True
Set objWorkbook = objExcel.Workbooks.Open(App.Path & "\xls\PriceList.xls")
Set objWorksheet = objWorkbook.ActiveSheet

Set objWorkbook = Nothing
Set objWorksheet = Nothing
Set objExcel = Nothing

'With MSFlexGrid1
'.Cols = 50
'.Rows = 100
'For i = 0 To .Rows - 1
'    .Row = i
'    For n = 0 To .Cols - 1
'        .Col = n
'        .Text = objWorksheet.Cells(i + 1, n + 1).Value
'    Next
'Next
'End With

'Me.SetFocus
'Unload Me
End Sub

Private Sub Form_Load()
Call LoadWorksheet
End Sub

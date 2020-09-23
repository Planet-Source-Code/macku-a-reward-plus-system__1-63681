Attribute VB_Name = "modProcedure"
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

Public Sub CenterScreen(ByRef myForm As Form, ByVal frmHeight As Integer, ByVal frmWidth As Integer)
    myForm.Move (frmWidth - myForm.Width) / 2, (frmHeight - myForm.Height) / 2
End Sub
Public Sub WriteListView3(ByRef vListView As ListView, ByVal recSet As ADODB.Recordset, ByVal myIcon As Byte, ByVal TotalFields As Byte)
Dim a As Variant
Dim b As Variant

Dim rsSet As New ADODB.Recordset

vListView.ListItems.clear

Do Until recSet.EOF
    '==
    rsSet.Open "select * from tblitems where itemcode = '" & recSet!itemcode & "'", dbconn, 1, 3
    
    '==
    If Not rsSet.RecordCount = 0 Then
        Set a = vListView.ListItems.Add(, , recSet.Fields(0), myIcon, myIcon)
    
        For b = 0 To (Val(TotalFields) - 1)
            If Not Val(b) = 0 Then
                '==
                If Val(b) = 5 Then
                    a.SubItems(Val(b)) = rsSet!srp
                Else
                    a.SubItems(Val(b)) = recSet.Fields(Val(b))
                End If
            End If
        Next
    '==
    End If
    
    '==
    Set rsSet = Nothing
    recSet.MoveNext
Loop

b = 0
Set recSet = Nothing
End Sub

Public Sub WriteListView(ByRef vListView As ListView, ByVal recSet As ADODB.Recordset, ByVal myIcon As Byte, ByVal TotalFields As Byte)
Dim a As Variant
Dim b As Variant

vListView.ListItems.clear

Do Until recSet.EOF
    Set a = vListView.ListItems.Add(, , recSet.Fields(0), myIcon, myIcon)
    
    For b = 0 To (Val(TotalFields) - 1)
        If Not Val(b) = 0 Then
            a.SubItems(Val(b)) = recSet.Fields(Val(b))
        End If
    Next
    
    recSet.MoveNext
Loop

b = 0
Set recSet = Nothing
    
End Sub

Public Sub WriteListView2(ByRef vListView As ListView, ByVal recSet As ADODB.Recordset, ByVal myIcon As Byte, ByVal TotalFields As Byte)
Dim a As Variant
Dim b As Variant

vListView.ListItems.clear

Do Until recSet.EOF
    Set a = vListView.ListItems.Add(, , recSet.Fields(0), myIcon, myIcon)
    
    For b = 0 To (Val(TotalFields) - 1)
        If Not Val(b) = 0 Then
            a.SubItems(Val(b)) = recSet.Fields(Val(b))
        End If
    Next
    
    recSet.MoveNext
Loop

b = 0
Set recSet = Nothing
End Sub
Public Function AddProgBar(pb As ProgressBar, sb As StatusBar, lPan As Long)

sb.Align = 2
sb.Refresh

pb.ZOrder 0
pb.Appearance = ccFlat
pb.BorderStyle = ccNone

pb.Left = sb.Panels(lPan).Left + 25
pb.Width = sb.Panels(lPan).Width - 45
pb.Top = sb.Top + 55
pb.Height = sb.Height - 100

End Function

Public Function CurrencyToText(curValue As Currency) As String
    Static Ones(10) As String
    Static Teens(10) As String
    Static Tens(10) As String
    Static Thousands(3) As String
    Dim i As Integer, nPosition As Integer
    Dim nNumber As Integer, nStars As Integer
    Dim bZeroValue As Boolean
    Dim stResult As String, stTemp As String, stStars As String
    Dim stBuffer As String


    If curValue > 999999.99 Then
        MsgBox "The limit of this function is 999999.99", vbExclamation, "Danger, Danger Will Robinson"
        Exit Function
    End If
    Ones(0) = "zero"
    Ones(1) = "one"
    Ones(2) = "two"
    Ones(3) = "three"
    Ones(4) = "four"
    Ones(5) = "five"
    Ones(6) = "six"
    Ones(7) = "seven"
    Ones(8) = "eight"
    Ones(9) = "nine"
    Teens(0) = "ten"
    Teens(1) = "eleven"
    Teens(2) = "twelve"
    Teens(3) = "thirteen"
    Teens(4) = "fourteen"
    Teens(5) = "fifteen"
    Teens(6) = "sixteen"
    Teens(7) = "seventeen"
    Teens(8) = "eighteen"
    Teens(9) = "nineteen"
    Tens(0) = ""
    Tens(1) = "ten"
    Tens(2) = "twenty"
    Tens(3) = "thirty"
    Tens(4) = "forty"
    Tens(5) = "fifty"
    Tens(6) = "sixty"
    Tens(7) = "seventy"
    Tens(8) = "eighty"
    Tens(9) = "ninty"
    Thousands(0) = ""
    Thousands(1) = "thousand"
    'Set the cents portion of the string
    stResult = "& " & Format((curValue - Int(curValue)) * 100, "00") & "/100"
    'Convert the dollar portion to a string
    stTemp = CStr(Int(curValue))
    'parse through string(Dollar ammount)


    For i = Len(stTemp) To 1 Step -1
        'Grab the value of this digit
        nNumber = Val(Mid(stTemp, i, 1))
        'Check the position(column) of this digi
        '     t
        'Ones, Tens, or Hundereds
        nPosition = (Len(stTemp) - i) + 1


        Select Case (nPosition Mod 3)
            Case 1 'Ones position
            bZeroValue = False


            If i = 1 Then
                stBuffer = Ones(nNumber) & " "
            ElseIf Mid(stTemp, i - 1, 1) = "1" Then
                stBuffer = Teens(nNumber) & " "
                i = i - 1 'Skip tens position
            ElseIf nNumber > 0 Then
                stBuffer = Ones(nNumber) & " "
            Else
                bZeroValue = True


                If i > 1 Then


                    If Mid(stTemp, i - 1, 1) <> "0" Then
                        bZeroValue = False
                    End If
                End If


                If i > 2 Then


                    If Mid(stTemp, i - 2, 1) <> "0" Then
                        bZeroValue = False
                    End If
                End If
                stBuffer = ""
            End If


            If bZeroValue = False And nPosition > 1 Then
                stBuffer = stBuffer & Thousands(nPosition / 3) & " "
            End If
            stResult = stBuffer & stResult
            Case 2 'Tens position
            'Numbers like twenty-five need to be hyp
            '     henated. So......
            'Check if the digit has a value other th
            '     an 0 AND check the next
            'digit to see if it has a value other th
            '     an 0
            'if both are true add the hyphen


            If nNumber > 0 And Val(Mid(stTemp, i + 1, 1)) = 0 Then
                stResult = Tens(nNumber) & " " & stResult
            ElseIf nNumber > 0 And Val(Mid(stTemp, i + 1, 1)) > 0 Then
                stResult = Tens(nNumber) & "-" & stResult
            End If
            Case 0 'Hundreds position


            If nNumber > 0 Then
                stResult = Ones(nNumber) & " hundred " & stResult
            End If
        End Select
Next i


If Len(stResult) > 0 Then
    stResult = UCase(Left(stResult, 1)) & Mid(stResult, 2)
End If
nStars = 125 - Len(stResult)


'For i = 0 To nStars
'    stStars = stStars + "*"
'Next
CurrencyToText = stResult & " " & stStars
End Function

Public Sub search_in_listview(ByRef sListView As ListView, ByVal sFindText As String)
Dim tmp_listtview As ListItem
Set tmp_listtview = sListView.FindItem(sFindText, lvwSubItem, lvwPartial, lvwPartial)
If Not tmp_listtview Is Nothing Then
    tmp_listtview.EnsureVisible
    tmp_listtview.Selected = True
End If
End Sub

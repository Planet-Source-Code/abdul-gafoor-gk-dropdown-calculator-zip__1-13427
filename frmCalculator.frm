VERSION 5.00
Begin VB.Form frmCalculator 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   Icon            =   "frmCalculator.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   1  'Arrow
   ScaleHeight     =   2565
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Variables for drawing buttons
Private Const BtnTxt As String = "789/456*123-0.%+C«="
Private BtnLoc() As RECT
Private BtnDownId As Integer

'Variables for calculation purpose
Private FVal As Double, SVal As Double
Private Op1 As String, Op2 As String
Private DispTxt As String

Private ScrX As Long, ScrY As Long

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim PressedKey As String, BtnId As Integer
    
    If (Shift = 0) Then
        Select Case KeyCode
            Case 48 To 57: PressedKey = CStr(Trim(KeyCode - 48))
            Case 96 To 105: PressedKey = CStr(Trim(KeyCode - 96))
            Case 110, 190: PressedKey = "."
            Case 107: PressedKey = "+"
            Case 109, 189: PressedKey = "-"
            Case 106: PressedKey = "*"
            Case 111, 191: PressedKey = "/"
            Case 67: PressedKey = "C"
            Case 8: PressedKey = "«"
            Case 13, 187: PressedKey = "="
            Case Else: Exit Sub
        End Select
    ElseIf (Shift = 1) Then
        Select Case KeyCode
            Case 187: PressedKey = "+"
            Case 56: PressedKey = "*"
            Case 53: PressedKey = "%"
            Case Else: Exit Sub
        End Select
    End If
    
    BtnId = InStr(1, BtnTxt, PressedKey) - 1
    Call ButtonStatus(BtnId, True)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Common.EditBox.SelStart = Len(Common.EditBox.Text)
        Unload Me
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not BtnDownId = -1) Then
        Call DoActionOnButton(BtnDownId)
        If (Mid(BtnTxt, BtnDownId + 1, 1) = "=") Then Exit Sub
        Call ButtonStatus(BtnDownId, False)
    End If
End Sub

Private Sub Form_Load()
    Dim TmpRect As RECT
    
    Width = (255 * 4 + 45 * 3 + 60 * 2)
    Height = (255 * 5 + 45 * 4 + 60 * 2)
    
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
    
    DispTxt = Common.EditBox.Text
    
    Me.ForeColor = Common.ButTxtCol
    
    Call SetRect(TmpRect, 0, 0, ScaleWidth / ScrX, ScaleHeight / ScrY)
    Call DrawEdge(hdc, TmpRect, BDR_RAISEDINNER, BF_RECT)
    
    Call SetCapture(Me.hwnd)
    
    Call SetButtonLocation
    Call DrawButtonFigures
    
    BtnDownId = -1
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Button = 1) And (Shift = 0) Then
        Dim BtnId As Integer
        BtnId = GetClickedButton(X, Y)
        
        If (Not BtnId = -1) Then
            Call ButtonStatus(BtnId, True)
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Check whether mouse is on the button or not
    'If yes, then do the appropriate action
    If (Not BtnDownId = -1) Then
        If IsMouseOnButton(X, Y, BtnDownId) Then
            Call DoActionOnButton(BtnDownId)
            If (Mid(BtnTxt, BtnDownId + 1, 1) = "=") Then Exit Sub
        End If
        
        Call ButtonStatus(BtnDownId, False)
    End If
    
    'Check whether user has clicked on calculator form or not
    'If yes, then set mouse capture again
    'Otherwise release mouse capture and unload form
    Dim blnMouseOver As Boolean
    blnMouseOver = X >= 0 And Y >= 0 And X <= ScaleWidth And Y <= ScaleHeight
    If blnMouseOver Then
        Call SetCapture(Me.hwnd)
    Else
        Call ReleaseCapture
        Call Form_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub DrawButtonFigures()
    Dim i As Integer
    
    For i = 0 To Len(BtnTxt) - 1
        Call DrawText(Me.hdc, Mid(BtnTxt, i + 1, 1), 1&, BtnLoc(i), DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
    Next i
End Sub

Private Sub SetButtonLocation()
    ReDim BtnLoc(0 To Len(BtnTxt))
    
    Dim LPos As Long, TPos As Long
    Dim i As Integer
    
    For i = 0 To Len(BtnTxt) - 1
        LPos = ((i Mod 4) * 45 + 60) + ((i Mod 4) * 255)
        TPos = (Int(i / 4) * 45 + 60) + (Int(i / 4) * 255)
        
        With BtnLoc(i)
            .Left = LPos / ScrX
            .Top = TPos / ScrY
            If (i = (Len(BtnTxt) - 1)) Then
                .Right = (LPos + 45 * 1 + 255 * 2) / ScrX
            Else
                .Right = (LPos + 255) / ScrX
            End If
            .Bottom = (TPos + 255) / ScrY
        End With
        
        Call DrawEdge(Me.hdc, BtnLoc(i), BDR_RAISEDINNER, BF_RECT)
    Next i
End Sub

Private Function GetClickedButton(X As Single, Y As Single) As Integer
    Dim i As Integer
    
    GetClickedButton = -1
    
    For i = 0 To Len(BtnTxt) - 1
        With BtnLoc(i)
            If IsMouseOnButton(X, Y, i) Then
                GetClickedButton = i
                Exit Function
            End If
        End With
    Next i
End Function

Private Function IsMouseOnButton(ByVal X As Single, ByVal Y As Single, ButtonId As Integer) As Boolean
    X = X / ScrX
    Y = Y / ScrY
    
    With BtnLoc(ButtonId)
        IsMouseOnButton = (X >= .Left And X <= .Right) And (Y >= .Top And Y <= .Bottom)
    End With
End Function

'This procedure draws visual effect of a button
'according to the status of button
Private Sub ButtonStatus(ButtonId As Integer, Status As Boolean)
    'Variable declarations
    Dim ClrVal As Long, Brsh As Long
    Dim TmpBtnLoc As RECT
    Dim TmpBtnTxt As String * 1
    
    'Fill the variables with approprite values
    ClrVal = GetBkColor(Me.hdc)
    Brsh = CreateSolidBrush(ClrVal)
    With BtnLoc(ButtonId)
        Call SetRect(TmpBtnLoc, .Left + 1, .Top + 1, .Right + 1, .Bottom + 1)
    End With
    TmpBtnTxt = Mid(BtnTxt, ButtonId + 1, 1)
    
    'Clear the specified button location temporarily
    Call FillRect(Me.hdc, BtnLoc(ButtonId), Brsh)
    
    'If status is true then set the visual effect of button in down position
    'Otherwise make it in normal position
    If Status Then
        Call DrawText(Me.hdc, TmpBtnTxt, 1, TmpBtnLoc, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
        Call DrawEdge(Me.hdc, BtnLoc(ButtonId), BDR_SUNKENOUTER, BF_RECT)
        Me.Refresh
        BtnDownId = ButtonId
    Else
        Call DrawText(Me.hdc, TmpBtnTxt, 1, BtnLoc(ButtonId), DT_SINGLELINE Or DT_CENTER Or DT_VCENTER)
        Call DrawEdge(Me.hdc, BtnLoc(ButtonId), BDR_RAISEDINNER, BF_RECT)
        Me.Refresh
        BtnDownId = -1
    End If
    
    'Delete unwanted objects to free up the memory
    Call DeleteObject(ClrVal)
    Call DeleteObject(Brsh)
End Sub

Private Sub DoActionOnButton(ButtonId)
    Dim KeyAscii As Integer
    KeyAscii = Asc(Mid(BtnTxt, ButtonId + 1, 1))
    
    Select Case KeyAscii
        Case 46, 48 To 57: Call NumberClicked(KeyAscii)
        Case 43, 45, 42, 47: Call OperatorClicked(KeyAscii)
        Case 37: Call PercentClicked
        Case 67: Call ClearClicked
        Case 171: Call BackClicked
        Case 61: Call EnterClicked
    End Select
End Sub

Private Sub NumberClicked(KeyAscii As Integer)
    If (Not InStr(1, DispTxt, ".") = 0) And (KeyAscii = 46) Then
        Exit Sub
    ElseIf (Val(DispTxt) = 0) And (KeyAscii = 46) Then
        DispTxt = "0"
    ElseIf (DispTxt = "") And (KeyAscii = 48) Then
        Exit Sub
    ElseIf (DispTxt = "0") Then
        DispTxt = ""
    End If
    
    DispTxt = DispTxt & Chr(KeyAscii)
    Call UpdateEditBox(DispTxt)
End Sub

Private Sub BackClicked()
    If (Not DispTxt = "") Then DispTxt = Mid(DispTxt, 1, Len(DispTxt) - 1)
    If (DispTxt = "0") Then DispTxt = ""
    If (DispTxt = "") Then DispTxt = "0"
    
    Call UpdateEditBox(DispTxt)
End Sub

Private Sub OperatorClicked(KeyAscii As Integer)
    If (Op1 = "") Then
        FVal = Val(Common.EditBox.Text)
        Op1 = Chr(KeyAscii)
        
        DispTxt = ""
    Else
        Op2 = Chr(KeyAscii)
        
        If ((Not DispTxt = "") Or (Op1 = Op2) Or (Op2 = "=")) Then
            SVal = Val(Common.EditBox.Text)
            FVal = Calculate(FVal, SVal, Op1)
            SVal = 0
            
            DispTxt = ""
        End If
        
        Op1 = Op2
    End If
    
    Call UpdateEditBox(CStr(FVal))
End Sub

Private Sub PercentClicked()
    If Not (Op1 = "") Then
        SVal = Val(Common.EditBox.Text)
        
        Select Case Op1
            Case "+": FVal = FVal + (SVal / 100)
            Case "-": FVal = FVal - (SVal / 100)
            Case "*": FVal = FVal * (SVal / 100)
            Case "/": FVal = FVal / (SVal / 100)
        End Select
        
        SVal = 0
        Op1 = ""
    Else
        FVal = 0
    End If
    
    Call UpdateEditBox(CStr(FVal))
    DispTxt = ""
End Sub

Private Sub ClearClicked()
    Call ClearVariables
    Call UpdateEditBox(0)
End Sub

Private Sub EnterClicked()
    Call OperatorClicked(Asc("="))
    Call ClearVariables
    Call Form_KeyPress(vbKeyEscape)
End Sub

Private Function Calculate(fv As Double, sv As Double, op As String) As Double
    Select Case op
        Case "+": Calculate = fv + sv
        Case "-": Calculate = fv - sv
        Case "*": Calculate = fv * sv
        Case "/": Calculate = fv / sv
    End Select
End Function

Private Sub ClearVariables()
    FVal = 0: SVal = 0
    Op1 = ""
    DispTxt = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FVal = 0: SVal = 0
    Op1 = "": Op2 = ""
    DispTxt = ""
End Sub

Private Function UpdateEditBox(Txt As String)
    Common.EditBox.Text = Txt
    Common.EditBox.Refresh
End Function

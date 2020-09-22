VERSION 5.00
Begin VB.UserControl DDCalculator 
   AutoRedraw      =   -1  'True
   ClientHeight    =   450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1515
   ScaleHeight     =   450
   ScaleWidth      =   1515
   ToolboxBitmap   =   "DDCalculator.ctx":0000
   Begin VB.TextBox txtEdit 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Marlett"
         Size            =   8.25
         Charset         =   2
         Weight          =   500
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1200
      ScaleHeight     =   285
      ScaleWidth      =   195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   190
   End
End
Attribute VB_Name = "DDCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**********************************************************************
'************** Project :  DDCalculator OCX     ***********************
'************** Version :  1.2                  ***********************
'************** Author  :  Abdul Gafoor.GK      ***********************
'************** Date    :  15/August/2000       ***********************
'**********************************************************************
'
'   This OCX Control is simple enough not to describe its use.
'   This is my first ActiveX Control and written out of mere
'   interest in ActiveX and COM objects.
'
'   You will notice that I used API functions for some tasks,
'   even though it is available in VB.  I just wanted to test the
'   the power of those functions since this is my first project,
'   which uses API functions
'
'   This version is dependent to some components from vbAccelerator.
'   (isubclass.cls, subclass.cls, msubclass.bas). With these
'   components I could subclass without any crash. Thanks to
'   vbAccelerator.  (Subclassing is needed for Context Menu of
'   Text box and Hot Tracking of Dropdown Button.)
'
'   If you like this control, please don't forget to send
'   your comments in 'gafoorgk@yahoo.com'
'
'**********************************************************************

Option Explicit

Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function DestroyMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function IsMenu Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal X As Long, ByVal Y As Long, ByVal nReserved As Long, ByVal hwnd As Long, lprc As Any) As Long
Private Const MF_STRING = &H0&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Const MF_BYCOMMAND = &H0&
Private Const TPM_LEFTALIGN = &H0&
Private Const TPM_RETURNCMD = &H100&
Private Const TPM_RIGHTBUTTON = &H2&

Private Declare Function TrackMouseEvent Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As tagTRACKMOUSEEVENT) As Long
Private Type tagTRACKMOUSEEVENT
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type
Private Const TME_LEAVE = &H2
Private Const HOVER_DEFAULT = &HFFFFFFFF

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const WM_CONTEXTMENU = &H7B
Private Const WM_MOUSELEAVE = &H2A3

Implements ISubclass

Private MouseTrack As tagTRACKMOUSEEVENT
Private IsMouseTracking As Boolean

Public Enum dcAppearanceConstants
    Flat
    [3D]
End Enum

Public Enum dcAlignmentConstants
    [Left Justify]
    [Right Justify]
    Center
End Enum

Private IsDropButDown As Boolean
Private IsCalcVisible As Boolean

Private ScrX As Long, ScrY As Long

'For holding the reference to context menu
Public TxtEditContextMenu As Long

'Default Property Values:
Private Const m_def_HotTracking = True
Private Const m_def_DropButtonForeColorOnMouse = &HFF0000
Private Const m_def_ForeColor = &H80000012
Private Const m_def_DropButtonForeColor = &H80000012
Private Const m_def_NumberFormat = ""
Private Const m_def_Value = 0
Private Const m_def_CalcButtonTextColor = &H80000012
Private Const m_def_Appearance = dcAppearanceConstants.[3D]

'Property Variables:
Private m_HotTracking As Boolean
Private m_DropButtonForeColorOnMouse As OLE_COLOR
Private m_ForeColor As OLE_COLOR
Private m_DropButtonForeColor As OLE_COLOR
Private m_NumberFormat As String
Private m_Value As Double 'Single
Private m_Appearance As dcAppearanceConstants

'Event Declarations:
Event Validate(Cancel As Boolean) 'MappingInfo=txtEdit,txtEdit,-1,Validate
Attribute Validate.VB_Description = "Occurs when a control loses focus to a control that causes validation."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtEdit,txtEdit,-1,MouseDown
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Attribute MouseDown.VB_UserMemId = -605
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtEdit,txtEdit,-1,MouseMove
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Attribute MouseMove.VB_UserMemId = -606
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=txtEdit,txtEdit,-1,MouseUp
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Attribute MouseUp.VB_UserMemId = -607
Event Click() 'MappingInfo=txtEdit,txtEdit,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Attribute Click.VB_UserMemId = -600
Attribute Click.VB_MemberFlags = "200"
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtEdit,txtEdit,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Attribute KeyDown.VB_UserMemId = -602
Event KeyPress(KeyAscii As Integer) 'MappingInfo=txtEdit,txtEdit,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Attribute KeyPress.VB_UserMemId = -603
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=txtEdit,txtEdit,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Attribute KeyUp.VB_UserMemId = -604
Event BeforeCalculatorOpen()
Event AfterCalculatorClose()
Event Change() 'MappingInfo=txtEdit,txtEdit,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event DblClick() 'MappingInfo=txtEdit,txtEdit,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Attribute DblClick.VB_UserMemId = -601

Private Property Get ISubclass_MsgResponse() As EMsgResponse
    ISubclass_MsgResponse = emrConsume
End Property

Private Property Let ISubclass_MsgResponse(ByVal RHS As EMsgResponse)
    'Not used
End Property

Private Function ISubclass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If (hwnd = picButton.hwnd) Then
        IsMouseTracking = False
        If m_HotTracking Then Call DrawDropDownSign
    ElseIf (hwnd = txtEdit.hwnd) Then
        'Use GetCursorPos instead of extracting HiWord and LoWord from lParam parameter
        Dim CurPos As POINTAPI
        Call GetCursorPos(CurPos)
        
        'Enable/Disable Copy command in popup menu according to the selection in text box
        If (EditBox.SelText = "") Then
            Call ModifyMenu(TxtEditContextMenu, 1&, MF_BYCOMMAND Or MF_DISABLED Or MF_GRAYED Or MF_STRING, 1&, "&Copy")
        Else
            Call ModifyMenu(TxtEditContextMenu, 1&, MF_BYCOMMAND Or MF_ENABLED Or MF_STRING, 1&, "&Copy")
        End If
        
        'Enable/Disable paste command in popup menu according to the value in clipboard
        If Clipboard.GetFormat(vbCFText) Then
            If (Val(Clipboard.GetText) = 0) Then
                Call ModifyMenu(TxtEditContextMenu, 2&, MF_BYCOMMAND Or MF_DISABLED Or MF_GRAYED Or MF_STRING, 2&, "&Paste")
            Else
                Call ModifyMenu(TxtEditContextMenu, 2&, MF_BYCOMMAND Or MF_ENABLED Or MF_STRING, 2&, "&Paste")
            End If
        Else
            Call ModifyMenu(TxtEditContextMenu, 2&, MF_BYCOMMAND Or MF_DISABLED Or MF_GRAYED Or MF_STRING, 2&, "&Paste")
        End If
        
        'Show popup menu
        Dim ReturnCmdId As Long
        If Not (lParam = -1) Then
            ReturnCmdId = TrackPopupMenu(TxtEditContextMenu, TPM_LEFTALIGN Or TPM_RETURNCMD Or TPM_RIGHTBUTTON, CurPos.X, CurPos.Y, ByVal 0&, txtEdit.hwnd, ByVal 0&)
            
            'Do action according to the selection
            Select Case ReturnCmdId
                Case 1: Clipboard.SetText EditBox.SelText
                Case 2: EditBox.SelText = Clipboard.GetText
            End Select
        End If
    End If
End Function

Private Sub picButton_GotFocus()
    txtEdit.SetFocus
End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (Shift = 0) And (Button = 1) Then
        Call DropDownButtonStatus(True)
    End If
End Sub

Private Sub picButton_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not IsMouseTracking Then
        If m_HotTracking Then
            IsMouseTracking = True
            Call TrackMouseEvent(MouseTrack)
            Call DrawDropDownSign
        End If
    End If
End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim IsMouseInButton As Boolean
    
    If (m_Appearance = [3D]) Then
        If (IsDropButDown) Then
            Call DropDownButtonStatus(False)
        Else
            Exit Sub
        End If
    End If
    
    With picButton
        IsMouseInButton = (X > .ScaleLeft And X < .ScaleWidth) And (Y > .ScaleTop And Y < .ScaleHeight)
    End With
    
    If IsMouseInButton Then
        If Shift = 0 And Button = 1 Then
            Call ShowCalculator
        End If
    End If
End Sub

Private Sub txtEdit_Change()
    RaiseEvent Change
End Sub

Private Sub txtEdit_GotFocus()
    txtEdit = CStr(m_Value)
    
    With txtEdit
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
    Dim strChar As String
    strChar = Chr(KeyAscii)
    
    If (Not IsNumeric(strChar)) Then
        If (strChar = "-") And (txtEdit.SelStart = 0) And (InStr(1, txtEdit, "-") = 0) Then
            Exit Sub
        ElseIf (strChar = ".") And (InStr(1, txtEdit, ".") = 0) Then
            Exit Sub
        ElseIf (KeyAscii = vbKeyBack) Then
            Exit Sub
        End If
        
        KeyAscii = 0
    End If
End Sub

Private Sub txtEdit_LostFocus()
    m_Value = Val(txtEdit)
    
    If IsCalcVisible Then
        txtEdit = CStr(m_Value)
    Else
        txtEdit = CStr(Format(m_Value, m_NumberFormat))
    End If
    
    txtEdit.SelStart = 0
End Sub

Private Sub UserControl_Initialize()
    KeyPreview = True
    
    ScrX = Screen.TwipsPerPixelX
    ScrY = Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    If ((Shift = 0) And (KeyCode = vbKeyF4)) Or _
       ((Shift = 4) And (KeyCode = vbKeyDown)) Then
        Call ShowCalculator
    ElseIf ((Shift = 2) And (KeyCode = vbKeyC)) Then
        If Not (txtEdit.SelText = "") Then Clipboard.SetText txtEdit.SelText
    ElseIf ((Shift = 2) And (KeyCode = vbKeyV)) Then
        If Clipboard.GetFormat(vbCFText) Then
            If Not (Val(Clipboard.GetText) = 0) Then txtEdit.SelText = Clipboard.GetText
        End If
    End If
End Sub

Private Sub UserControl_Resize()
    Height = 285
    If Width < 300 Then Width = 300
    
    Call Draw3DFlatEdge(m_Appearance)
End Sub

Private Sub Draw3DFlatEdge(TypeOfEdge As dcAppearanceConstants)
    Dim Rct As RECT
    
    If TypeOfEdge = [3D] Then
        With txtEdit
            .Left = 2 * ScrX + 15
            .Top = 2 * ScrY + 15
            .Width = ScaleWidth - (4 * ScrX) - 190 - 15
            .Height = ScaleHeight - (4 * ScrY) - 15
        End With
        
        Call SetRect(Rct, 0, 0, ScaleWidth / ScrX, ScaleHeight / ScrY)
        
        Cls
        Call DrawEdge(hdc, Rct, EDGE_SUNKEN, BF_RECT)
        Refresh
        
        With picButton
            .Left = ScaleWidth - 190 - 30
            .Top = ScaleTop + (2 * ScrX)
            .Width = 190
            .Height = ScaleHeight - (4 * ScrX)
            
            Call SetRect(Rct, 0, 0, .ScaleWidth / ScrX, .ScaleHeight / ScrY)
            
            .Cls
            Call DrawEdge(.hdc, Rct, EDGE_RAISED, BF_RECT)
            .Refresh
        End With
    Else
        With txtEdit
            .Left = (1 * ScrX) + 15
            .Top = (1 * ScrY) + 15
            .Width = ScaleWidth - (2 * ScrX) - 190 - 15
            .Height = ScaleHeight - (2 * ScrY) - 15
        End With
        
        Call SetRect(Rct, 0, 0, ScaleWidth / ScrX, ScaleHeight / ScrY)
        
        Cls
        Call DrawEdge(hdc, Rct, BDR_SUNKENOUTER, BF_RECT + BF_FLAT + BF_MONO)
        Refresh
        
        With picButton
            .Left = ScaleWidth - 190
            .Top = 0
            .Width = 190
            .Height = Height
            
            Call SetRect(Rct, 0, 0, .ScaleWidth / ScrX, .ScaleHeight / ScrY)
            
            .Cls
            Call DrawEdge(.hdc, Rct, EDGE_RAISED, BF_RECT + BF_FLAT)
            .Refresh
        End With
    End If
    
    Call DrawDropDownSign
End Sub

Private Sub DropDownButtonStatus(Status As Boolean)
    Dim R As RECT
    
    If Not m_Appearance = Flat Then
        With picButton
            .Cls
            
            Call SetRect(R, 0, 0, .ScaleWidth / ScrX, .ScaleHeight / ScrY)
            
            If Status Then
                Call DrawEdge(.hdc, R, EDGE_RAISED, BF_RECT Or BF_FLAT)
                IsDropButDown = True
                Call DrawDropDownSign
            Else
                Call DrawEdge(.hdc, R, EDGE_RAISED, BF_RECT)
                IsDropButDown = False
                Call DrawDropDownSign
            End If
            
            .Refresh
        End With
    End If
End Sub

Private Sub ShowCalculator()
    Dim EditPos As RECT
    
    Call GetWindowRect(UserControl.hwnd, EditPos)
    
    Load frmCalculator
    With frmCalculator
        .Left = (EditPos.Right * ScrX) - .Width
        .Top = EditPos.Bottom * ScrY
        If (.Top + .Height) > Screen.Height Then
            .Top = EditPos.Top * ScrY - .Height
        End If
        
        'Show the calculator
        IsCalcVisible = True
        RaiseEvent BeforeCalculatorOpen
        .Show vbModal
        
        m_Value = Val(txtEdit)
        IsCalcVisible = False
        RaiseEvent AfterCalculatorClose
    End With
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    Common.ButTxtCol = PropBag.ReadProperty("CalcButtonTextColor", m_def_CalcButtonTextColor)
    m_Value = PropBag.ReadProperty("Value", m_def_Value)
    txtEdit.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Set txtEdit.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    txtEdit.Alignment = PropBag.ReadProperty("Alignment", 0)
    txtEdit.Locked = PropBag.ReadProperty("Locked", False)
    txtEdit.RightToLeft = PropBag.ReadProperty("RightToLeft", False)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    txtEdit.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    txtEdit.LinkItem = PropBag.ReadProperty("LinkItem", "")
    txtEdit.LinkMode = PropBag.ReadProperty("LinkMode", 0)
    txtEdit.LinkTimeout = PropBag.ReadProperty("LinkTimeout", 50)
    txtEdit.LinkTopic = PropBag.ReadProperty("LinkTopic", "")
    m_NumberFormat = PropBag.ReadProperty("NumberFormat", m_def_NumberFormat)
    m_DropButtonForeColor = PropBag.ReadProperty("DropButtonForeColor", m_def_DropButtonForeColor)
    m_HotTracking = PropBag.ReadProperty("HotTracking", m_def_HotTracking)
    m_DropButtonForeColorOnMouse = PropBag.ReadProperty("DropButtonForeColorOnMouse", m_def_DropButtonForeColorOnMouse)
    
    If Ambient.UserMode Then
        TxtEditContextMenu = CreatePopupMenu()
        Call AppendMenu(TxtEditContextMenu, MF_STRING, 1&, "&Copy")
        Call AppendMenu(TxtEditContextMenu, MF_STRING, 2&, "&Paste")
        
        Set Common.EditBox = UserControl.txtEdit
        
        With MouseTrack
            .cbSize = Len(MouseTrack)
            .dwFlags = TME_LEAVE
            .hwndTrack = picButton.hwnd
            .dwHoverTime = HOVER_DEFAULT
        End With
        Call TrackMouseEvent(MouseTrack)
        IsMouseTracking = True
        
        AttachMessage Me, picButton.hwnd, WM_MOUSELEAVE
        AttachMessage Me, txtEdit.hwnd, WM_CONTEXTMENU
    End If
    
    Call Draw3DFlatEdge(m_Appearance)
    
    Call DrawDropDownSign
    If UserControl.Enabled Then txtEdit.ForeColor = m_ForeColor Else txtEdit.ForeColor = vbGrayText
    
    txtEdit = CStr(Format(m_Value, m_NumberFormat))
End Sub

Private Sub UserControl_Terminate()
    Set Common.EditBox = Nothing
    
    If IsMenu(TxtEditContextMenu) Then
        Call DestroyMenu(TxtEditContextMenu)
    End If
    
    On Error Resume Next
    DetachMessage Me, picButton.hwnd, WM_MOUSELEAVE
    DetachMessage Me, txtEdit.hwnd, WM_CONTEXTMENU
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("CalcButtonTextColor", Common.ButTxtCol, m_def_CalcButtonTextColor)
    Call PropBag.WriteProperty("Value", m_Value, m_def_Value)
    Call PropBag.WriteProperty("BackColor", txtEdit.BackColor, &H80000005)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H80000005)
    Call PropBag.WriteProperty("Font", txtEdit.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Alignment", txtEdit.Alignment, 0)
    Call PropBag.WriteProperty("Locked", txtEdit.Locked, False)
    Call PropBag.WriteProperty("RightToLeft", txtEdit.RightToLeft, False)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", txtEdit.MousePointer, 0)
    Call PropBag.WriteProperty("LinkItem", txtEdit.LinkItem, "")
    Call PropBag.WriteProperty("LinkMode", txtEdit.LinkMode, 0)
    Call PropBag.WriteProperty("LinkTimeout", txtEdit.LinkTimeout, 50)
    Call PropBag.WriteProperty("LinkTopic", txtEdit.LinkTopic, "")
    Call PropBag.WriteProperty("NumberFormat", m_NumberFormat, m_def_NumberFormat)
    Call PropBag.WriteProperty("DropButtonForeColor", m_DropButtonForeColor, m_def_DropButtonForeColor)
    Call PropBag.WriteProperty("HotTracking", m_HotTracking, m_def_HotTracking)
    Call PropBag.WriteProperty("DropButtonForeColorOnMouse", m_DropButtonForeColorOnMouse, m_def_DropButtonForeColorOnMouse)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get Appearance() As dcAppearanceConstants
Attribute Appearance.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute Appearance.VB_UserMemId = -520
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As dcAppearanceConstants)
    m_Appearance = New_Appearance
    PropertyChanged "Appearance"
    
    Call Draw3DFlatEdge(m_Appearance)
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Appearance = m_def_Appearance
    Common.ButTxtCol = m_def_CalcButtonTextColor
    m_Value = m_def_Value
    UserControl.BackColor = txtEdit.BackColor
    txtEdit = CStr(m_Value)
    m_NumberFormat = m_def_NumberFormat
    m_DropButtonForeColor = m_def_DropButtonForeColor
    m_ForeColor = m_def_ForeColor
    m_HotTracking = m_def_HotTracking
    m_DropButtonForeColorOnMouse = m_def_DropButtonForeColorOnMouse
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14
Public Function About() As Variant
Attribute About.VB_UserMemId = -552
Attribute About.VB_MemberFlags = "40"
    frmAbout.Show 1
End Function

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CalcButtonTextColor() As OLE_COLOR
Attribute CalcButtonTextColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    CalcButtonTextColor = Common.ButTxtCol
End Property

Public Property Let CalcButtonTextColor(ByVal New_CalcButtonTextColor As OLE_COLOR)
    Common.ButTxtCol = New_CalcButtonTextColor
    PropertyChanged "CalcButtonTextColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get Value() As Double
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Misc"
Attribute Value.VB_MemberFlags = "200"
    Value = m_Value
End Property

Public Property Let Value(ByVal New_Value As Double)
    m_Value = New_Value
    txtEdit = CStr(Format(m_Value, m_NumberFormat))
    PropertyChanged "Value"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
Attribute BackColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute BackColor.VB_UserMemId = -501
    BackColor = txtEdit.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    txtEdit.BackColor() = New_BackColor
    UserControl.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
    Call Draw3DFlatEdge(m_Appearance)
End Property

Private Sub txtEdit_DblClick()
    'RaiseEvent DblClick
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_ProcData.VB_Invoke_Property = ";Font"
Attribute Font.VB_UserMemId = -512
    Set Font = txtEdit.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set txtEdit.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
Attribute ForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
Attribute ForeColor.VB_UserMemId = -513
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_ForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    
    If UserControl.Enabled Then txtEdit.ForeColor() = m_ForeColor
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,Alignment
Public Property Get Alignment() As dcAlignmentConstants
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
Attribute Alignment.VB_ProcData.VB_Invoke_Property = ";Misc"
    Alignment = txtEdit.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As dcAlignmentConstants)
    txtEdit.Alignment() = New_Alignment
    PropertyChanged "Alignment"
    
    If Ambient.UserMode Then
        On Error Resume Next
        AttachMessage Me, txtEdit.hwnd, WM_CONTEXTMENU
    End If
End Property

Private Sub txtEdit_Click()
    RaiseEvent Click
End Sub

Private Sub txtEdit_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
Attribute Locked.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Locked = txtEdit.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    txtEdit.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,RightToLeft
Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
Attribute RightToLeft.VB_ProcData.VB_Invoke_Property = ";Behavior"
Attribute RightToLeft.VB_UserMemId = -611
    RightToLeft = txtEdit.RightToLeft
End Property

Public Property Let RightToLeft(ByVal New_RightToLeft As Boolean)
    txtEdit.RightToLeft() = New_RightToLeft
    PropertyChanged "RightToLeft"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
Attribute Enabled.VB_ProcData.VB_Invoke_Property = ";Behavior"
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
    Call DrawDropDownSign
    If UserControl.Enabled Then txtEdit.ForeColor = m_ForeColor Else txtEdit.ForeColor = vbGrayText
End Property

Private Sub txtEdit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
Attribute MouseIcon.VB_ProcData.VB_Invoke_Property = ";Misc"
    Set MouseIcon = txtEdit.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set txtEdit.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Private Sub txtEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,MousePointer
Public Property Get MousePointer() As VBRUN.MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
Attribute MousePointer.VB_ProcData.VB_Invoke_Property = ";Misc"
    MousePointer = txtEdit.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As VBRUN.MousePointerConstants)
    txtEdit.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub txtEdit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkItem
Public Property Get LinkItem() As String
Attribute LinkItem.VB_Description = "Returns/sets the data passed to a destination control in a DDE conversation with another application."
Attribute LinkItem.VB_ProcData.VB_Invoke_Property = ";DDE"
    LinkItem = txtEdit.LinkItem
End Property

Public Property Let LinkItem(ByVal New_LinkItem As String)
    txtEdit.LinkItem() = New_LinkItem
    PropertyChanged "LinkItem"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkMode
Public Property Get LinkMode() As LinkModeConstants
Attribute LinkMode.VB_Description = "Returns/sets the type of link used for a DDE conversation and activates the connection."
Attribute LinkMode.VB_ProcData.VB_Invoke_Property = ";DDE"
    LinkMode = txtEdit.LinkMode
End Property

Public Property Let LinkMode(ByVal New_LinkMode As LinkModeConstants)
    txtEdit.LinkMode() = New_LinkMode
    PropertyChanged "LinkMode"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkTimeout
Public Property Get LinkTimeout() As Integer
Attribute LinkTimeout.VB_Description = "Returns/sets the amount of time a control waits for a response to a DDE message."
Attribute LinkTimeout.VB_ProcData.VB_Invoke_Property = ";DDE"
    LinkTimeout = txtEdit.LinkTimeout
End Property

Public Property Let LinkTimeout(ByVal New_LinkTimeout As Integer)
    txtEdit.LinkTimeout() = New_LinkTimeout
    PropertyChanged "LinkTimeout"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkTopic
Public Property Get LinkTopic() As String
Attribute LinkTopic.VB_Description = "Returns/sets the source application and topic for a destination control."
Attribute LinkTopic.VB_ProcData.VB_Invoke_Property = ";DDE"
    LinkTopic = txtEdit.LinkTopic
End Property

Public Property Let LinkTopic(ByVal New_LinkTopic As String)
    txtEdit.LinkTopic() = New_LinkTopic
    PropertyChanged "LinkTopic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkExecute
Public Sub LinkExecute(ByVal Command As String)
Attribute LinkExecute.VB_Description = "Sends a command string to the source application in a DDE conversation."
    txtEdit.LinkExecute Command
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkPoke
Public Sub LinkPoke()
Attribute LinkPoke.VB_Description = "Transfers contents of Label, PictureBox, or TextBox to source application in DDE conversation."
    txtEdit.LinkPoke
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkRequest
Public Sub LinkRequest()
Attribute LinkRequest.VB_Description = "Asks the source DDE application to update the contents of a Label, PictureBox, or Textbox control."
    txtEdit.LinkRequest
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=txtEdit,txtEdit,-1,LinkSend
Public Sub LinkSend()
Attribute LinkSend.VB_Description = "Transfers contents of PictureBox to destination application in DDE conversation."
    txtEdit.LinkSend
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get NumberFormat() As String
Attribute NumberFormat.VB_ProcData.VB_Invoke_Property = ";Behavior"
    NumberFormat = m_NumberFormat
End Property

Public Property Let NumberFormat(ByVal New_NumberFormat As String)
    m_NumberFormat = New_NumberFormat
    PropertyChanged "NumberFormat"
    
    txtEdit = CStr(Format(m_Value, m_NumberFormat))
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=picButton,picButton,-1,ForeColor
Public Property Get DropButtonForeColor() As OLE_COLOR
Attribute DropButtonForeColor.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DropButtonForeColor = m_DropButtonForeColor
End Property

Public Property Let DropButtonForeColor(ByVal New_DropButtonForeColor As OLE_COLOR)
    m_DropButtonForeColor = New_DropButtonForeColor
    PropertyChanged "DropButtonForeColor"
    
    picButton.ForeColor() = m_DropButtonForeColor
    Call DrawDropDownSign
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_ProcData.VB_Invoke_Property = ";Behavior"
    HotTracking = m_HotTracking
End Property

Public Property Let HotTracking(ByVal New_HotTracking As Boolean)
    m_HotTracking = New_HotTracking
    PropertyChanged "HotTracking"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get DropButtonForeColorOnMouse() As OLE_COLOR
Attribute DropButtonForeColorOnMouse.VB_ProcData.VB_Invoke_Property = ";Appearance"
    DropButtonForeColorOnMouse = m_DropButtonForeColorOnMouse
End Property

Public Property Let DropButtonForeColorOnMouse(ByVal New_DropButtonForeColorOnMouse As OLE_COLOR)
    m_DropButtonForeColorOnMouse = New_DropButtonForeColorOnMouse
    PropertyChanged "DropButtonForeColorOnMouse"
End Property

Private Sub DrawDropDownSign()
    Dim Rct As RECT
    
    With picButton
        'Clear the location of dropdown sign
        Call SetRect(Rct, 2, 2, .ScaleWidth / ScrX - 2, .ScaleHeight / ScrY - 2)
        Call FillRect(.hdc, Rct, CreateSolidBrush(GetBkColor(.hdc)))
        
        'Draw the drop down sign
        If UserControl.Enabled Then
            If m_HotTracking And IsMouseTracking Then
                .ForeColor = m_DropButtonForeColorOnMouse
            Else
                .ForeColor = m_DropButtonForeColor
            End If
            
            If Not IsDropButDown Then
                Call SetRect(Rct, 0, 0, .ScaleWidth / ScrX, .ScaleHeight / ScrY - 1)
                Call DrawText(.hdc, ByVal "6", 1, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            Else
                Call SetRect(Rct, 1, 1, .ScaleWidth / ScrX + 1, .ScaleHeight / ScrY)
                Call DrawText(.hdc, "6", 1, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            End If
        Else
            Call SetRect(Rct, 1, 1, .ScaleWidth / ScrX + 1, .ScaleHeight / ScrY)
            .ForeColor = vb3DHighlight
            Call DrawText(.hdc, "6", 1, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
            
            Call SetRect(Rct, 0, 0, .ScaleWidth / ScrX, .ScaleHeight / ScrY - 1)
            .ForeColor = vb3DShadow
            Call DrawText(.hdc, "6", 1, Rct, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE)
        End If
        
        .Refresh
    End With
End Sub

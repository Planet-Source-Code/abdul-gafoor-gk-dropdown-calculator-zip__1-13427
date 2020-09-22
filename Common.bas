Attribute VB_Name = "Common"
Option Explicit

Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_BOTTOM = &H8
Public Const BF_FLAT = &H4000      ' For flat rather than 3D borders
Public Const BF_LEFT = &H1
Public Const BF_MONO = &H8000      ' For monochrome borders.
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const EDGE_RAISED = BDR_RAISEDOUTER Or BDR_RAISEDINNER
Public Const EDGE_SUNKEN = BDR_SUNKENOUTER Or BDR_SUNKENINNER
Public Const BF_RECT = BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM

Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Const DT_CENTER = &H1
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20

Public Declare Function GetBkColor Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long

Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private mvarEditBox As TextBox
Private mvarButTxtCol As OLE_COLOR
Private mvarButTxtTrackCol As OLE_COLOR
Private mvarCalcButTracking As Boolean

'Holds a reference to edit box in user control
Public Property Get EditBox() As TextBox
    Set EditBox = mvarEditBox
End Property

Public Property Set EditBox(ByVal vNewValue As TextBox)
    Set mvarEditBox = vNewValue
End Property

'Holds the value of color of text in the calculator button
Public Property Get ButTxtCol() As OLE_COLOR
    ButTxtCol = mvarButTxtCol
End Property

Public Property Let ButTxtCol(ByVal vNewValue As OLE_COLOR)
    mvarButTxtCol = vNewValue
End Property

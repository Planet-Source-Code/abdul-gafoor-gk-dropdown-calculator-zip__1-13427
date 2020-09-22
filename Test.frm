VERSION 5.00
Object = "*\A..\Dropdown Calculator\DDCalc.vbp"
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Dropdown Calculator"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3600
   ForeColor       =   &H00C00000&
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3600
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   " Alignment "
      Height          =   1455
      Index           =   3
      Left            =   1920
      TabIndex        =   8
      Top             =   240
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "Right"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Center"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Format "
      Height          =   1095
      Index           =   2
      Left            =   1920
      TabIndex        =   7
      Top             =   1800
      Width           =   1455
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Enabled "
      Height          =   1095
      Index           =   1
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "False"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   13
         Top             =   720
         Width           =   735
      End
      Begin VB.OptionButton Option1 
         Caption         =   "True"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Appearance "
      Height          =   1095
      Index           =   0
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1455
      Begin VB.OptionButton Option1 
         Caption         =   "Flat"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         Caption         =   "3D"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
   Begin DDCalc.DDCalculator DDCalculator1 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumberFormat    =   "\R\s\. ##0.00"
      HotTracking     =   0   'False
      DropButtonForeColorOnMouse=   192
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Option1(0).Value = True
    Option1(2).Value = True
    Option1(5).Value = True
    
    Text1 = DDCalculator1.NumberFormat
End Sub

Private Sub Option1_Click(Index As Integer)
    With DDCalculator1
        Select Case Index
            Case 0: .Appearance = [3D]
            Case 1: .Appearance = Flat
            
            Case 2: .Alignment = [Left Justify]
            Case 3: .Alignment = Center
            Case 4: .Alignment = [Right Justify]
            
            Case 5: .Enabled = True
            Case 6: .Enabled = False
        End Select
    End With
End Sub

Private Sub Text1_Change()
    DDCalculator1.NumberFormat = Text1
End Sub

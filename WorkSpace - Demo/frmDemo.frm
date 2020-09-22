VERSION 5.00
Begin VB.Form frmDemo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "WorkSpace Class - Demo"
   ClientHeight    =   1590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   7
      Left            =   2018
      TabIndex        =   7
      Tag             =   "BottomCenter"
      Top             =   1320
      Width           =   195
   End
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Height          =   255
      Index           =   4
      Left            =   2018
      TabIndex        =   4
      Tag             =   "Center"
      Top             =   668
      Width           =   195
   End
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Height          =   195
      Index           =   1
      Left            =   2018
      TabIndex        =   1
      Tag             =   "TopCenter"
      Top             =   0
      Width           =   195
   End
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Caption         =   "Bottom Right"
      Height          =   255
      Index           =   8
      Left            =   3000
      TabIndex        =   8
      Tag             =   "BottomRight"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Caption         =   "Middle Right"
      Height          =   255
      Index           =   5
      Left            =   3030
      TabIndex        =   5
      Tag             =   "MiddleRight"
      Top             =   668
      Width           =   1185
   End
   Begin VB.OptionButton optPosition 
      Caption         =   "Bottom Left"
      Height          =   255
      Index           =   6
      Left            =   0
      TabIndex        =   6
      Tag             =   "BottomLeft"
      Top             =   1320
      Width           =   1155
   End
   Begin VB.OptionButton optPosition 
      Caption         =   "Middle Left"
      Height          =   255
      Index           =   3
      Left            =   0
      TabIndex        =   3
      Tag             =   "MiddleLeft"
      Top             =   668
      Width           =   1125
   End
   Begin VB.OptionButton optPosition 
      Alignment       =   1  'Right Justify
      Caption         =   "Top Right"
      Height          =   255
      Index           =   2
      Left            =   3210
      TabIndex        =   2
      Tag             =   "TopRight"
      Top             =   0
      Width           =   1005
   End
   Begin VB.OptionButton optPosition 
      Caption         =   "Top Left"
      Height          =   255
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Tag             =   "TopLeft"
      Top             =   0
      Width           =   945
   End
   Begin VB.Label lblCenter 
      Alignment       =   2  'Center
      Caption         =   "Cen     ter"
      Height          =   225
      Left            =   1680
      TabIndex        =   11
      Top             =   690
      Width           =   825
   End
   Begin VB.Label lblBottomCenter 
      Alignment       =   2  'Center
      Caption         =   "Bottom Center"
      Height          =   225
      Left            =   1620
      TabIndex        =   10
      Top             =   1140
      Width           =   1065
   End
   Begin VB.Label lblTopCenter 
      Alignment       =   2  'Center
      Caption         =   "Top Center"
      Height          =   225
      Left            =   1680
      TabIndex        =   9
      Top             =   210
      Width           =   885
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim wsWorkSpace As WorkSpace

Private Sub Form_Load()
    '** Create WorkSpace Class Object
        Set wsWorkSpace = New WorkSpace
    
    '** Set Default Option and Position
        optPosition(4).Value = True
        
    '** Inform user what he/she can do
        MsgBox "Hello!," & String(2, vbCrLf) & _
            vbTab & "This class gets the systems avaiable desktop workspace and provides property values to return the data." & vbCrLf & _
            "This is done entirerly with Win32 API calls and not with the SystemInfo Control." & String(2, vbCrLf) & _
            "Select an option to shift the form to that position on the screen." & String(2, vbCrLf) & "Enjoy!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '** Destroy WorkSpace Class Object
        Set wsWorkSpace = Nothing
End Sub

Private Sub lblTopCenter_Click()
    '** Set Top Center Position
        optPosition(1).Value = True
End Sub

Private Sub lblCenter_Click()
    '** Set Center Position
        optPosition(4).Value = True
End Sub

Private Sub lblBottomCenter_Click()
    '** Set Bottom Center Position
        optPosition(7).Value = True
End Sub

Private Sub optPosition_Click(Index As Integer)
    With wsWorkSpace
        Select Case Index
            Case 0  '** Set form position to Top Left
                .PositionForm Me, wsTop + wsLeft
            Case 1  '** Set form position to Top Center
                .PositionForm Me, wsTop + wsCenterX
            Case 2  '** Set form position to Top Right
                .PositionForm Me, wsTop + wsRight
            Case 3  '** Set form position to Middle Left
                .PositionForm Me, wsCenterY + wsLeft
            Case 4  '** Set form position to Center
                .PositionForm Me, wsCenterY + wsCenterX
            Case 5  '** Set form position to Middle Right
                .PositionForm Me, wsCenterY + wsRight
            Case 6  '** Set form position to Bottom Left
                .PositionForm Me, wsBottom + wsLeft
            Case 7  '** Set form position to Bottom Center
                .PositionForm Me, wsBottom + wsCenterX
            Case 8  '** Set form position to Bottom Right
                .PositionForm Me, wsBottom + wsRight
        End Select
    End With
End Sub

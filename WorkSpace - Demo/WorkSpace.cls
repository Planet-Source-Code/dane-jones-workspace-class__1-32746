VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WorkSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ********************************************* _
    Author: Dane Robert Jones _
    Email:  whitewolf@techie.com _
    Web:    http://danerjones.tripod.com _
    Disclaimer: The author is not responsible _
        for <yadda yadda>.. _
    Distribution of this code can be made freely _
        as long as this header and the original _
        code remains intact. _
  *********************************************
  
'** Define Constants
    Private Const SPI_GETWORKAREA = 48

'** Define Enumerations
    Public Enum wsPositions
        '** Vertical Positions
            wsTop = &H1
            wsCenterY = &H2
            wsBottom = &H4
        '** Horizontal Positions
            wsLeft = &H8
            wsCenterX = &H10
            wsRight = &H20
    End Enum

'** Define Type Structures
    Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
    End Type
    Private Type RECT_Twip
        Left As Single
        Top As Single
        Right As Single
        Bottom As Single
    End Type
    
'** Define Variables
    Private AvailableWorkSpace As RECT_Twip
    Private bolAutoRefresh As Boolean
    Private wsPosition As wsPositions
    
'** Define DLL Functions
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long

Private Sub Class_Initialize()
    '** Set Default Values
        bolAutoRefresh = False
        wsPosition = wsBottom + wsRight
        
    '** Get workspace dimentions from the system
        Call Refresh
End Sub

Public Sub Refresh()
    Dim WS As RECT
    Dim retVal As Long
    
    '** Get WorkSpace metrics from system
        retVal = SystemParametersInfo(SPI_GETWORKAREA, 0&, WS, 0&)
    
    '** Convert to twips and save
        With AvailableWorkSpace
            .Left = WS.Left * Screen.TwipsPerPixelX
            .Top = WS.Top * Screen.TwipsPerPixelY
            .Right = WS.Right * Screen.TwipsPerPixelX
            .Bottom = WS.Bottom * Screen.TwipsPerPixelY
        End With
    
End Sub

'** If autorefresh is true then refresh the workspace matrix
    Private Sub DoRefresh()
        If AutoRefresh Then Call Refresh
    End Sub

Friend Sub PositionForm(ByRef wsForm As Form, Optional ByVal wsPosition As wsPositions = 0)
    Dim bolTemoAutoRefresh As Boolean
    Dim sngLeft As Single, sngTop As Single
    
    '** Verify the supplied data
        If wsForm Is Nothing Then _
            Exit Sub
        If wsPosition = 0 Then _
            wsPosition = Position
        
    '** If the auto refresh feature is on then get a refresh _
        and temporarily disable the autorefresh so don't waste _
        processor cycles.
        bolTemoAutoRefresh = AutoRefresh
        
        If bolTemoAutoRefresh Then
            AutoRefresh = False
            Call Refresh
        End If
    
    '** Set the Forms Position
    '** Set Vertical Position
        Select Case True
            Case wsPosition And wsTop
                sngTop = Top
            Case wsPosition And wsCenterY
                sngTop = CenterY - wsForm.Height / 2
            Case wsPosition And wsBottom
                sngTop = Bottom - wsForm.Height
        End Select
        
    '** Set Horizontal Position
        Select Case True
            Case wsPosition And wsLeft
                sngLeft = Left
            Case wsPosition And wsCenterX
                sngLeft = CenterX - wsForm.Width / 2
            Case wsPosition And wsRight
                sngLeft = Right - wsForm.Width
        End Select
    
    '** Set Forms Position in one swoop to prevent screen flicker
        wsForm.Move sngLeft, sngTop
        
    '** Save Position
        Position = wsPosition
    
    '** Reset the AutoRefresh to prior state
        AutoRefresh = bolTemoAutoRefresh
End Sub

'** Position Property
    Friend Property Get Position() As wsPositions
        Position = wsPosition
    End Property

    Friend Property Let Position(ByVal vNewValue As wsPositions)
        wsPosition = vNewValue
    End Property

'** AutoRefresh Property
    Public Property Get AutoRefresh() As Boolean
        AutoRefresh = bolAutoRefresh
    End Property
    
    Public Property Let AutoRefresh(ByVal vNewValue As Boolean)
        bolAutoRefresh = vNewValue
    End Property

'** Left Property
    Public Property Get Left() As Single
        Call DoRefresh
        Left = AvailableWorkSpace.Left
    End Property
    
    Private Property Let Left(ByVal vNewValue As Single)
        AvailableWorkSpace.Left = vNewValue
    End Property

'** Top Property
    Public Property Get Top() As Single
        Call DoRefresh
        Top = AvailableWorkSpace.Top
    End Property
    
    Private Property Let Top(ByVal vNewValue As Single)
        AvailableWorkSpace.Top = vNewValue
    End Property

'** Right Property
    Public Property Get Right() As Single
        Call DoRefresh
        Right = AvailableWorkSpace.Right
    End Property
    
    Private Property Let Right(ByVal vNewValue As Single)
        AvailableWorkSpace.Right = vNewValue
    End Property

'** Bottom Property
    Public Property Get Bottom() As Single
        Call DoRefresh
        Bottom = AvailableWorkSpace.Bottom
    End Property
    
    Private Property Let Bottom(ByVal vNewValue As Single)
        AvailableWorkSpace.Bottom = vNewValue
    End Property

'** CenterX Property
    Public Property Get CenterX() As Single
        Call DoRefresh
        With AvailableWorkSpace
            CenterX = (.Right - .Left) / 2 + .Left
        End With
    End Property
    
'** CenterY Property
    Public Property Get CenterY() As Single
        Call DoRefresh
        With AvailableWorkSpace
            CenterY = (.Bottom - .Top) / 2 + .Top
        End With
    End Property

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'movement class
Public Enum Direction

    drLeft
    drRight
    drUp
    drDown
    
End Enum

Private drDirection As Direction

Private mintDistance As Integer

Public Property Let Distance(ByVal pintArg As Integer)

    mintDistance = pintArg
    
End Property

Public Property Get Distance() As Integer

    Distance = mintDistance
    
End Property

Public Property Get Direction() As Direction

    Direction = drDirection
    
End Property

Public Property Let Direction(ByVal drType As Direction)

    Select Case drType
    
        Case drLeft, drRight, drUp, drDown
        
        Case Else
            Err.Raise Number:=3, _
                Description:="Invalid direction"
    End Select
    
    drDirection = drType
    
End Property

Public Function ObjectMovement(Sprite As PictureBox)
        
    Select Case drDirection
    
        Case drLeft
            Sprite.Left = Sprite.Left - Distance
        Case drRight
            Sprite.Left = Sprite.Left + Distance
        Case drUp
            Sprite.Top = Sprite.Top - Distance
        Case drDown
            Sprite.Top = Sprite.Top + Distance

    End Select
    
End Function

Public Sub InitObjectMovement(ByVal PicBar As PictureBox)

    With PicBar
    
        .ScaleMode = vbPixels
        .AutoRedraw = True
        
    End With
    
End Sub

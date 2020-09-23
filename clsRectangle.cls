VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRectangle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mintLength As Integer
Private mintWidth As Integer


Public Property Get area()
    
    area = Length * Width

End Property

Public Property Get perimeter()
    
    perimeter = Length * 2 + Width * 2

End Property

Public Property Get Length() As Integer

    Length = mintLength
    
End Property

Public Property Get Width() As Integer

    Width = mintWidth
    
End Property

Public Property Let Length(ByVal pintArg As Integer)

    mintLength = pintArg
    
End Property

Public Property Let Width(ByVal pintArg As Integer)

    mintWidth = pintArg
    
End Property

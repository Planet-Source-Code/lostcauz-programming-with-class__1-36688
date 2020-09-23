VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarValue As Integer
Private mvarForeColor As Long
Private mvarBackColor As Long
Private mvarDrawText As Boolean

Public Event ProgressDone()


Public Property Let BackColor(ByVal vData As Double)

    mvarBackColor = vData
    
End Property


Public Property Get BackColor() As Double

    mvarBackColor = BackColor
    
End Property


Public Property Let ForeColor(ByVal vData As Double)

    mvarForeColor = vData
    
End Property


Public Property Get ForeColor() As Double

    mvarForeColor = ForeColor
    
End Property


Public Property Let value(ByVal vData As Integer)

    mvarValue = vData
    
End Property


Public Property Get value() As Integer

    value = mvarValue
    
End Property


Public Property Get DrawText() As Boolean

    DrawText = mvarDrawText
    
End Property


Public Property Let DrawText(ByVal vData As Boolean)

    mvarDrawText = vData
    
End Property


Private Sub Class_Initialize()

    mvarValue = 0
    mvarBackColor = &H8000000F
    mvarForeColor = &HFF0000
    mvarDrawText = True
    
End Sub


Public Sub InitProgress(ByVal PicBar As PictureBox)

    With PicBar
    
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .FontBold = True
        .BackColor = mvarBackColor
        .ForeColor = mvarForeColor
        
    End With
    
End Sub


Public Sub DrawStatus(ByVal PicBar As PictureBox)

    PicBar.Cls

    If mvarDrawText Then
        PicBar.CurrentX = (PicBar.ScaleWidth / 2) - (PicBar.TextWidth(CInt((PicBar.TextWidth(mvarValue / PicBar.Width) * 100))) / 2)
        PicBar.CurrentY = (PicBar.ScaleHeight / 2) - (PicBar.TextHeight("1") / 2)
    End If

    If mvarValue >= PicBar.ScaleWidth Then
        RaiseEvent ProgressDone

        If mvarDrawText Then
            'PicBar.Print "100%"
            Form1.Label9.Caption = "100%" 'added this to remedy blackout of percentage
        End If
    Else

        If mvarDrawText Then
            'PicBar.Print CStr(Round(((mvarValue / PicBar.ScaleWidth) * 100), 0) & "%")
            Form1.Label9.Caption = CStr(Round(((mvarValue / PicBar.ScaleWidth) * 100), 0) & "%")
        End If
    End If
    
    PicBar.Line (0, 0)-((mvarValue / PicBar.ScaleWidth) * PicBar.ScaleWidth, PicBar.ScaleHeight), mvarForeColor Xor mvarBackColor, BF
    PicBar.Refresh
    
End Sub

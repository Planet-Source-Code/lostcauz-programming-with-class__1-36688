VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Collection of Classes"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   7755
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDown 
      Caption         =   "D"
      Height          =   255
      Left            =   3840
      TabIndex        =   19
      Top             =   960
      Width           =   255
   End
   Begin VB.CommandButton cmdUp 
      Caption         =   "U"
      Height          =   255
      Left            =   3840
      TabIndex        =   18
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton cmdRight 
      Caption         =   "R"
      Height          =   255
      Left            =   4080
      TabIndex        =   17
      Top             =   720
      Width           =   255
   End
   Begin VB.CommandButton cmdLeft 
      Caption         =   "L"
      Height          =   255
      Left            =   3600
      TabIndex        =   16
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FF00FF&
      Height          =   375
      Left            =   5640
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   15
      Top             =   840
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   3600
      Top             =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3600
      ScaleHeight     =   315
      ScaleWidth      =   3675
      TabIndex        =   12
      Top             =   2640
      Width           =   3735
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFC0C0&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   40
         Width           =   3735
      End
   End
   Begin VB.CommandButton cmdRectangle 
      Caption         =   "Calculate"
      Height          =   495
      Left            =   1920
      TabIndex        =   10
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   9
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Height          =   495
      Left            =   2640
      TabIndex        =   2
      Top             =   600
      Width           =   495
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   3120
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   7560
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label8 
      Caption         =   "Movement Class"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   120
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   3360
      X2              =   3360
      Y1              =   120
      Y2              =   3120
   End
   Begin VB.Label Label7 
      Caption         =   "Progress Bar Class"
      Height          =   255
      Left            =   3600
      TabIndex        =   11
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Width"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Length"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label4 
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "Rectangle Class"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1320
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "StopWatch Class"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   2295
   End
   Begin VB.Label Label1 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'the stopwatch and progressbar class were used from other
'PSC submissions with only slight modifications
'I wrote the other 2 classes for this example
'hopefully will assist in learning classes
'feedback is appreciated - lostcauz

'stopwatch class implementation
Dim randy As New clsStopWatch
    
'rectangle class implementation
Dim rowdy As New clsRectangle

'ProgressBar class implementation
Private WithEvents progress As clsProgressBar
Attribute progress.VB_VarHelpID = -1

'movement class implementation
Dim demi As New clsMovement

Private Sub cmdDown_Click() 'movement

    demi.Direction = drDown
    demi.ObjectMovement Picture2

End Sub

Private Sub cmdLeft_Click() 'movement

    demi.Direction = drLeft
    demi.ObjectMovement Picture2

End Sub

Private Sub cmdRectangle_Click() 'rectangle class

    rowdy.Length = Val(Text1.Text)
    rowdy.Width = Val(Text2.Text)
    
    Label4.Caption = "Area is: " & rowdy.area & Chr(13) _
                    & "Perimeter is: " & rowdy.perimeter
    
End Sub

Private Sub cmdRight_Click() 'movement

    demi.Direction = drRight
    demi.ObjectMovement Picture2

End Sub

Private Sub cmdStart_Click() 'stopwatch class
    
    randy.StartTimer
    
End Sub

Private Sub cmdStop_Click() 'stopwatch class
    
    randy.StopTimer
    Label1.Caption = "Exactly " & randy.ElapsedMilliseconds _
                        & " milliseconds" & Chr(13) _
                        & "or about " & randy.ElapsedSeconds _
                        & " seconds"
    'end stopwatch
    
End Sub


Private Sub cmdUp_Click() 'movement

    demi.Direction = drUp
    demi.ObjectMovement Picture2

End Sub

Private Sub progress_ProgressDone() 'ProgressBar

    Timer1.Enabled = False
    
End Sub


Private Sub Timer1_Timer() 'ProgressBar

    progress.value = progress.value + 10
    progress.DrawStatus Picture1
    
End Sub

Private Sub Form_Load()

    Set progress = New clsProgressBar 'ProgressBar
    
    progress.BackColor = vbBlack
    progress.ForeColor = vbRed
    progress.InitProgress Picture1    'end ProgressBar
    
    demi.InitObjectMovement Picture2 'movement
    demi.Distance = 120 'movement
    
End Sub

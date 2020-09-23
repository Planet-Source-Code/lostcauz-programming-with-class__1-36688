VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
    '-- Local private variables
    Private m_lStartTime As Long
    Private m_lEndTime As Long
    Private m_lLastLapTime As Long


Public Sub StopTimer()
    On Error GoTo StopTimer_Error
    m_lEndTime = GetTickCount()

StopTimer_Error:
End Sub


Public Sub StartTimer()
    On Error GoTo StartTimer_Error
    
    Dim lStoppedTime As Long
    
    '-- If there is an endtime, we need to calculate how much time
    ' has elapsed since it was stopped and adjust the start time
    ' and last lap time accordingly. We don't want to
    ' include time that passed while the watch was stopped.

    If m_lEndTime > 0 Then
        
        '-- How long were we stopped?
        lStoppedTime = GetTickCount() - m_lEndTime
        
        '-- Adjust the start time.
        m_lStartTime = m_lStartTime + lStoppedTime
        
        '-- Adjust the LapTime.
        m_lLastLapTime = m_lLastLapTime + lStoppedTime
        
    Else
        
        '-- First time we've started. Just captu
        '     re the start time.
        m_lStartTime = GetTickCount()
        
    End If
    
    '-- Clear the endtime.
    m_lEndTime = 0

StartTimer_Error:

End Sub


Public Property Get ElapsedMilliseconds() As Long

    On Error GoTo ElapsedMilliseconds_Error

    If m_lStartTime = 0 Then
        '-- The timer hasn't started yet. Return
        '     0.
        ElapsedMilliseconds = 0
        Exit Property
    End If

    If m_lEndTime = 0 Then
        '-- The user has not clicked stop yet. G
        '     ive an elapsed time.
        ElapsedMilliseconds = GetTickCount() - m_lStartTime
    Else
        '-- There is a stop time. Just calculate
        '     the difference.
        ElapsedMilliseconds = m_lEndTime - m_lStartTime
    End If

ElapsedMilliseconds_Error:

End Property


Public Property Get laptime() As Long
    '-- Return the number of seconds since t
    '     he last LapTime.
    On Error GoTo Laptime_Error
    
    Dim lCurrentLapTime As Long
    Dim lRetVal As Long
    
    lCurrentLapTime = Me.ElapsedMilliseconds
    


    If m_lLastLapTime = 0 Then
        '-- First Lap. Just return the Elapsed M
        '     illiseconds.
        lRetVal = lCurrentLapTime
    Else
        lRetVal = lCurrentLapTime - m_lLastLapTime
    End If
    
    '-- Save the last lap time.
    m_lLastLapTime = lCurrentLapTime
    
    '-- Return the lap time.
    laptime = lRetVal

Laptime_Error:

End Property

Public Property Get ElapsedSeconds() As Long
    
    If m_lStartTime = 0 Then
        '-- The timer hasn't started yet. Return
        '     0.
        ElapsedSeconds = 0
        Exit Property
    End If

    If m_lEndTime = 0 Then
    
        '-- The user has not clicked stop yet. Give an elapsed time.
        ElapsedSeconds = ((GetTickCount() - m_lStartTime) * 0.001)
    Else
        '-- There is a stop time. Just calculate the difference.
        ElapsedSeconds = ((m_lEndTime - m_lStartTime) * 0.001)
    End If

End Property


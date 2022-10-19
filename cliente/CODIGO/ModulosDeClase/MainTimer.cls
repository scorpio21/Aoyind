VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

''
' The MainTimer of the game.
' Controls client´s intervals.
'
' @author Nacho (Integer) integer-x@hotmail.com
' @version 1.0
' @date 20060428

'*****************************************************
'Timer class for Argentum Online
'Designed and implemented by Nacho (Integer)
'integer-x@hotmail.com
'*****************************************************

Option Explicit

''
' How many timers we are going to use-
'
' @see See MainTimer.CreateAll
Private Const CANTIDADTIMERS As Byte = 19

''
' A Timer data structure.
'
' @param Interval How long, in miliseconds, a cicle lasts.
' @param CurrentTick Current Tick in which the Timer is.
' @param StartTick Tick in which current cicle has started.
' @param Run True if the timer is active.

Private Type Timer

    Interval As Long
    CurrentTick As Long
    StartTick As Long
    Run As Boolean

End Type

'Timers
Dim Timer(1 To CANTIDADTIMERS) As Timer

''
' Timer´s Index.
'
' @param Attack                 Controls the Combat system.
' @param Work                   Controls the Work system.
' @param UseItemWithU           Controls the usage of items with the "U" key.
' @param UseItemWithDblClick    Controls the usage of items with double click.
' @param SendRPU                Controls the use of the "L" to request a pos update.
' @param CastSpell              Controls the casting of spells.
' @param Arrows                 Controls the shooting of arrows.
Public Enum TimersIndex

    PuedeGolpe = 1
    Work = 2
    PuedeUsar = 3
    PuedeUsarDobleClick = 4
    SendRPU = 5
    PuedeLanzarHechizo = 6
    PuedeFlechas = 7
    PuedeGolpeMagia = 8
    Hide = 9
    Buy = 10
    Montar = 11
    Anclar = 12
    Telep = 13
    HabilitaLanzarHechizo = 14
    PuedeMagiaGolpe = 15
    PuedeGolpeUsar = 16
    PuedeMover = 17
    PuedeMoverEquitando = 18
    PuedeRPUMover = 19

End Enum

''
' Window´s API Function.
' A milisecond pricision counter.
'
' @return   Miliseconds since midnight.

Private Declare Function GetTickCount Lib "kernel32" () As Long

''
' Sets a new interval for a timer.
'
' @param TimerIndex Timer´s Index
' @param Interval New lenght for the Timer´s cicle in miliseconds.
' @remarks  Must be done after creating the timer and before using it, otherwise, Interval will be 0

Public Sub SetInterval(ByVal TimerIndex As TimersIndex, ByVal Interval As Long)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified:
    'Desc: Sets a new interval for a timer.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Interval = Interval

End Sub

''
' Retrieves the interval of a timer.
'
' @param TimerIndex Timer´s Index
'
' @return The requested timer's interval.

Public Function GetInterval(ByVal TimerIndex As TimersIndex) As Long

    '*************************************************
    'Author: Juan Martín Sotuyo Dodero (Maraxus)
    'Last modified: 13/10/2008
    'Desc: Retrieves the interval of a timer.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    GetInterval = Timer(TimerIndex).Interval

End Function

''
' Starts the timer.
'
' @param TimerIndex Timer´s Index
' @remarks  Do this for starting the timer.

Public Sub Start(ByVal TimerIndex As TimersIndex)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified:
    'Desc: Starts running the Timer.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.

End Sub

''
' Resume the timer form where we had stoped.
'
' @param TimerIndex Timer´s Index
' @remarks  Makes the Timer continue from where it had stoped.

Public Sub Continue(ByVal TimerIndex As TimersIndex)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified:
    'Desc: Continue our timer from where we had stopped it.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = True 'We allow the Timer to run.
    
    'We´ll set the StartTick again
    Timer(TimerIndex).StartTick = Timer(TimerIndex).StartTick + Timer(TimerIndex).CurrentTick

End Sub

''
' Pauses the Timer
'
' @param TimerIndex Timer´s Index
' @remarks  Can be resumed with Continue()

Public Sub Pause(ByVal TimerIndex As TimersIndex)

    '*************************************************
    'Author: Nacho (Integer)
    'Last modified:
    'Desc: Pauses the timer.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).Run = False 'Don´t allow the timer to run.
    
    Timer(TimerIndex).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimerIndex).StartTick 'Calculates acutual tick

End Sub

''
' Check if the timer has already completed it´s cicle.
'
' @param TimerIndex Timer´s Index
' @param Restart If true, restart if we timer has ticked
' @return True if the interval has already been elapsed
' @remarks  Can´t be done if the timer is stoped or if it had never been started.

Public Function Check(ByVal TimerIndex As TimersIndex, _
                      Optional Restart As Boolean = False) As Boolean

    '*************************************************
    'Author: Nacho Agustin (Integer)
    'Last modified: 08/26/06
    'Modification: NIGO: Added Restart as boolean
    'Desc: Checks if the Timer has alredy "ticked"
    'Returns: True if it has ticked, False if not.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    If Not Timer(TimerIndex).Run Then Exit Function 'If we weren´t allowed to run, exit function
    
    Select Case TimerIndex

        Case TimersIndex.PuedeGolpe
            Timer(TimersIndex.PuedeGolpe).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeGolpe).StartTick 'Calcutates CurrentTick
                
            If Abs(Timer(TimersIndex.PuedeGolpe).CurrentTick) >= Timer(TimersIndex.PuedeGolpe).Interval Then
                Timer(TimersIndex.PuedeMagiaGolpe).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeMagiaGolpe).StartTick 'Calcutates CurrentTick

                If Abs(Timer(TimersIndex.PuedeMagiaGolpe).CurrentTick) >= Timer(TimersIndex.PuedeMagiaGolpe).Interval Then
                    Check = True
                    Timer(TimersIndex.PuedeGolpe).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Timer(TimersIndex.PuedeGolpeMagia).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Exit Function

                End If

            End If

        Case TimersIndex.PuedeLanzarHechizo
            Timer(TimersIndex.PuedeLanzarHechizo).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeLanzarHechizo).StartTick 'Calcutates CurrentTick
                
            If Abs(Timer(TimersIndex.PuedeLanzarHechizo).CurrentTick) >= Timer(TimersIndex.PuedeLanzarHechizo).Interval Then
                Timer(TimersIndex.PuedeGolpeMagia).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeGolpeMagia).StartTick 'Calcutates CurrentTick

                If Abs(Timer(TimersIndex.PuedeGolpeMagia).CurrentTick) >= Timer(TimersIndex.PuedeGolpeMagia).Interval Then
                    Check = True
                    Timer(TimersIndex.PuedeLanzarHechizo).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Timer(TimersIndex.PuedeMagiaGolpe).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Timer(TimersIndex.HabilitaLanzarHechizo).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Exit Function

                End If

            End If

        Case TimersIndex.PuedeMover

            If UserEquitando = True Then
                Timer(TimersIndex.PuedeMoverEquitando).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeMoverEquitando).StartTick

                If Abs(Timer(TimersIndex.PuedeMoverEquitando).CurrentTick) >= Timer(TimersIndex.PuedeMoverEquitando).Interval Then
                    Check = True
                    Timer(TimersIndex.PuedeMoverEquitando).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Exit Function

                End If

            Else
                Timer(TimersIndex.PuedeMover).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimersIndex.PuedeMover).StartTick

                If Abs(Timer(TimersIndex.PuedeMover).CurrentTick) >= Timer(TimersIndex.PuedeMover).Interval Then
                    Check = True
                    Timer(TimersIndex.PuedeMover).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer
                    Exit Function

                End If

            End If

        Case Else
            Timer(TimerIndex).CurrentTick = (GetTickCount() And &H7FFFFFFF) - Timer(TimerIndex).StartTick 'Calcutates CurrentTick
            
            If Abs(Timer(TimerIndex).CurrentTick) >= Timer(TimerIndex).Interval Then
                Check = True
                Timer(TimerIndex).StartTick = (GetTickCount() And &H7FFFFFFF) 'Restart Timer (Nicer than calling Restart() )
                Exit Function

            End If

    End Select
    
End Function

Public Function CheckV(ByVal TimerIndex As TimersIndex) As Boolean

    '*************************************************
    'Author: Nacho Agustin (Integer)
    'Last modified: 08/26/06
    'Modification: NIGO: Added Restart as boolean
    'Desc: Checks if the Timer has alredy "ticked"
    'Returns: True if it has ticked, False if not.
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Function
    
    If Not Timer(TimerIndex).Run Then Exit Function 'If we weren´t allowed to run, exit function

    If Abs((GetTickCount() And &H7FFFFFFF) - Timer(TimerIndex).StartTick) >= Timer(TimerIndex).Interval Then
        CheckV = True 'We have Ticked!

    End If

End Function

''
' Restarts timer.
'
' @param TimerIndex Timer´s Index

Public Sub Restart(ByVal TimerIndex As TimersIndex)

    '*************************************************
    'Author: Nacho Agustin (Integer)
    'Last modified:
    'Desc: Restarts timer
    '*************************************************
    If TimerIndex < 1 Or TimerIndex > CANTIDADTIMERS Then Exit Sub
    
    Timer(TimerIndex).StartTick = (GetTickCount() And &H7FFFFFFF)

End Sub

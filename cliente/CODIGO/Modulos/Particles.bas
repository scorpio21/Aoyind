Attribute VB_Name = "Particles"
Option Explicit

Public Type TLVERTEX2

        x As Single
        y As Single
        z As Single
        rhw As Single
        color As Long
        tu As Single
        tv As Single

End Type

'Blood list

Public Type BloodData

        v(0 To 5) As TLVERTEX2
        Life As Long
        TileX As Byte
        TileY As Byte

End Type

Public LastBlood        As Long
Public BloodList()      As BloodData


Public LastDamage       As Integer    'Last damage counter text index used
Dim z As Integer
Public SangreTexture(1 To 4) As Direct3DTexture8
Private Type Effect
    x As Single                 'Location of effect
    y As Single
    GoToX As Single             'Location to move to
    GoToY As Single
    KillWhenAtTarget As Boolean     'If the effect is at its target (GoToX/Y), then Progression is set to 0
    KillWhenTargetLost As Boolean   'Kill the effect if the target is lost (sets progression = 0)
    Gfx As Byte                 'Particle texture used
    used As Boolean             'If the effect is in use
    EffectNum As Byte           'What number of effect that is used
    Modifier As Integer         'Misc variable (depends on the effect)
    FloatSize As Long           'The size of the particles
    Direction As Integer        'Misc variable (depends on the effect)
    Particles() As Particle     'Information on each particle
    Progression As Single       'Progression state, best to design where 0 = effect ends
    PartVertex() As TLVERTEX    'Used to point render particles
    PreviousFrame As Long       'Tick time of the last frame
    ParticleCount As Integer    'Number of particles total
    ParticlesLeft As Integer    'Number of particles left - only for non-repetitive effects
    BindToChar As Integer       'Setting this value will bind the effect to move towards the character
    BindSpeed As Single         'How fast the effect moves towards the character
    BoundToMap As Byte          'If the effect is bound to the map or not (used only by the map editor)
    TargetAA As Single
    R As Single
End Type
Public NumEffects As Byte   'Maximum number of effects at once
Public Effect() As Effect   'List of all the active effects
 
'Constants With The Order Number For Each Effect
Public Const EffectNum_Fire As Byte = 1             'Burn baby, burn! Flame from a central point that blows in a specified direction
Public Const EffectNum_Snow As Byte = 2             'Snow that covers the screen - weather effect
Public Const EffectNum_Heal As Byte = 3             'Healing effect that can bind to a character, ankhs float up and fade
Public Const EffectNum_Bless As Byte = 4            'Following three effects are same: create a circle around the central point
Public Const EffectNum_Protection As Byte = 5       ' (often the character) and makes the given particle on the perimeter
Public Const EffectNum_Strengthen As Byte = 6       ' which float up and fade out
Public Const EffectNum_Rain As Byte = 7             'Exact same as snow, but moves much faster and more alpha value - weather effect
Public Const EffectNum_EquationTemplate As Byte = 8 'Template for creating particle effects through equations - a page with some equations can be found here: [url=http://www.vbgore.com/modules.php?name=Forums&file=viewtopic&t=221]http://www.vbgore.com/modules.php?name= ... opic&t=221[/url]
Public Const EffectNum_Waterfall As Byte = 9        'Waterfall effect
Public Const EffectNum_Summon As Byte = 10          'Summon effect
Public Const EffectNum_Necro As Byte = 11
Public Const EffectNum_Atom As Byte = 12
Public Const EffectNum_MeditMAX As Byte = 13
Public Const EffectNum_PortalGroso As Byte = 14
Public Const EffectNum_RedFountain As Byte = 15
Public Const EffectNum_Smoke As Byte = 16
Public Const EffectNum_BloodSpray As Byte = 17
Public Const EffectNum_BloodSplatter As Byte = 18


Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
 
Function Effect_EquationTemplate_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'Particle effect template for effects as described on the
'wiki page: [url=http://www.vbgore.com/Particle_effect_equations]http://www.vbgore.com/Particle_effect_equations[/url]
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Begin]http://www.vbgore.com/CommonCode.Partic ... late_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_EquationTemplate_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_EquationTemplate  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True                     'Enable the effect
    Effect(EffectIndex).x = x                           'Set the effect's X coordinate
    Effect(EffectIndex).y = y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_EquationTemplate_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_EquationTemplate_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Reset]http://www.vbgore.com/CommonCode.Partic ... late_Reset[/url]
'*****************************************************************
Dim x As Single
Dim y As Single
Dim R As Single
    
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    R = (Index / 20) * exp(Index / Effect(EffectIndex).Progression Mod 3)
    x = R * Cos(Index)
    y = R * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + x, Effect(EffectIndex).y + y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 1, 0.2 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_EquationTemplate_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_EquationTemplate_Update]http://www.vbgore.com/CommonCode.Partic ... ate_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_EquationTemplate_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Function Effect_Bless_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Begin]http://www.vbgore.com/CommonCode.Partic ... less_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Bless_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Bless     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Bless_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Bless_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Reset]http://www.vbgore.com/CommonCode.Partic ... less_Reset[/url]
'*****************************************************************
Dim a As Single
Dim x As Single
Dim y As Single
 
    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    x = Effect(EffectIndex).x - (Cos(a) * Effect(EffectIndex).Modifier)
    y = Effect(EffectIndex).y + (Sin(a) * Effect(EffectIndex).Modifier)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt x, y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0, 5, 5, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    
 
 
End Sub
 
Private Sub Effect_Bless_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Bless_Update]http://www.vbgore.com/CommonCode.Partic ... ess_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Bless_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Function Effect_Fire_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin]http://www.vbgore.com/CommonCode.Partic ... Fire_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Fire_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Fire      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).x = x           'Set the effect's X coordinate
    Effect(EffectIndex).y = y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Fire_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Fire_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset]http://www.vbgore.com/CommonCode.Partic ... Fire_Reset[/url]
'*****************************************************************
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x - 10 + Rnd * 20, Effect(EffectIndex).y - 10 + Rnd * 20, -Sin((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, Cos((Effect(EffectIndex).Direction + (Rnd * 70) - 35) * DegreeToRadian) * 8, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.4 + (Rnd * 0.2), 0.03 + (Rnd * 0.07)
 
End Sub
 
Private Sub Effect_Fire_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update]http://www.vbgore.com/CommonCode.Partic ... ire_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then
 
                    'Reset the particle
                    Effect_Fire_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Private Function Effect_FToDW(f As Single) As Long
'*****************************************************************
'Converts a float to a D-Word, or in Visual Basic terms, a Single to a Long
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_FToDW]http://www.vbgore.com/CommonCode.Particles.Effect_FToDW[/url]
'*****************************************************************
Dim buf As D3DXBuffer
 
    'Converts a single into a long (Float to DWORD)
    Set buf = D3DX.CreateBuffer(4)
    D3DX.BufferSetData buf, 0, 4, 1, f
    D3DX.BufferGetData buf, 0, 4, 1, Effect_FToDW
 
End Function
 
Function Effect_Heal_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Begin]http://www.vbgore.com/CommonCode.Partic ... Heal_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Heal_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Heal      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).x = x           'Set the effect's X coordinate
    Effect(EffectIndex).y = y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Progression = Progression   'Loop the effect
    Effect(EffectIndex).KillWhenAtTarget = True     'End the effect when it reaches the target (progression = 0)
    Effect(EffectIndex).KillWhenTargetLost = True   'End the effect if the target is lost (progression = 0)
    
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Heal_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Heal_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Reset]http://www.vbgore.com/CommonCode.Partic ... Heal_Reset[/url]
'*****************************************************************
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x - 10 + Rnd * 20, Effect(EffectIndex).y - 10 + Rnd * 20, -Sin((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), Cos((180 + (Rnd * 90) - 45) * 0.0174533) * 8 + (Rnd * 3), 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0.8, 0.2, 0.2, 0.6 + (Rnd * 0.2), 0.01 + (Rnd * 0.5)
    
End Sub
 
Private Sub Effect_Heal_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Heal_Update]http://www.vbgore.com/CommonCode.Partic ... eal_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
Dim I As Integer
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression <> 0 Then
 
                    'Reset the particle
                    Effect_Heal_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Sub Effect_Kill(ByVal EffectIndex As Integer, Optional ByVal KillAll As Boolean = False)
'*****************************************************************
'Kills (stops) a single effect or all effects
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Kill]http://www.vbgore.com/CommonCode.Particles.Effect_Kill[/url]
'*****************************************************************
Dim LoopC As Long
 
    'Check If To Kill All Effects
    If KillAll = True Then
 
        'Loop Through Every Effect
        For LoopC = 1 To NumEffects
 
            'Stop The Effect
            Effect(LoopC).used = False
 
        Next
        
    Else
 
        'Stop The Selected Effect
        Effect(EffectIndex).used = False
        
    End If
 
End Sub
 
Private Function Effect_NextOpenSlot() As Integer
'*****************************************************************
'Finds the next open effects index
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_NextOpenSlot]http://www.vbgore.com/CommonCode.Partic ... xtOpenSlot[/url]
'*****************************************************************
Dim EffectIndex As Integer
 
    'Find The Next Open Effect Slot
    Do
        EffectIndex = EffectIndex + 1   'Check The Next Slot
        If EffectIndex > NumEffects Then    'Dont Go Over Maximum Amount
            Effect_NextOpenSlot = -1
            Exit Function
        End If
    Loop While Effect(EffectIndex).used = True    'Check Next If Effect Is In Use
 
    'Return the next open slot
    Effect_NextOpenSlot = EffectIndex
 
    'Clear the old information from the effect
    Erase Effect(EffectIndex).Particles()
    Erase Effect(EffectIndex).PartVertex()
    ZeroMemory Effect(EffectIndex), LenB(Effect(EffectIndex))
    Effect(EffectIndex).GoToX = -30000
    Effect(EffectIndex).GoToY = -30000
 
End Function
 
Function Effect_Protection_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Begin]http://www.vbgore.com/CommonCode.Partic ... tion_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Protection_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Protection    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Protection_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Protection_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Reset]http://www.vbgore.com/CommonCode.Partic ... tion_Reset[/url]
'*****************************************************************
Dim a As Single
Dim x As Single
Dim y As Single
 
    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier)
    y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt x, y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_UpdateOffset(ByVal EffectIndex As Integer)
'***************************************************
'Update an effect's position if the screen has moved
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_UpdateOffset]http://www.vbgore.com/CommonCode.Partic ... dateOffset[/url]
'***************************************************
 
    Effect(EffectIndex).x = Effect(EffectIndex).x + (LastOffsetX - ParticleOffsetX)
    Effect(EffectIndex).y = Effect(EffectIndex).y + (LastOffsetY - ParticleOffsetY)
 
End Sub
 
Private Sub Effect_UpdateBinding(ByVal EffectIndex As Integer)
 
'***************************************************
'Updates the binding of a particle effect to a target, if
'the effect is bound to a character
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_UpdateBinding]http://www.vbgore.com/CommonCode.Partic ... ateBinding[/url]
'***************************************************
Dim TargetI As Integer
Dim TargetA As Single
 
    'Update position through character binding
    If Effect(EffectIndex).BindToChar > 0 Then
 
        'Store the character index
        TargetI = Effect(EffectIndex).BindToChar
 
        'Check for a valid binding index
        If TargetI > LastChar Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        ElseIf charlist(TargetI).ACTIVE = 0 Then
            Effect(EffectIndex).BindToChar = 0
            If Effect(EffectIndex).KillWhenTargetLost Then
                Effect(EffectIndex).Progression = 0
                Exit Sub
            End If
        Else
 
            'Calculate the X and Y positions
            Effect(EffectIndex).GoToX = Engine_TPtoSPX(charlist(Effect(EffectIndex).BindToChar).Pos.x) + 16
            Effect(EffectIndex).GoToY = Engine_TPtoSPY(charlist(Effect(EffectIndex).BindToChar).Pos.y)
 
        End If
 
    End If
 
    'Move to the new position if needed
    If Effect(EffectIndex).GoToX > -30000 Or Effect(EffectIndex).GoToY > -30000 Then
        If Effect(EffectIndex).GoToX <> Effect(EffectIndex).x Or Effect(EffectIndex).GoToY <> Effect(EffectIndex).y Then
 
            'Calculate the angle
            TargetA = Engine_GetAngle(Effect(EffectIndex).x, Effect(EffectIndex).y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
 
            'Update the position of the effect
            Effect(EffectIndex).x = Effect(EffectIndex).x - Sin(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
            Effect(EffectIndex).y = Effect(EffectIndex).y + Cos(TargetA * DegreeToRadian) * Effect(EffectIndex).BindSpeed
 
            'Check if the effect is close enough to the target to just stick it at the target
            If Effect(EffectIndex).GoToX > -30000 Then
                If Abs(Effect(EffectIndex).x - Effect(EffectIndex).GoToX) < 6 Then Effect(EffectIndex).x = Effect(EffectIndex).GoToX
            End If
            If Effect(EffectIndex).GoToY > -30000 Then
                If Abs(Effect(EffectIndex).y - Effect(EffectIndex).GoToY) < 6 Then Effect(EffectIndex).y = Effect(EffectIndex).GoToY
            End If
 
            'Check if the position of the effect is equal to that of the target
            If Effect(EffectIndex).x = Effect(EffectIndex).GoToX Then
                If Effect(EffectIndex).y = Effect(EffectIndex).GoToY Then
 
                    'For some effects, if the position is reached, we want to end the effect
                    If Effect(EffectIndex).KillWhenAtTarget Then
                        Effect(EffectIndex).BindToChar = 0
                        Effect(EffectIndex).Progression = 0
                        Effect(EffectIndex).GoToX = Effect(EffectIndex).x
                        Effect(EffectIndex).GoToY = Effect(EffectIndex).y
                    End If
                    Exit Sub    'The effect is at the right position, don't update
 
                End If
            End If
 
        End If
    End If
 
End Sub
 
Private Sub Effect_Protection_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Protection_Update]http://www.vbgore.com/CommonCode.Partic ... ion_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Protection_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Public Sub Effect_Render(ByVal EffectIndex As Integer, Optional ByVal SetRenderStates As Boolean = True)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Render]http://www.vbgore.com/CommonCode.Partic ... ect_Render[/url]
'*****************************************************************
 
   Dim count As Long
    Dim I As Long

    'Check if we have the device

    If D3DDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub
    'Set the render state for the size of the particle
    D3DDevice.SetRenderState D3DRS_POINTSIZE, Effect(EffectIndex).FloatSize

    'Set the render state to point blitting

    If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE

    'Set the last texture to a random number to force the engine to reload the texture
    'LastTexture = -65489
    'Check what type of rendering to do (blood or everything else)

    If Effect(EffectIndex).EffectNum = EffectNum_BloodSpray Then
        count = Effect(EffectIndex).ParticleCount \ 4
        D3DDevice.SetTexture 0, SangreTexture(1)
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex(0), LenB(Effect(EffectIndex).PartVertex(0))

        For I = 0 To count - 1

            With Effect(EffectIndex).Particles(I)

                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(I).y = Effect(EffectIndex).PartVertex(I).y + .sngZ
                Effect(EffectIndex).PartVertex(I).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            End With

        Next I

        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex(0), LenB(Effect(EffectIndex).PartVertex(0))
        D3DDevice.SetTexture 0, SangreTexture(4)
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex(count - 1), Len(Effect(EffectIndex).PartVertex(0))

        For I = count To count - 1 + count

            With Effect(EffectIndex).Particles(I)

                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(I).y = Effect(EffectIndex).PartVertex(I).y + .sngZ
                Effect(EffectIndex).PartVertex(I).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            End With

        Next I

        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex(count - 1), LenB(Effect(EffectIndex).PartVertex(0))
        D3DDevice.SetTexture 0, SangreTexture(2)
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex((count * 2) - 1), LenB(Effect(EffectIndex).PartVertex(0))

        For I = (count * 2) To (count * 2) - 1 + count

            With Effect(EffectIndex).Particles(I)

                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(I).y = Effect(EffectIndex).PartVertex(I).y + .sngZ
                Effect(EffectIndex).PartVertex(I).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            End With

        Next I

        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex((count * 2) - 1), LenB(Effect(EffectIndex).PartVertex(0))
        D3DDevice.SetTexture 0, SangreTexture(3)
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex((count * 3) - 1), LenB(Effect(EffectIndex).PartVertex(0))

        For I = (count * 3) To Effect(EffectIndex).ParticleCount

            With Effect(EffectIndex).Particles(I)

                If .sngZ < 1 Then Effect(EffectIndex).PartVertex(I).y = Effect(EffectIndex).PartVertex(I).y + .sngZ
                Effect(EffectIndex).PartVertex(I).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
            End With

        Next I

        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, count, Effect(EffectIndex).PartVertex((count * 3) - 1), LenB(Effect(EffectIndex).PartVertex(0))
    Else
        'Set the texture
        D3DDevice.SetTexture 0, ParticleTexture(Effect(EffectIndex).Gfx)
        'Draw all the particles at once
        D3DDevice.DrawPrimitiveUP D3DPT_POINTLIST, Effect(EffectIndex).ParticleCount, Effect(EffectIndex).PartVertex(0), LenB(Effect(EffectIndex).PartVertex(0))

        'Reset the render state back to normal

        If SetRenderStates Then D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
End Sub
 
Function Effect_Snow_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Begin]http://www.vbgore.com/CommonCode.Partic ... Snow_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Snow_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Snow      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Snow_Reset EffectIndex, LoopC, 1
    Next LoopC
 
    'Set the initial time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Snow_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Reset]http://www.vbgore.com/CommonCode.Partic ... Snow_Reset[/url]
'*****************************************************************
 
    If FirstReset = 1 Then
 
        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), Rnd * (ScreenHeight + 50), Rnd * 5, 5 + Rnd * 3, 0, 0
 
    Else
 
        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), -15 - Rnd * 185, Rnd * 5, 5 + Rnd * 3, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > ScreenWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > ScreenHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (ScreenWidth + 50)
 
    End If
 
    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.8, 0
 
End Sub
 
Private Sub Effect_Snow_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Snow_Update]http://www.vbgore.com/CommonCode.Partic ... now_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (ScreenWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (ScreenHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
 
            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Reset the particle
                Effect_Snow_Reset EffectIndex, LoopC
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Function Effect_Strengthen_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Begin]http://www.vbgore.com/CommonCode.Partic ... then_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Strengthen_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Strengthen    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Strengthen_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Strengthen_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Reset]http://www.vbgore.com/CommonCode.Partic ... then_Reset[/url]
'*****************************************************************
Dim a As Single
Dim x As Single
Dim y As Single
 
    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier)
    y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier)
 
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt x, y, 0, Rnd * -1, 0, -2
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_Strengthen_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Strengthen_Update]http://www.vbgore.com/CommonCode.Partic ... hen_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then
 
                    'Reset the particle
                    Effect_Strengthen_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Sub Effect_UpdateAll()
'*****************************************************************
'Updates all of the effects and renders them
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_UpdateAll]http://www.vbgore.com/CommonCode.Partic ... _UpdateAll[/url]
'*****************************************************************
Dim LoopC As Long
 
    'Make sure we have effects
    If NumEffects = 0 Then Exit Sub
 
    'Set the render state for the particle effects
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
 
    'Update every effect in use
    For LoopC = 1 To NumEffects
 
        'Make sure the effect is in use
        If Effect(LoopC).used Then
        
            'Update the effect position if the screen has moved
            Effect_UpdateOffset LoopC
        
            'Update the effect position if it is binded
            Effect_UpdateBinding LoopC
 
            'Find out which effect is selected, then update it
            If Effect(LoopC).EffectNum = EffectNum_Fire Then Effect_Fire_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Snow Then Effect_Snow_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Heal Then Effect_Heal_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Bless Then Effect_Bless_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Protection Then Effect_Protection_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Strengthen Then Effect_Strengthen_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Rain Then Effect_Rain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_EquationTemplate Then Effect_EquationTemplate_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Waterfall Then Effect_Waterfall_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Summon Then Effect_Summon_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Necro Then Effect_Necro_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Atom Then Effect_Atom_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_MeditMAX Then Effect_MeditMAX_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_PortalGroso Then Effect_PortalGroso_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_RedFountain Then Effect_RedFountain_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_Smoke Then Effect_Smoke_Update LoopC
            If Effect(LoopC).EffectNum = EffectNum_BloodSpray Then Effect_BloodSpray_Update LoopC
            
            'Render the effect
            If NieveOn = True And ZonaActual <> 1 Then
            Effect_Render LoopC, False
            Else
               If Effect(LoopC).EffectNum <> EffectNum_Snow And ZonaActual = 1 Then
            Effect_Render LoopC, False
               End If
           End If
        End If
 
    Next
    
    'Set the render state back for normal rendering
    D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
 
End Sub
 
Function Effect_Rain_Begin(ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Begin]http://www.vbgore.com/CommonCode.Partic ... Rain_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Rain_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Rain      'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(10)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Rain_Reset EffectIndex, LoopC, 1
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Rain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long, Optional ByVal FirstReset As Byte = 0)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Reset]http://www.vbgore.com/CommonCode.Partic ... Rain_Reset[/url]
'*****************************************************************
 
    If FirstReset = 1 Then
 
        'The very first reset
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * (ScreenWidth + 400)), Rnd * (ScreenHeight + 50), Rnd * 5, 25 + Rnd * 12, 0, 0
 
    Else
 
        'Any reset after first
        Effect(EffectIndex).Particles(Index).ResetIt -200 + (Rnd * 1200), -15 - Rnd * 185, Rnd * 5, 25 + Rnd * 12, 0, 0
        If Effect(EffectIndex).Particles(Index).sngX < -20 Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngX > ScreenWidth Then Effect(EffectIndex).Particles(Index).sngY = Rnd * (ScreenHeight + 50)
        If Effect(EffectIndex).Particles(Index).sngY > ScreenHeight Then Effect(EffectIndex).Particles(Index).sngX = Rnd * (ScreenWidth + 50)
 
    End If
 
    'Set the color
    Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 1, 0.4, 0
 
End Sub
 
Private Sub Effect_Rain_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Rain_Update]http://www.vbgore.com/CommonCode.Partic ... ain_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check if the particle is in use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if to reset the particle
            If Effect(EffectIndex).Particles(LoopC).sngX < -200 Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngX > (ScreenWidth + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
            If Effect(EffectIndex).Particles(LoopC).sngY > (ScreenHeight + 200) Then Effect(EffectIndex).Particles(LoopC).sngA = 0
 
            'Time for a reset, baby!
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Reset the particle
                Effect_Rain_Reset EffectIndex, LoopC
 
            Else
 
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub
 
Public Sub Effect_Begin(ByVal EffectIndex As Integer, ByVal x As Single, ByVal y As Single, ByVal GfxIndex As Byte, ByVal Particles As Byte, Optional ByVal Direction As Single = 180, Optional ByVal BindToMap As Boolean = False)
'*****************************************************************
'A very simplistic form of initialization for particle effects
'Should only be used for starting map-based effects
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Begin]http://www.vbgore.com/CommonCode.Particles.Effect_Begin[/url]
'*****************************************************************
'Actualizado: Lord Fers
 
Dim RetNum As Byte
 
    Select Case EffectIndex
        Case EffectNum_Fire
            RetNum = Effect_Fire_Begin(x, y, GfxIndex, Particles, Direction, 1)
        Case EffectNum_Bless
            RetNum = Effect_Bless_Begin(x, y, GfxIndex, Particles, 80, 1000)
        Case EffectNum_Waterfall
            RetNum = Effect_Waterfall_Begin(x, y, GfxIndex, 1000)
            
        Case EffectNum_Necro
              RetNum = Effect_Necro_Begin(x, y, GfxIndex, Particles, 6, 500)
              
              Case EffectNum_Atom
              RetNum = Effect_Atom_Begin(x, y, GfxIndex, Particles, 30, 1000)
              
               Case EffectNum_MeditMAX
              RetNum = Effect_MeditMAX_Begin(x, y, 1, 10, 100, 30, 10)
              
              Case EffectNum_PortalGroso
              RetNum = Effect_PortalGroso_Begin(x, y, 1, 1000, 10)
              
              
    End Select
    
    'Bind the effect to the map if needed
    If BindToMap Then Effect(RetNum).BoundToMap = 1
    
End Sub
 
Function Effect_Waterfall_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Begin]http://www.vbgore.com/CommonCode.Partic ... fall_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Waterfall_Begin = EffectIndex
 
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Waterfall     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Waterfall_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Waterfall_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Reset]http://www.vbgore.com/CommonCode.Partic ... fall_Reset[/url]
'*****************************************************************
 
    If Int(Rnd * 10) = 1 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + (Rnd * 60), Effect(EffectIndex).y + (Rnd * 130), 0, 8 + (Rnd * 6), 0, 0
    Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + (Rnd * 60), Effect(EffectIndex).y + (Rnd * 10), 0, 8 + (Rnd * 6), 0, 0
    End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.1, 0.1, 0.9, 0.6 + (Rnd * 0.4), 0
    
End Sub
 
Private Sub Effect_Waterfall_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Waterfall_Update]http://www.vbgore.com/CommonCode.Partic ... all_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
 
    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime
 
                'Check if the particle is ready to die
                If (.sngY > Effect(EffectIndex).y + 140) Or (.sngA = 0) Then
    
                    'Reset the particle
                    Effect_Waterfall_Reset EffectIndex, LoopC
    
                Else
 
                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).x = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).y = .sngY
    
                End If
    
            End If
            
        End With
 
    Next LoopC
 
End Sub
 
Function Effect_Summon_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 0) As Integer
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Begin]http://www.vbgore.com/CommonCode.Partic ... mmon_Begin[/url]
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long
 
    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function
 
    'Return the index of the used slot
    Effect_Summon_Begin = EffectIndex
 
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Summon    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True                     'Enable the effect
    Effect(EffectIndex).x = x                           'Set the effect's X coordinate
    Effect(EffectIndex).y = y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect
 
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
 
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(8)    'Size of the particles
 
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
 
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Summon_Reset EffectIndex, LoopC
    Next LoopC
 
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
End Function
 
Private Sub Effect_Summon_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Reset]http://www.vbgore.com/CommonCode.Partic ... mmon_Reset[/url]
'*****************************************************************
Dim x As Single
Dim y As Single
Dim R As Single
    
    If Effect(EffectIndex).Progression > 1000 Then
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 1.4
    Else
        Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.5
    End If
    R = (Index / 30) * exp(Index / Effect(EffectIndex).Progression)
    x = R * Cos(Index)
    y = R * Sin(Index)
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + x, Effect(EffectIndex).y + y, 0, 0, 0, 0
    Effect(EffectIndex).Particles(Index).ResetColor 0, Rnd, 0, 0.9, 0.2 + (Rnd * 0.2)
 
End Sub
 
Private Sub Effect_Summon_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: [url=http://www.vbgore.com/CommonCode.Particles.Effect_Summon_Update]http://www.vbgore.com/CommonCode.Partic ... mon_Update[/url]
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long
 
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
 
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
 
        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then
 
            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
 
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then
 
                'Check if the effect is ending
                If Effect(EffectIndex).Progression < 1800 Then
 
                    'Reset the particle
                    Effect_Summon_Reset EffectIndex, LoopC
 
                Else
 
                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False
 
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
 
                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False
 
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0
 
                End If
 
            Else
            
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY
 
            End If
 
        End If
 
    Next LoopC
 
End Sub


Function Effect_Necro_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Direction As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Necro_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Necro     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles        'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).x = x          'Set the effect's X coordinate
    Effect(EffectIndex).y = y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Direction = Direction       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)
    Effect(EffectIndex).TargetAA = 0
    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        'Effect_Necro_Reset EffectIndex, LoopC
    Next LoopC
    
    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Necro_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    'Static TargetA As Single
    Dim Co As Single
    Dim sI As Single
    'Calculate the angle
    
    If Effect(EffectIndex).TargetAA = 0 And Effect(EffectIndex).GoToX <> -30000 Then Effect(EffectIndex).TargetAA = Engine_GetAngle(Effect(EffectIndex).x, Effect(EffectIndex).y, Effect(EffectIndex).GoToX, Effect(EffectIndex).GoToY) + 180
    
    sI = Sin(Effect(EffectIndex).TargetAA * DegreeToRadian)
    Co = Cos(Effect(EffectIndex).TargetAA * DegreeToRadian)
    
    'Reset the particle
    If RandomNumber(1, 2) = 2 Then
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * 20, Si * Sin(Effect(EffectIndex).Progression * 3) * 20, 0, 0
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + Co * Sin(Effect(EffectIndex).Progression) * 25, Effect(EffectIndex).y + sI * Sin(Effect(EffectIndex).Progression) * 25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2 + (Rnd * 0.5), 1, 0.5 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x, Effect(EffectIndex).y, Co * Sin(Effect(EffectIndex).Progression * 3) * -20, Si * Sin(Effect(EffectIndex).Progression * 3) * -20, 0, 0
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + Co * Sin(Effect(EffectIndex).Progression) * -25, Effect(EffectIndex).y + sI * Sin(Effect(EffectIndex).Progression) * -25, 0, 0, 0, 0
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2 + (Rnd * 0.5), 0.2, 0.7 + (Rnd * 0.2), 0.1 + (Rnd * 4.09)
    End If
    

End Sub

Private Sub Effect_Necro_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.5 And RandomNumber(1, 3) = 3 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Necro_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub


Function Effect_Atom_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Atom_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_Atom    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles           'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(16)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Atom_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Atom_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

Dim a As Single
Dim x As Single
Dim y As Single
Dim R As Single
    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    R = Rnd * 4
    If R < 1 Then
        x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier) / 3 + (Cos(a) * Effect(EffectIndex).Modifier)
        y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(Index).ResetColor 0.2, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 2 Then
        x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier)
        y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier) / 3 + (Sin(a) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(Index).ResetColor 1, 1, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 3 Then
        x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier) / 3
        y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier)
        Effect(EffectIndex).Particles(Index).ResetColor 1, 0.2, 0.2, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    ElseIf R < 4 Then
        x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier)
        y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier) / 3
        
        Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 1, 0.6 + (Rnd * 0.4), 0.06 + (Rnd * 0.2)
    End If
    
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt x, y, 0, 0, 0, -1
    

End Sub

Private Sub Effect_Atom_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate the time difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check if particle is in use
        If Effect(EffectIndex).Particles(LoopC).used Then

            'Update the particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_Atom_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_MeditMAX_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Size As Byte = 30, Optional ByVal Time As Single = 10, Optional R As Single) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_MeditMAX_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_MeditMAX    'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles  'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic
    Effect(EffectIndex).Modifier = Size         'How large the circle is
    Effect(EffectIndex).Progression = Time      'How long the effect will last
    
    Effect(EffectIndex).R = R

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_MeditMAX_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_MeditMAX_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim a As Single
Dim x As Single
Dim y As Single
Dim AccX As Single

    'Get the positions
    a = Rnd * 360 * DegreeToRadian
    
    Do While a > 3 And a < 3.27
        Randomize 1000
        z = z + 1
        If z > 6 Then z = 1
        a = Rnd * 60 * DegreeToRadian * 6
    Loop
    
    x = Effect(EffectIndex).x - (Sin(a) * Effect(EffectIndex).Modifier)
    y = Effect(EffectIndex).y + (Cos(a) * Effect(EffectIndex).Modifier) / 2

    
    Effect(EffectIndex).Particles(Index).ResetIt x, y, -(Sgn(Sin(a)) * 4 * (Rnd - 0.2)), Rnd * -1, 0, -1.8
    Effect(EffectIndex).Particles(Index).ResetColor 0.9, 0.9, 0.7 * Rnd, 0.7 + (Rnd * 0.3), 0.05 + (Rnd * 0.1)

End Sub
Public Sub Effect_MeditMAX_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    'If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_MeditMAX_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_PortalGroso_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal Progression As Single = 1) As Integer
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_PortalGroso_Begin = EffectIndex
    
    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_PortalGroso  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True                     'Enable the effect
    Effect(EffectIndex).x = x                           'Set the effect's X coordinate
    Effect(EffectIndex).y = y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx                       'Set the graphic
    Effect(EffectIndex).Progression = Progression       'If we loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_PortalGroso_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_PortalGroso_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
Dim x As Single
Dim y As Single
Dim R As Single
Dim ind As Integer
    Effect(EffectIndex).Progression = Effect(EffectIndex).Progression + 0.1
    ind = CInt(Index / 10) * 10
    R = ((Index + 100) / 4) * exp((Index + 100) / 2000)
    x = R * Cos(Index) * 0.25 '* 0.3 * 0.25
    y = R * Sin(Index) * 0.25 '* 0.2 * 0.25
    'Reset the particle
    'If Rnd * 20 < 1 Then
    '    Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x + x, Effect(EffectIndex).y + y, 0, 0, 0, -1.5 * (ind / Effect(EffectIndex).ParticleCount)
    'Else
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + x, Effect(EffectIndex).y + y, 0, 0, 0, 0
    'End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0, (0.7 * ind / Effect(EffectIndex).ParticleCount), 1, IIf(ind / Effect(EffectIndex).ParticleCount / 7 < 0.03, 0.03, ind / Effect(EffectIndex).ParticleCount / 7)

End Sub

Private Sub Effect_PortalGroso_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long
Dim Owner As Integer
    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'For LoopC = 1 To LastChar
    '    If EffectIndex = CharList(LoopC).AuraIndex Then
    '        Owner = LoopC
    '    End If
    'Next
    
    'If ClientSetup.bGraphics < 2 Then Effect(EffectIndex).Used = False
    
    'If Owner = 0 Then
    '    Effect(EffectIndex).Used = False
    'End If
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then

            ''Update The Particle
            'If EffectIndex <> CharList(UserCharIndex).AuraIndex Then
                Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime
            'Else
             '   Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime, True
            'End If
            
            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0 Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Then

                    'Reset the particle
                    Effect_PortalGroso_Reset EffectIndex, LoopC

                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else

                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub

Function Effect_RedFountain_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer) As Integer

Dim EffectIndex As Integer
Dim LoopC As Long

'Get the next open effect slot

    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_RedFountain_Begin = EffectIndex

    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_RedFountain     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles          'Set the number of particles
    Effect(EffectIndex).used = True             'Enabled the effect
    Effect(EffectIndex).x = x                   'Set the effect's X coordinate
    Effect(EffectIndex).y = y                   'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx               'Set the graphic

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(20)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_RedFountain_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_RedFountain_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

    'If Int(Rnd * 10) < 6 Then
        Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + (Rnd * 10) - 5, Effect(EffectIndex).y - (Rnd * 10), 0, 1, 0, -1 - Rnd * 0.25
    'Else
        'Effect(EffectIndex).Particles(index).ResetIt Effect(EffectIndex).x + (Rnd * 10) - 5, Effect(EffectIndex).y - (Rnd * 10), 1 + (Rnd * 5), -15 - (Rnd * 3), 0, 1.1 + Rnd * 0.1
    'End If
    Effect(EffectIndex).Particles(Index).ResetColor 0.9, Rnd * 0.7, 0.1, 0.6 + (Rnd * 0.4), 0.035 + Rnd * 0.01
    
End Sub

Private Sub Effect_RedFountain_Update(ByVal EffectIndex As Integer)
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime

    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime

    'Go through the particle loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
    
        With Effect(EffectIndex).Particles(LoopC)
    
            'Check if the particle is in use
            If .used Then
    
                'Update The Particle
                .UpdateParticle ElapsedTime

                'Check if the particle is ready to die
                If (.sngA < 0) Or (.sngY > Effect(EffectIndex).y + 100) Then
    
                    'Reset the particle
                    Effect_RedFountain_Reset EffectIndex, LoopC
    
                Else

                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(.sngR, .sngG, .sngB, .sngA)
                    Effect(EffectIndex).PartVertex(LoopC).x = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).y = .sngY
    
                End If
    
            End If
            
        End With

    Next LoopC

End Sub


Function Effect_Smoke_Begin(ByVal x As Single, ByVal y As Single, ByVal Gfx As Integer, ByVal Particles As Integer, Optional ByVal radius As Integer = 180, Optional ByVal Progression As Single = 1) As Integer
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Begin
'*****************************************************************
Dim EffectIndex As Integer
Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot
    If EffectIndex = -1 Then Exit Function

    'Return the index of the used slot
    Effect_Smoke_Begin = EffectIndex

    'Set The Effect's Variables
    Effect(EffectIndex).EffectNum = EffectNum_Smoke     'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles - Particles * 0.25 * (2)        'Set the number of particles
    Effect(EffectIndex).used = True     'Enabled the effect
    Effect(EffectIndex).x = x          'Set the effect's X coordinate
    Effect(EffectIndex).y = y           'Set the effect's Y coordinate
    Effect(EffectIndex).Gfx = Gfx       'Set the graphic
    Effect(EffectIndex).Modifier = radius       'The direction the effect is animat
    Effect(EffectIndex).Progression = Progression   'Loop the effect

    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount

    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(15)    'Size of the particles

    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles
    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_Smoke_Reset EffectIndex, LoopC
    Next LoopC

    'Set The Initial Time
    Effect(EffectIndex).PreviousFrame = timeGetTime

End Function

Private Sub Effect_Smoke_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Reset
'*****************************************************************
    Dim v As Single
    v = Rnd * 20
    'Reset the particle
    Effect(EffectIndex).Particles(Index).ResetIt Effect(EffectIndex).x + Effect(EffectIndex).Modifier * RandomNumber(-1, 1) * Rnd / 2, Effect(EffectIndex).y - Effect(EffectIndex).Modifier, 0, 0, 0, Rnd * -1.5
    Effect(EffectIndex).Particles(Index).ResetColor 0.2, 0.2, 0.2, 1, 0

End Sub

Private Sub Effect_Smoke_Update(ByVal EffectIndex As Integer)
'*****************************************************************
'More info: http://www.vbgore.com/CommonCode.Particles.Effect_Fire_Update
'*****************************************************************
Dim ElapsedTime As Single
Dim LoopC As Long

    'Calculate The Time Difference
    ElapsedTime = (timeGetTime - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = timeGetTime
    
    'Update the life span
    If Effect(EffectIndex).Progression > 0 Then Effect(EffectIndex).Progression = Effect(EffectIndex).Progression - ElapsedTime
    
    'Go Through The Particle Loop
    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        'Check If Particle Is In Use
        If Effect(EffectIndex).Particles(LoopC).used Then

            'Update The Particle
            Effect(EffectIndex).Particles(LoopC).UpdateParticle ElapsedTime

            'Check if the particle is ready to die
            If Effect(EffectIndex).Particles(LoopC).sngA <= 0.3 Or Effect(EffectIndex).Particles(LoopC).sngY + Effect(EffectIndex).Modifier * 3 < Effect(EffectIndex).y Then

                'Check if the effect is ending
                If Effect(EffectIndex).Progression > 0 Or Effect(EffectIndex).Progression = -5000 Then
                    
                    
                    'Reset the particle
                    Effect_Smoke_Reset EffectIndex, LoopC
                    
                Else

                    'Disable the particle
                    Effect(EffectIndex).Particles(LoopC).used = False

                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1

                    'Check if the effect is out of particles
                    If Effect(EffectIndex).ParticlesLeft = 0 Then Effect(EffectIndex).used = False

                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).PartVertex(LoopC).color = 0

                End If

            Else
                
                'Set the particle information on the particle vertex
                Effect(EffectIndex).PartVertex(LoopC).color = D3DColorMake(Effect(EffectIndex).Particles(LoopC).sngR, Effect(EffectIndex).Particles(LoopC).sngG, Effect(EffectIndex).Particles(LoopC).sngB, Effect(EffectIndex).Particles(LoopC).sngA)
                Effect(EffectIndex).PartVertex(LoopC).x = Effect(EffectIndex).Particles(LoopC).sngX
                Effect(EffectIndex).PartVertex(LoopC).y = Effect(EffectIndex).Particles(LoopC).sngY

            End If

        End If

    Next LoopC

End Sub


Function Effect_BloodSpray_Begin(ByVal x As Single, _
                                 ByVal y As Single, _
                                 ByVal Particles As Integer, _
                                 ByVal Direction As Single, _
                                 Optional ByVal Intensity As Single = 1) As Integer

    Dim EffectIndex As Integer
    Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function
    'Return the index of the used slot
    Effect_BloodSpray_Begin = EffectIndex
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_BloodSpray  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True                     'Enable the effect
    Effect(EffectIndex).x = x                           'Set the effect's X coordinate
    Effect(EffectIndex).y = y                           'Set the effect's Y coordinate
    Effect(EffectIndex).Direction = Direction           'Direction
    Effect(EffectIndex).Modifier = Intensity
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles

    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        Effect_BloodSpray_Reset EffectIndex, LoopC

    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = GetTickCount
End Function

Private Sub Effect_BloodSpray_Reset(ByVal EffectIndex As Integer, ByVal Index As Long)

'Reset the particle

    With Effect(EffectIndex)
        .Particles(Index).ResetIt .x + (Rnd * 16) - 8, .y + (Rnd * 32) - 16, Sin((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), -Cos((.Direction - 10 + (Rnd * 20)) * DegreeToRadian) * (30 * .Modifier * Rnd), 0, 0, -10, -2 - (Rnd * 30), 8 + Rnd * 4
        .Particles(Index).ResetColor 1, 1, 1, 0.8, 0
    End With

End Sub


Private Sub Effect_BloodSpray_Update(ByVal EffectIndex As Integer)

    Dim ElapsedTime As Single
    Dim LoopC As Long
    Dim TileX As Long
    Dim TileY As Long

    'Calculate the time difference
    ElapsedTime = (GetTickCount - Effect(EffectIndex).PreviousFrame) * 0.01
    Effect(EffectIndex).PreviousFrame = GetTickCount

    'Go through the particle loop

    For LoopC = 0 To Effect(EffectIndex).ParticleCount

        With Effect(EffectIndex).Particles(LoopC)

            'Check if particle is in Use

            If .used Then
                'Update the particle
                .UpdateParticle ElapsedTime
                'Don't pass any walls/etc
                TileX = Engine_SPtoTPX(.sngX)
                TileY = Engine_SPtoTPY(.sngY)

                If TileX < 1 Then
                    .sngZ = 1.1
                ElseIf TileY < 1 Then
                    .sngZ = 1.1
                ElseIf TileX > 92 Then
                    .sngZ = 1.1
                ElseIf TileY > 92 Then
                    .sngZ = 1.1
                End If

                If .sngZ <> 1.1 Then
                    'If MapData(TileX, TileY).BlockedAttack Then
                    '.sngZ = 1.1
                    'End If
                End If

                'Blood trails

                If LoopC = 0 Or LoopC Mod 15 = 0 Then
                    If Int(Rnd * 3) = 0 Then
                        If Int(Rnd * 2) = 0 Then
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 2
                        Else
                            Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 1
                        End If
                    End If
                End If

                'Check if to kill off the particle

                If .sngZ > 1 Then
                    'Disable the particle
                    .used = False
                    'Subtract from the total particle count
                    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticlesLeft - 1
                    'Clear the color (dont leave behind any artifacts)
                    Effect(EffectIndex).Particles(LoopC).sngA = 0

                    'Check if we lost all the particles

                    If Effect(EffectIndex).ParticlesLeft <= 0 Then Effect(EffectIndex).used = False
                    'Create the blood splatter
                    Engine_Blood_Create .sngX + ParticleOffsetX, .sngY + ParticleOffsetY, 0
                Else
                    'Set the particle information on the particle vertex
                    Effect(EffectIndex).PartVertex(LoopC).color = 1258291200
                    Effect(EffectIndex).PartVertex(LoopC).x = .sngX
                    Effect(EffectIndex).PartVertex(LoopC).y = .sngY
                End If
            End If

        End With

    Next LoopC

End Sub

Function Effect_BloodSplatter_Begin(ByVal x As Single, _
                                    ByVal y As Single, _
                                    ByVal Particles As Integer) As Integer

    Dim EffectIndex As Integer
    Dim LoopC As Long

    'Get the next open effect slot
    EffectIndex = Effect_NextOpenSlot

    If EffectIndex = -1 Then Exit Function
    'Return the index of the used slot
    Effect_BloodSplatter_Begin = EffectIndex
    'Set the effect's variables
    Effect(EffectIndex).EffectNum = EffectNum_BloodSplatter  'Set the effect number
    Effect(EffectIndex).ParticleCount = Particles       'Set the number of particles
    Effect(EffectIndex).used = True                     'Enable the effect
    Effect(EffectIndex).x = x                           'Set the effect's X coordinate
    Effect(EffectIndex).y = y                           'Set the effect's Y coordinate
    'Set the number of particles left to the total avaliable
    Effect(EffectIndex).ParticlesLeft = Effect(EffectIndex).ParticleCount
    'Set the float variables
    Effect(EffectIndex).FloatSize = Effect_FToDW(7)    'Size of the particles
    'Redim the number of particles
    ReDim Effect(EffectIndex).Particles(0 To Effect(EffectIndex).ParticleCount)
    ReDim Effect(EffectIndex).PartVertex(0 To Effect(EffectIndex).ParticleCount)

    'Create the particles

    For LoopC = 0 To Effect(EffectIndex).ParticleCount
        Set Effect(EffectIndex).Particles(LoopC) = New Particle
        Effect(EffectIndex).Particles(LoopC).used = True
        Effect(EffectIndex).PartVertex(LoopC).rhw = 1
        'Effect_BloodSplatter_Reset EffectIndex, loopc

    Next LoopC

    'Set the initial time
    Effect(EffectIndex).PreviousFrame = GetTickCount
End Function

Public Function Engine_GetAngle(ByVal CenterX As Integer, _
                                ByVal CenterY As Integer, _
                                ByVal TargetX As Integer, _
                                ByVal TargetY As Integer) As Single

'************************************************************
'Gets the angle between two points in a 2d plane
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetAngle
'************************************************************

    Dim SideA As Single
    Dim SideC As Single

    On Error GoTo ErrOut

    'Check for horizontal lines (90 or 270 degrees)

    If CenterY = TargetY Then

        'Check for going right (90 degrees)

        If CenterX < TargetX Then
            Engine_GetAngle = 90
            'Check for going left (270 degrees)
        Else
            Engine_GetAngle = 270
        End If

        'Exit the function

        Exit Function

    End If

    'Check for horizontal lines (360 or 180 degrees)

    If CenterX = TargetX Then

        'Check for going up (360 degrees)

        If CenterY > TargetY Then
            Engine_GetAngle = 360
            'Check for going down (180 degrees)
        Else
            Engine_GetAngle = 180
        End If

        'Exit the function

        Exit Function

    End If

    'Calculate Side C
    SideC = Sqr(Abs(TargetX - CenterX) ^ 2 + Abs(TargetY - CenterY) ^ 2)

    'Side B = CenterY

    'Calculate Side A
    SideA = Sqr(Abs(TargetX - CenterX) ^ 2 + TargetY ^ 2)

    'Calculate the angle
    Engine_GetAngle = (SideA ^ 2 - CenterY ^ 2 - SideC ^ 2) / (CenterY * SideC * -2)

    Engine_GetAngle = (Atn(-Engine_GetAngle / Sqr(-Engine_GetAngle * Engine_GetAngle + 1)) + 1.5708) * 57.29583

    'If the angle is >180, subtract from 360

    If TargetX < CenterX Then Engine_GetAngle = 360 - Engine_GetAngle

    Exit Function

    'Check for error
ErrOut:
    'Return a 0 saying there was an error
    Engine_GetAngle = 0

    Exit Function

End Function

Sub Engine_Blood_Create(ByVal x As Single, ByVal y As Single, ByVal Size As Byte)

'*****************************************************************
'Creates a puddle of blood on the ground
'*****************************************************************

    Dim TileX As Integer
    Dim TileY As Integer

    Const texwidth As Single = 64
    Const texheight As Single = 64
    Const NumLarge As Long = 2
    Const NumMedium As Long = 12
    Const NumSmall As Long = 8
    Const Large1X As Single = 0
    Const Large1Y As Single = 0
    Const Large1W As Single = 32
    Const Large1H As Single = 16
    Const Large2X As Single = 0
    Const Large2Y As Single = 17
    Const Large2W As Single = 32
    Const Large2H As Single = 16
    Const Med1X As Single = 0
    Const Med1Y As Single = 34
    Const Med1W As Single = 14
    Const Med1H As Single = 6
    Const Med2X As Single = 0
    Const Med2Y As Single = 41
    Const Med2W As Single = 12
    Const Med2H As Single = 7
    Const Med3X As Single = 0
    Const Med3Y As Single = 49
    Const Med3W As Single = 11
    Const Med3H As Single = 8
    Const Med4X As Single = 15
    Const Med4Y As Single = 34
    Const Med4W As Single = 12
    Const Med4H As Single = 5
    Const Med5X As Single = 15
    Const Med5Y As Single = 40
    Const Med5W As Single = 8
    Const Med5H As Single = 9
    Const Med6X As Single = 12
    Const Med6Y As Single = 50
    Const Med6W As Single = 9
    Const Med6H As Single = 7
    Const Med7X As Single = 22
    Const Med7Y As Single = 50
    Const Med7W As Single = 10
    Const Med7H As Single = 7
    Const Med8X As Single = 33
    Const Med8Y As Single = 0
    Const Med8W As Single = 16
    Const Med8H As Single = 7
    Const Med9X As Single = 33
    Const Med9Y As Single = 8
    Const Med9W As Single = 14
    Const Med9H As Single = 7
    Const Med10X As Single = 33
    Const Med10Y As Single = 29
    Const Med10W As Single = 17
    Const Med10H As Single = 8
    Const Med11X As Single = 33
    Const Med11Y As Single = 38
    Const Med11W As Single = 15
    Const Med11H As Single = 9
    Const Med12X As Single = 33
    Const Med12Y As Single = 48
    Const Med12W As Single = 11
    Const Med12H As Single = 6
    Const Small1X As Single = 28
    Const Small1Y As Single = 34
    Const Small1W As Single = 4
    Const Small1H As Single = 6
    Const Small2X As Single = 24
    Const Small2Y As Single = 41
    Const Small2W As Single = 6
    Const Small2H As Single = 4
    Const Small3X As Single = 33
    Const Small3Y As Single = 16
    Const Small3W As Single = 10
    Const Small3H As Single = 3
    Const Small4X As Single = 44
    Const Small4Y As Single = 16
    Const Small4W As Single = 4
    Const Small4H As Single = 5
    Const Small5X As Single = 33
    Const Small5Y As Single = 20
    Const Small5W As Single = 8
    Const Small5H As Single = 4
    Const Small6X As Single = 42
    Const Small6Y As Single = 22
    Const Small6W As Single = 4
    Const Small6H As Single = 3
    Const Small7X As Single = 33
    Const Small7Y As Single = 25
    Const Small7W As Single = 8
    Const Small7H As Single = 3
    Const Small8X As Single = 42
    Const Small8Y As Single = 26
    Const Small8W As Single = 5
    Const Small8H As Single = 2

    Dim BloodIndex As Integer
    Dim I As Long
    Dim l As Long

    'Find the tile
    TileX = ((x - 288) \ 32) + 1
    TileY = ((y - 288) \ 32) + 1

    If TileX < 1 Then TileX = 1
    If TileX > 100 Then TileX = 100
    If TileY < 1 Then TileY = 1
    If TileY > 100 Then TileY = 100

    'Check if there is too much blood on this tile already

    If MapData(TileX, TileY).Blood > 40 Then Exit Sub

    'Get the next open blood slot

    Do
        BloodIndex = BloodIndex + 1

        'Update LastBlood if we go over the size of the current array

        If BloodIndex > LastBlood Then
            LastBlood = BloodIndex
            ReDim Preserve BloodList(1 To LastBlood)

            Exit Do

        End If

    Loop While BloodList(BloodIndex).Life > 0

    'Set the blood's lfie
    BloodList(BloodIndex).Life = GetTickCount + 1800

    'Get a random size if none is specified

    If Size < 1 Or Size > 3 Then
        Size = Int(Rnd * (NumLarge + NumSmall + NumMedium)) + 1

        If Size <= NumLarge Then
            Size = 3
        ElseIf Size <= NumLarge + NumMedium Then
            Size = 2
        Else
            Size = 1
        End If
    End If

    With BloodList(BloodIndex)

        'Set up the general blood information

        For l = 0 To 5
            .v(l).color = -1
            .v(l).rhw = 1
            .v(l).x = x
            .v(l).y = y

        Next l

        ' 3____4
        ' 0|\\ | 0 = 3
        ' | \\ | 1 = 5
        ' | \\ |
        ' | \\ |
        ' 2|____\\|
        ' 1 5
        'Large blood

        If Size = 3 Then
            I = Int(Rnd * NumLarge) + 1

            Select Case I

                Case 1
                    .v(4).x = x + Large1W
                    .v(2).y = y + Large1H
                    .v(0).tu = Large1X / texwidth
                    .v(0).tv = Large1Y / texheight
                    .v(5).tu = (Large1X + Large1W) / texwidth
                    .v(5).tv = (Large1Y + Large1H) / texheight

                Case 2
                    .v(4).x = x + Large2W
                    .v(2).y = y + Large2H
                    .v(0).tu = Large2X / texwidth
                    .v(0).tv = Large2Y / texheight
                    .v(5).tu = (Large2X + Large2W) / texwidth
                    .v(5).tv = (Large2Y + Large2H) / texheight
            End Select

            'Medium blood
        ElseIf Size = 2 Then
            I = Int(Rnd * NumMedium) + 1

            Select Case I

                Case 1
                    .v(4).x = x + Med1W
                    .v(2).y = y + Med1H
                    .v(0).tu = Med1X / texwidth
                    .v(0).tv = Med1Y / texheight
                    .v(5).tu = (Med1X + Med1W) / texwidth
                    .v(5).tv = (Med1Y + Med1H) / texheight

                Case 2
                    .v(4).x = x + Med2W
                    .v(2).y = y + Med2H
                    .v(0).tu = Med2X / texwidth
                    .v(0).tv = Med2Y / texheight
                    .v(5).tu = (Med2X + Med2W) / texwidth
                    .v(5).tv = (Med2Y + Med2H) / texheight

                Case 3
                    .v(4).x = x + Med3W
                    .v(2).y = y + Med3H
                    .v(0).tu = Med3X / texwidth
                    .v(0).tv = Med3Y / texheight
                    .v(5).tu = (Med3X + Med3W) / texwidth
                    .v(5).tv = (Med3Y + Med3H) / texheight

                Case 4
                    .v(4).x = x + Med4W
                    .v(2).y = y + Med4H
                    .v(0).tu = Med4X / texwidth
                    .v(0).tv = Med4Y / texheight
                    .v(5).tu = (Med4X + Med4W) / texwidth
                    .v(5).tv = (Med4Y + Med4H) / texheight

                Case 5
                    .v(4).x = x + Med5W
                    .v(2).y = y + Med5H
                    .v(0).tu = Med5X / texwidth
                    .v(0).tv = Med5Y / texheight
                    .v(5).tu = (Med5X + Med5W) / texwidth
                    .v(5).tv = (Med5Y + Med5H) / texheight

                Case 6
                    .v(4).x = x + Med6W
                    .v(2).y = y + Med6H
                    .v(0).tu = Med6X / texwidth
                    .v(0).tv = Med6Y / texheight
                    .v(5).tu = (Med6X + Med6W) / texwidth
                    .v(5).tv = (Med6Y + Med6H) / texheight

                Case 7
                    .v(4).x = x + Med7W
                    .v(2).y = y + Med7H
                    .v(0).tu = Med7X / texwidth
                    .v(0).tv = Med7Y / texheight
                    .v(5).tu = (Med7X + Med7W) / texwidth
                    .v(5).tv = (Med7Y + Med7H) / texheight

                Case 8
                    .v(4).x = x + Med8W
                    .v(2).y = y + Med8H
                    .v(0).tu = Med8X / texwidth
                    .v(0).tv = Med8Y / texheight
                    .v(5).tu = (Med8X + Med8W) / texwidth
                    .v(5).tv = (Med8Y + Med8H) / texheight

                Case 9
                    .v(4).x = x + Med9W
                    .v(2).y = y + Med9H
                    .v(0).tu = Med9X / texwidth
                    .v(0).tv = Med9Y / texheight
                    .v(5).tu = (Med9X + Med9W) / texwidth
                    .v(5).tv = (Med9Y + Med9H) / texheight

                Case 10
                    .v(4).x = x + Med10W
                    .v(2).y = y + Med10H
                    .v(0).tu = Med10X / texwidth
                    .v(0).tv = Med10Y / texheight
                    .v(5).tu = (Med10X + Med10W) / texwidth
                    .v(5).tv = (Med10Y + Med10H) / texheight

                Case 11
                    .v(4).x = x + Med11W
                    .v(2).y = y + Med11H
                    .v(0).tu = Med11X / texwidth
                    .v(0).tv = Med11Y / texheight
                    .v(5).tu = (Med11X + Med11W) / texwidth
                    .v(5).tv = (Med11Y + Med11H) / texheight

                Case 12
                    .v(4).x = x + Med12W
                    .v(2).y = y + Med12H
                    .v(0).tu = Med12X / texwidth
                    .v(0).tv = Med12Y / texheight
                    .v(5).tu = (Med12X + Med12W) / texwidth
                    .v(5).tv = (Med12Y + Med12H) / texheight
            End Select

            'Small blood
        Else
            I = Int(Rnd * NumSmall) + 1

            Select Case I

                Case 1
                    .v(4).x = x + Small1W
                    .v(2).y = y + Small1H
                    .v(0).tu = Small1X / texwidth
                    .v(0).tv = Small1Y / texheight
                    .v(5).tu = (Small1X + Small1W) / texwidth
                    .v(5).tv = (Small1Y + Small1H) / texheight

                Case 2
                    .v(4).x = x + Small2W
                    .v(2).y = y + Small2H
                    .v(0).tu = Small2X / texwidth
                    .v(0).tv = Small2Y / texheight
                    .v(5).tu = (Small2X + Small2W) / texwidth
                    .v(5).tv = (Small2Y + Small2H) / texheight

                Case 3
                    .v(4).x = x + Small3W
                    .v(2).y = y + Small3H
                    .v(0).tu = Small3X / texwidth
                    .v(0).tv = Small3Y / texheight
                    .v(5).tu = (Small3X + Small3W) / texwidth
                    .v(5).tv = (Small3Y + Small3H) / texheight

                Case 4
                    .v(4).x = x + Small4W
                    .v(2).y = y + Small4H
                    .v(0).tu = Small4X / texwidth
                    .v(0).tv = Small4Y / texheight
                    .v(5).tu = (Small4X + Small4W) / texwidth
                    .v(5).tv = (Small4Y + Small4H) / texheight

                Case 5
                    .v(4).x = x + Small5W
                    .v(2).y = y + Small5H
                    .v(0).tu = Small5X / texwidth
                    .v(0).tv = Small5Y / texheight
                    .v(5).tu = (Small5X + Small5W) / texwidth
                    .v(5).tv = (Small5Y + Small5H) / texheight

                Case 6
                    .v(4).x = x + Small6W
                    .v(2).y = y + Small6H
                    .v(0).tu = Small6X / texwidth
                    .v(0).tv = Small6Y / texheight
                    .v(5).tu = (Small6X + Small6W) / texwidth
                    .v(5).tv = (Small6Y + Small6H) / texheight

                Case 7
                    .v(4).x = x + Small7W
                    .v(2).y = y + Small7H
                    .v(0).tu = Small7X / texwidth
                    .v(0).tv = Small7Y / texheight
                    .v(5).tu = (Small7X + Small7W) / texwidth
                    .v(5).tv = (Small7Y + Small7H) / texheight

                Case 8
                    .v(4).x = x + Small8W
                    .v(2).y = y + Small8H
                    .v(0).tu = Small8X / texwidth
                    .v(0).tv = Small8Y / texheight
                    .v(5).tu = (Small8X + Small8W) / texwidth
                    .v(5).tv = (Small8Y + Small8H) / texheight
            End Select

        End If

        'These variables are the same no blood used
        .v(4).tu = .v(5).tu
        .v(4).tv = .v(0).tv
        .v(2).tu = .v(0).tu
        .v(2).tv = .v(5).tv
        .v(5).x = .v(4).x
        .v(5).y = .v(2).y
        .v(3) = .v(0)
        .v(1) = .v(5)
        'Find the blood tile location
        .TileX = TileX
        .TileY = TileY
        MapData(.TileX, .TileY).Blood = MapData(.TileX, .TileY).Blood + 1
    End With

End Sub

Sub Engine_Blood_Erase(ByVal BloodIndex As Long)

'*****************************************************************
'Erases a blood splatter by index
'*****************************************************************

    With BloodList(BloodIndex)
        'Set the life to 0 to not use it
        BloodList(BloodIndex).Life = 0

        'Erase the blood from the tile

        If .TileX > 0 Then
            If .TileY > 0 Then
                If .TileX <= 100 Then
                    If .TileY <= 100 Then
                        MapData(.TileX, .TileY).Blood = MapData(.TileX, .TileY).Blood - 1
                    End If
                End If
            End If
        End If

    End With

    'Resize the array if needed

    If BloodIndex = LastBlood Then

        Do Until BloodList(LastBlood).Life > 0
            LastBlood = LastBlood - 1

            If LastBlood = 0 Then Exit Do
        Loop

        If LastBlood <> BloodIndex Then
            If LastBlood <> 0 Then
                ReDim Preserve BloodList(1 To LastBlood)
            Else
                Erase BloodList
            End If
        End If
    End If

End Sub

Public Sub Engine_Render_Blood()

'*****************************************************************
'Batch render the blood on the ground
'*****************************************************************

    Dim BloodVB As Direct3DVertexBuffer8    'Vertex buffer
    Dim BloodVL() As TLVERTEX2   'Vertex list
    Dim BloodCount As Long
    Dim Alpha As Long
    Dim I As Long
    Dim J As Long
    Dim Tex As Direct3DTexture8
    Dim srdesc As D3DSURFACE_DESC
    Dim ColGraR As Integer
    Dim ColGraG As Integer
    Dim ColGraB As Integer
    'Check if Render Blood option is enabled

    'If Config.renderBlood = 0 Then Exit Sub

    'Check for any blood
    Select Case Colorsangre
    Case 2
        ColGraR = 0
        ColGraG = 255
        ColGraB = 0
    Case 3
        ColGraR = 0
        ColGraG = 0
        ColGraB = 255
    Case 4
        ColGraR = 0
        ColGraG = 0
        ColGraB = 0
    Case Else
        ColGraR = 255
        ColGraG = 0
        ColGraB = 0

    End Select

    If LastBlood = 0 Then Exit Sub

    Set Tex = D3DX.CreateTextureFromFileEx(D3DDevice, App.path & "\RECURSOS\25005.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, &HFF000000, ByVal 0, ByVal 0)


    Tex.GetLevelDesc 0, srdesc

    'Create the vertex list
    ReDim BloodVL(1 To LastBlood * 6) As TLVERTEX2

    For I = 1 To LastBlood

        If BloodList(I).Life <> 0 Then
            If BloodList(I).Life > GetTickCount Then
                If BloodList(I).Life - GetTickCount > 3000 Then
                    Alpha = 255
                Else
                    Alpha = (BloodList(I).Life - GetTickCount) / 7

                    If Alpha > 255 Then Alpha = 255
                End If

                For J = 1 To 6
                    BloodVL((BloodCount * 6) + J) = BloodList(I).v(J - 1)

                    With BloodVL((BloodCount * 6) + J)
                        .x = .x - ParticleOffsetX
                        .y = .y - ParticleOffsetY
                        .color = D3DColorARGB(Alpha, ColGraR, ColGraG, ColGraB)
                    End With

                Next J

                BloodCount = BloodCount + 1
            Else
                Engine_Blood_Erase I
            End If
        End If

    Next I

    D3DDevice.SetTexture 0, Tex
    D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(Alpha, 0, 0, 0)

    'Check if any blood was found in use

    If BloodCount = 0 Then Exit Sub

    'Create the vertex buffer
    Set BloodVB = D3DDevice.CreateVertexBuffer(28 * BloodCount * 6, 0, FVF, D3DPOOL_MANAGED)
    D3DVertexBuffer8SetData BloodVB, 0, 28 * BloodCount * 6, 0, BloodVL(1)

    'Draw the blood
    D3DDevice.SetStreamSource 0, BloodVB, 28

    D3DDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, BloodCount * 2

End Sub




Public Function Engine_SPtoTPX(ByVal x As Long) As Long
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
    Engine_SPtoTPX = UserPos.x + x \ TilePixelWidth - 1024 \ 2
End Function

Public Function Engine_SPtoTPY(ByVal y As Long) As Long
'************************************************************
'Screen Position to Tile Position
'Takes the screen pixel position and returns the tile position
'************************************************************
    Engine_SPtoTPY = UserPos.y + y \ TilePixelHeight - 768 \ 2
End Function




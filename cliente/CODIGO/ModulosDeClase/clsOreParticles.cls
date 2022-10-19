VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOreParticles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type Particle
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Private Type particle_group
    ACTIVE As Boolean
    ID As Long
    MAP_X As Integer
    MAP_Y As Integer
    char_index As Long
    frame_counter As Single
    frame_speed As Single
    stream_type As Byte
    particle_stream() As Particle
    particle_count As Long
    grh_index_list() As Long
    grh_index_count As Long
    alpha_blend As Boolean
    alive_counter As Long
    never_die As Boolean
    x1 As Integer
    x2 As Integer
    y1 As Integer
    y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    rgb_list(0 To 3) As Long
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Dim base_tile_size As Integer
 
Dim particle_group_list() As particle_group
Dim particle_group_count As Long
Dim particle_group_last As Long
 
Private Function Particle_Group_Next_Open() As Long
 
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until particle_group_list(LoopC).ACTIVE = False
        If LoopC = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Particle_Group_Next_Open = LoopC
Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function
 
Public Function Particle_Group_Create(ByVal MAP_X As Integer, ByVal MAP_Y As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
   
 
    If (MAP_X <> -1) And (MAP_Y <> -1) Then
    If Map_Particle_Group_Get(MAP_X, MAP_Y) = 0 Then
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, MAP_X, MAP_Y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, MAP_X, MAP_Y, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If
    End If
End Function
 
Private Function Particle_Group_Make(ByVal particle_group_index As Long, ByVal MAP_X As Integer, ByVal MAP_Y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
 
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
 
    particle_group_list(particle_group_index).ACTIVE = True
 
    If (MAP_X <> -1) And (MAP_Y <> -1) Then
        particle_group_list(particle_group_index).MAP_X = MAP_X
        particle_group_list(particle_group_index).MAP_Y = MAP_Y
    End If
 
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
 
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
 
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
   
    particle_group_list(particle_group_index).stream_type = stream_type
 
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
   
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
   
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    MapData(MAP_X, MAP_Y).particle_group = particle_group_index
End Function
 
Private Function Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef rgb_list() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional grh_resizex As Integer, Optional grh_resizey As Integer)
 
    If no_move = False Then
                If temp_particle.alive_counter = 0 Then
                    InitGrh temp_particle.Grh, grh_index, alpha_blend
                    temp_particle.X = RandomNumber(x1, x2)
                    temp_particle.Y = RandomNumber(y1, y2)
                    temp_particle.vector_x = RandomNumber(vecx1, vecx2)
                    temp_particle.vector_y = RandomNumber(vecy1, vecy2)
                    temp_particle.angle = angle
                    temp_particle.alive_counter = RandomNumber(life1, life2)
                    temp_particle.friction = fric
                Else
 
                    If gravity = True Then
                        temp_particle.vector_y = temp_particle.vector_y + grav_strength
                        If temp_particle.Y > 0 Then
                            temp_particle.vector_y = bounce_strength
                        End If
                    End If
 
                   If spin = True Then temp_particle.Grh.angle = temp_particle.Grh.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)
                    If temp_particle.angle >= 360 Then
                        temp_particle.angle = 0
                    End If
                               
                    If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
                    If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
                End If
 
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)
 
         temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If
   
    If grh_resize = True Then
 
        If temp_particle.Grh.GrhIndex Then
             Grh_Render_Advance temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(), True, True, alpha_blend
            Exit Function
        End If
    End If
 
    If temp_particle.Grh.GrhIndex Then
        Grh_Render temp_particle.Grh, _
                temp_particle.X + screen_x, _
                temp_particle.Y + screen_y, _
                rgb_list(), 1, False, True _
                , True
    End If
End Function

Public Function Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer)
 
    Dim LoopC As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean
 
    If UserMinHP = 0 And GoingHome <> 1 Then
        temp_rgb(0) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(1) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(2) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
        temp_rgb(3) = D3DColorARGB(particle_group_list(particle_group_index).alpha_blend, 255, 255, 255)
    Else
        temp_rgb(0) = particle_group_list(particle_group_index).rgb_list(0)
        temp_rgb(1) = particle_group_list(particle_group_index).rgb_list(1)
        temp_rgb(2) = particle_group_list(particle_group_index).rgb_list(2)
        temp_rgb(3) = particle_group_list(particle_group_index).rgb_list(3)
    End If
    
    
   
    
       
    If particle_group_list(particle_group_index).alive_counter Then
 
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If
 
        For LoopC = 1 To particle_group_list(particle_group_index).particle_count
 
            Particle_Render particle_group_list(particle_group_index).particle_stream(LoopC), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).x1, particle_group_list(particle_group_index).y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).x2, _
                            particle_group_list(particle_group_index).y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
                           
        Next LoopC
       
        If no_move = False Then
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If
   
    Else
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Function
 
Private Function Particle_Group_Destroy(ByVal particle_group_index As Long)
 
    Dim temp As particle_group
    Dim I As Integer
   
    If particle_group_list(particle_group_index).MAP_X > 0 And particle_group_list(particle_group_index).MAP_Y > 0 Then
        MapData(particle_group_list(particle_group_index).MAP_X, particle_group_list(particle_group_index).MAP_Y).particle_group = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            For I = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
                If charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = particle_group_index Then
                    charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = 0
                    Exit For
                End If
            Next I
        End If
    End If
   
    particle_group_list(particle_group_index) = temp
 
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).ACTIVE
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Function
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Function
 
Private Function Char_Particle_Group_Make(ByVal particle_group_index As Long, ByVal char_index As Integer, ByVal particle_char_index As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)
                               
 
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1
 
    particle_group_list(particle_group_index).ACTIVE = True
 
    particle_group_list(particle_group_index).char_index = char_index
 
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)
 
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If
 
    particle_group_list(particle_group_index).alpha_blend = alpha_blend
 
    particle_group_list(particle_group_index).stream_type = stream_type
 
    particle_group_list(particle_group_index).frame_speed = frame_speed
   
    particle_group_list(particle_group_index).x1 = x1
    particle_group_list(particle_group_index).y1 = y1
    particle_group_list(particle_group_index).x2 = x2
    particle_group_list(particle_group_index).y2 = y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2
 
    particle_group_list(particle_group_index).rgb_list(0) = rgb_list(0)
    particle_group_list(particle_group_index).rgb_list(1) = rgb_list(1)
    particle_group_list(particle_group_index).rgb_list(2) = rgb_list(2)
    particle_group_list(particle_group_index).rgb_list(3) = rgb_list(3)
   
    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey
 
    particle_group_list(particle_group_index).ID = ID
 
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)
   
    charlist(char_index).particle_group(particle_char_index) = particle_group_index
End Function

Public Function Char_Particle_Group_Create(ByVal char_index As Integer, ByRef grh_index_list() As Long, ByRef rgb_list() As Long, _
                                        Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                        Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                        Optional ByVal frame_speed As Single = 0.5, Optional ByVal ID As Long, _
                                        Optional ByVal x1 As Integer, Optional ByVal y1 As Integer, Optional ByVal angle As Integer, _
                                        Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                        Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                        Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                        Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                        Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                        Optional bounce_strength As Long, Optional ByVal x2 As Integer, Optional ByVal y2 As Integer, _
                                        Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                        Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                        Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                        Optional grh_resizex As Integer, Optional grh_resizey As Integer) As Long
 
    Dim char_part_free_index As Integer
   
    If Not Char_Check(char_index) Then Exit Function
    char_part_free_index = Char_Particle_Group_Next_Open(char_index)
   
    If char_part_free_index > 0 Then
        Char_Particle_Group_Create = Particle_Group_Next_Open
        Char_Particle_Group_Make Char_Particle_Group_Create, char_index, char_part_free_index, particle_count, stream_type, grh_index_list(), rgb_list(), alpha_blend, alive_counter, frame_speed, ID, x1, y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, x2, y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
       
    End If
 
End Function
 
Private Function Char_Particle_Group_Next_Open(ByVal char_index As Integer) As Integer
 
On Error GoTo ErrorHandler:
    Dim LoopC As Long
   
    LoopC = 1
    Do Until charlist(char_index).particle_group(LoopC) = 0
        If LoopC = charlist(char_index).particle_count Then
            Char_Particle_Group_Next_Open = charlist(char_index).particle_count + 1
            charlist(char_index).particle_count = Char_Particle_Group_Next_Open
            ReDim Preserve charlist(char_index).particle_group(1 To Char_Particle_Group_Next_Open) As Long
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop
   
    Char_Particle_Group_Next_Open = LoopC
 
Exit Function
 
ErrorHandler:
    charlist(char_index).particle_count = 1
    ReDim charlist(char_index).particle_group(1 To 1) As Long
    Char_Particle_Group_Next_Open = 1
 
End Function
 
Private Function Map_Particle_Group_Get(ByVal MAP_X As Integer, ByVal MAP_Y As Integer) As Long
 
    If InMapBounds(MAP_X, MAP_Y) Then
        Map_Particle_Group_Get = MapData(MAP_X, MAP_Y).particle_group
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
 
Private Function Char_Check(ByVal char_index As Integer) As Boolean
 
    If char_index > 0 And char_index <= LastChar Then
        Char_Check = (charlist(char_index).Heading > 0)
    End If
   
End Function
 
Private Sub Grh_Render_Advance(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByVal Height As Integer, ByVal Width As Integer, ByRef rgb_list() As Long, Optional ByVal h_center As Boolean, Optional ByVal v_center As Boolean, Optional ByVal alpha_blend As Boolean = False)
 
    Dim tile_width As Integer
    Dim tile_height As Integer
    Dim grh_index As Long
 
    If Grh.Started Then
        Grh.FrameCounter = Grh.FrameCounter + (timerTicksPerFrame * Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = 1
        End If
    End If
   
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
   
    If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
        screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (base_tile_size \ 2)) + base_tile_size \ 2
    End If
   
    If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
        screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * base_tile_size) + base_tile_size
    End If
    
    Device_Box_Textured_Render_Advance grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).PixelWidth, GrhData(grh_index).PixelHeight, _
        rgb_list, _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        Width, Height, alpha_blend, Grh.angle
 
End Sub
 
Private Sub Grh_Render(ByRef Grh As Grh, ByVal screen_x As Integer, ByVal screen_y As Integer, ByRef rgb_list() As Long, Optional ByVal h_centered As Boolean = True, Optional ByVal v_centered As Boolean = True, Optional ByVal alpha_blend As Boolean = False, Optional ByVal KillAnim As Boolean = 0)
 
On Error Resume Next
 
    Dim grh_index As Long
   
    If Grh.GrhIndex = 0 Then Exit Sub
 
    If Grh.Started = 1 Then
        Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
        If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
            Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
            If Grh.Loops <> -1 Then
                If Grh.Loops > 0 Then
                    Grh.Loops = Grh.Loops - 1
                Else
                    Grh.Started = 0
                End If
            End If
        End If
    End If
 
    If Grh.FrameCounter = 0 Then Grh.FrameCounter = 1
    grh_index = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    If grh_index <= 0 Then Exit Sub
    If GrhData(grh_index).FileNum = 0 Then Exit Sub
       
    If h_centered Then
        If GrhData(Grh.GrhIndex).TileWidth <> 1 Then
            screen_x = screen_x - Int(GrhData(Grh.GrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If
    End If
   
    If v_centered Then
        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            screen_y = screen_y - Int(GrhData(Grh.GrhIndex).TileHeight * 32) + 32
        End If
    End If
    
    Device_Box_Textured_Render grh_index, _
        screen_x, screen_y, _
        GrhData(grh_index).PixelWidth, GrhData(grh_index).PixelHeight, _
        rgb_list(), _
        GrhData(grh_index).sX, GrhData(grh_index).sY, _
        alpha_blend, Grh.angle
     
End Sub
 
Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, ByVal dest_width As Integer, Optional ByVal dest_height As Integer, _
                                            Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
 
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEXDos
    Static light_value(0 To 3) As Long
    Dim srdesc As D3DSURFACE_DESC
    Dim texture As Direct3DTexture8
    
    If GrhIndex = 0 Then Exit Sub
    Set texture = SurfaceDB.Surface(CLng(GrhData(GrhIndex).FileNum))
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
 
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .right = src_x + src_width
        .top = src_y
    End With
 
    Dim texwidth As Long, texheight As Long
    texture.GetLevelDesc 0, srdesc
    texwidth = srdesc.Width
    texheight = srdesc.Height
 
    With dest_rect
        .bottom = dest_y + dest_height
        .left = dest_x
        .right = dest_x + dest_width
        .top = dest_y
    End With
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, angle
 
    D3DDevice.SetTexture 0, texture
    
    If alpha_blend Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0)
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
       D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
       D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
 
End Sub
 
 Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEXDos, ByRef dest As RECT, ByRef src As RECT, ByRef rgb_list() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'
' * v1      * v3
' |\        |
' |  \      |
' |    \    |
' |      \  |
' |        \|
' * v0      * v2
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim temp As Single
    
    If angle > 0 Then
        'Center coordinates on screen of the square
        x_center = dest.left + (dest.right - dest.left) / 2
        y_center = dest.top + (dest.bottom - dest.top) / 2
        
        'Calculate radius
        radius = Sqr((dest.right - x_center) ^ 2 + (dest.bottom - y_center) ^ 2)
        
        'Calculate left and right points
        temp = (dest.right - x_center) / radius
        right_point = Atn(temp / Sqr(-temp * temp + 1))
        left_point = PI - right_point
    End If
    
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If
    
    
    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, src.left / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.left
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If
    
    
    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, src.left / Textures_Width, src.top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.right
        y_Cor = dest.bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If
    
    
    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, (src.right + 1) / Textures_Width, (src.bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.right
        y_Cor = dest.top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If
    
    
    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, (src.right + 1) / Textures_Width, src.top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, rgb_list(3), 0, 1, 1)
    End If

End Sub
Private Sub Device_Box_Textured_Render(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                            ByVal src_height As Integer, ByRef rgb_list() As Long, ByVal src_x As Integer, _
                                            ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single)
 
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEXDos
    Static light_value(0 To 3) As Long
    Dim srdesc As D3DSURFACE_DESC
    Dim texture As Direct3DTexture8
    
    If GrhIndex = 0 Then Exit Sub
    Set texture = SurfaceDB.Surface(CLng(GrhData(GrhIndex).FileNum))
    
    light_value(0) = rgb_list(0)
    light_value(1) = rgb_list(1)
    light_value(2) = rgb_list(2)
    light_value(3) = rgb_list(3)
 
    With src_rect
        .bottom = src_y + src_height
        .left = src_x
        .right = src_x + src_width
        .top = src_y
    End With
 
    Dim texwidth As Long, texheight As Long
    texture.GetLevelDesc 0, srdesc
    texwidth = srdesc.Width
    texheight = srdesc.Height
 
    With dest_rect
        .bottom = dest_y + src_height
        .left = dest_x
        .right = dest_x + src_width
        .top = dest_y
    End With
    
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), texwidth, texheight, angle
 
    D3DDevice.SetTexture 0, texture
    
    If alpha_blend Then
        D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If
   
    D3DDevice.SetRenderState D3DRS_TEXTUREFACTOR, D3DColorARGB(255, 0, 0, 0)
  D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))
   
    If alpha_blend Then
       D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
       D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If
 
 
End Sub


Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal z As Single, _
                                            ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                            ByVal tv As Single) As TLVERTEXDos
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'**************************************************************
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.z = z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function






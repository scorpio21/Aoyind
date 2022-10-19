Attribute VB_Name = "modOREParticles"
Private Type RGB
    R As Long
    G As Long
    b As Long
End Type
 
Private Type Stream
    Name As String
    NumOfParticles As Long
    NumGrhs As Long
    ID As Long
    x1 As Long
    y1 As Long
    x2 As Long
    y2 As Long
    angle As Long
    vecx1 As Long
    vecx2 As Long
    vecy1 As Long
    vecy2 As Long
    life1 As Long
    life2 As Long
    friction As Long
    spin As Byte
    spin_speedL As Single
    spin_speedH As Single
    AlphaBlend As Byte
    gravity As Byte
    grav_strength As Long
    bounce_strength As Long
    XMove As Byte
    YMove As Byte
    move_x1 As Long
    move_x2 As Long
    move_y1 As Long
    move_y2 As Long
    grh_list() As Long
    colortint(0 To 3) As RGB
    Speed As Single
    life_counter As Long
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
 
Private TotalStreams As Long
Private StreamData() As Stream
 
Sub CargarParticulas()
Dim StreamFile As String
Dim LoopC As Long
Dim I As Long
Dim GrhListing As String
Dim TempSet As String
Dim ColorSet As Long
   
StreamFile = App.path & "\INIT\Particles.ini"
TotalStreams = Val(GetVar(StreamFile, "INIT", "Total"))
 
ReDim StreamData(1 To TotalStreams) As Stream
 
    For LoopC = 1 To TotalStreams
        StreamData(LoopC).Name = GetVar(StreamFile, Val(LoopC), "Name")
        StreamData(LoopC).NumOfParticles = GetVar(StreamFile, Val(LoopC), "NumOfParticles")
        StreamData(LoopC).x1 = GetVar(StreamFile, Val(LoopC), "X1")
        StreamData(LoopC).y1 = GetVar(StreamFile, Val(LoopC), "Y1")
        StreamData(LoopC).x2 = GetVar(StreamFile, Val(LoopC), "X2")
        StreamData(LoopC).y2 = GetVar(StreamFile, Val(LoopC), "Y2")
        StreamData(LoopC).angle = GetVar(StreamFile, Val(LoopC), "Angle")
        StreamData(LoopC).vecx1 = GetVar(StreamFile, Val(LoopC), "VecX1")
        StreamData(LoopC).vecx2 = GetVar(StreamFile, Val(LoopC), "VecX2")
        StreamData(LoopC).vecy1 = GetVar(StreamFile, Val(LoopC), "VecY1")
        StreamData(LoopC).vecy2 = GetVar(StreamFile, Val(LoopC), "VecY2")
        StreamData(LoopC).life1 = GetVar(StreamFile, Val(LoopC), "Life1")
        StreamData(LoopC).life2 = GetVar(StreamFile, Val(LoopC), "Life2")
        StreamData(LoopC).friction = GetVar(StreamFile, Val(LoopC), "Friction")
        StreamData(LoopC).spin = GetVar(StreamFile, Val(LoopC), "Spin")
        StreamData(LoopC).spin_speedL = GetVar(StreamFile, Val(LoopC), "Spin_SpeedL")
        StreamData(LoopC).spin_speedH = GetVar(StreamFile, Val(LoopC), "Spin_SpeedH")
        StreamData(LoopC).AlphaBlend = GetVar(StreamFile, Val(LoopC), "AlphaBlend")
        StreamData(LoopC).gravity = GetVar(StreamFile, Val(LoopC), "Gravity")
        StreamData(LoopC).grav_strength = GetVar(StreamFile, Val(LoopC), "Grav_Strength")
        StreamData(LoopC).bounce_strength = GetVar(StreamFile, Val(LoopC), "Bounce_Strength")
        StreamData(LoopC).XMove = GetVar(StreamFile, Val(LoopC), "XMove")
        StreamData(LoopC).YMove = GetVar(StreamFile, Val(LoopC), "YMove")
        StreamData(LoopC).move_x1 = GetVar(StreamFile, Val(LoopC), "move_x1")
        StreamData(LoopC).move_x2 = GetVar(StreamFile, Val(LoopC), "move_x2")
        StreamData(LoopC).move_y1 = GetVar(StreamFile, Val(LoopC), "move_y1")
        StreamData(LoopC).move_y2 = GetVar(StreamFile, Val(LoopC), "move_y2")
        StreamData(LoopC).life_counter = GetVar(StreamFile, Val(LoopC), "life_counter")
        StreamData(LoopC).Speed = Val(GetVar(StreamFile, Val(LoopC), "Speed"))
        StreamData(LoopC).grh_resize = Val(GetVar(StreamFile, Val(LoopC), "resize"))
        StreamData(LoopC).grh_resizex = Val(GetVar(StreamFile, Val(LoopC), "rx"))
        StreamData(LoopC).grh_resizey = Val(GetVar(StreamFile, Val(LoopC), "ry"))
        StreamData(LoopC).NumGrhs = GetVar(StreamFile, Val(LoopC), "NumGrhs")
       
        ReDim StreamData(LoopC).grh_list(1 To StreamData(LoopC).NumGrhs)
        GrhListing = GetVar(StreamFile, Val(LoopC), "Grh_List")
       
        For I = 1 To StreamData(LoopC).NumGrhs
            StreamData(LoopC).grh_list(I) = ReadField(str(I), GrhListing, 44)
        Next I
        StreamData(LoopC).grh_list(I - 1) = StreamData(LoopC).grh_list(I - 1)
        For ColorSet = 1 To 4
            TempSet = GetVar(StreamFile, Val(LoopC), "ColorSet" & ColorSet)
            StreamData(LoopC).colortint(ColorSet - 1).R = ReadField(1, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).G = ReadField(2, TempSet, 44)
            StreamData(LoopC).colortint(ColorSet - 1).b = ReadField(3, TempSet, 44)
        Next ColorSet
    Next LoopC
 
End Sub
 
Public Function General_Particle_Create(ByVal ParticulaInd As Long, ByVal x As Integer, ByVal y As Integer, Optional ByVal particle_life As Long = 0) As Long
   
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).b)
 
General_Particle_Create = ParticlesORE.Particle_Group_Create(x, y, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function
 
Public Function General_Char_Particle_Create(ByVal ParticulaInd As Long, ByVal char_index As Integer, Optional ByVal particle_life As Long = 0) As Long
 
Dim rgb_list(0 To 3) As Long
rgb_list(0) = RGB(StreamData(ParticulaInd).colortint(0).R, StreamData(ParticulaInd).colortint(0).G, StreamData(ParticulaInd).colortint(0).b)
rgb_list(1) = RGB(StreamData(ParticulaInd).colortint(1).R, StreamData(ParticulaInd).colortint(1).G, StreamData(ParticulaInd).colortint(1).b)
rgb_list(2) = RGB(StreamData(ParticulaInd).colortint(2).R, StreamData(ParticulaInd).colortint(2).G, StreamData(ParticulaInd).colortint(2).b)
rgb_list(3) = RGB(StreamData(ParticulaInd).colortint(3).R, StreamData(ParticulaInd).colortint(3).G, StreamData(ParticulaInd).colortint(3).b)
 
General_Char_Particle_Create = ParticlesORE.Char_Particle_Group_Create(char_index, StreamData(ParticulaInd).grh_list, rgb_list(), StreamData(ParticulaInd).NumOfParticles, ParticulaInd, _
    StreamData(ParticulaInd).AlphaBlend, IIf(particle_life = 0, StreamData(ParticulaInd).life_counter, particle_life), StreamData(ParticulaInd).Speed, , StreamData(ParticulaInd).x1, StreamData(ParticulaInd).y1, StreamData(ParticulaInd).angle, _
    StreamData(ParticulaInd).vecx1, StreamData(ParticulaInd).vecx2, StreamData(ParticulaInd).vecy1, StreamData(ParticulaInd).vecy2, _
    StreamData(ParticulaInd).life1, StreamData(ParticulaInd).life2, StreamData(ParticulaInd).friction, StreamData(ParticulaInd).spin_speedL, _
    StreamData(ParticulaInd).gravity, StreamData(ParticulaInd).grav_strength, StreamData(ParticulaInd).bounce_strength, StreamData(ParticulaInd).x2, _
    StreamData(ParticulaInd).y2, StreamData(ParticulaInd).XMove, StreamData(ParticulaInd).move_x1, StreamData(ParticulaInd).move_x2, StreamData(ParticulaInd).move_y1, _
    StreamData(ParticulaInd).move_y2, StreamData(ParticulaInd).YMove, StreamData(ParticulaInd).spin_speedH, StreamData(ParticulaInd).spin, StreamData(ParticulaInd).grh_resize, StreamData(ParticulaInd).grh_resizex, StreamData(ParticulaInd).grh_resizey)
 
End Function






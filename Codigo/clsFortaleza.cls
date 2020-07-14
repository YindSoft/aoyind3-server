VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFortaleza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%   Creado por El Yind   %%%%%
'%%%%%       30/01/2012       %%%%%
'%%%%% Build: 2               %%%%%
'%%%%% Modificado: 14/02/2012 %%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit
Private Id As Integer
Private Nombre As String
Private X As Integer, Y As Integer
Private SpawnPos As WorldPos
Private Fecha As Date
Private pIdClan As Integer

Private Const WAV_CONQUISTA As Integer = 219

Private Const NPC_CAZA As Integer = 625
Private Const NPC_CLERO As Integer = 626
Private Const NPC_PALA As Integer = 624
Private Const NPC_GUERRE As Integer = 627

Private Const CANT_NPC As Integer = 8
Private Const CANT_NPC_CONTRATABLES As Integer = 2

Private Const GUERRE1 As Integer = 1
Private Const GUERRE2 As Integer = 2
Private Const CLERO1 As Integer = 3
Private Const CLERO2 As Integer = 4
Private Const PALA1 As Integer = 5
Private Const PALA2 As Integer = 6
Private Const CAZA1 As Integer = 7
Private Const CAZA2 As Integer = 8
Private NPC(1 To CANT_NPC) As Integer
Private NPC_Pos(1 To CANT_NPC) As WorldPos
Private Destino(1 To CANT_NPC) As Position
Private NPC_Num(1 To CANT_NPC) As Integer
Private NPC_Path(1 To CANT_NPC) As clsTree
Private NPC_Respawn(1 To CANT_NPC) As Date

Private ReyIndex As Integer
Private NPCRey As Integer

Private Const HECHI_REMO As Integer = 10
Private Const HECHI_INMO As Integer = 24
Private Const HECHI_TORMENTA As Integer = 15
Private Const HECHI_DESCARGA As Integer = 23
Private Const HECHI_REMOINVI As Integer = 38
Private Const HECHI_CURAR As Integer = 5

Private Const TIEMPO_RESPAWN As Integer = 90 'segundos
Private Const TIEMPO_AVISOS As Integer = 60 'segundos
Private UltimoAviso As Date

Private EsConquistable As Boolean
Private DescOriginal As String

Public Sub Init(ByVal mId As Integer)
Dim Datos As clsMySQLRecordSet
Dim i As Integer
i = mySQL.SQLQuery("SELECT * FROM fortalezas WHERE Id=" & mId, Datos)
Id = mId
Nombre = Datos("Nombre")
X = Datos("X")
Y = Datos("Y")
SpawnPos.map = 1
SpawnPos.X = Datos("SpawnX")
SpawnPos.Y = Datos("SpawnY")
Fecha = Datos("Fecha")
IdClan = Datos("IdClan")
NPCRey = Datos("NPCRey")
'Posiciones por defecto de los npc de defensa
NPC_Pos(CAZA1).X = X - 2
NPC_Pos(CAZA1).Y = Y - 26
NPC_Pos(CAZA2).X = X + 4
NPC_Pos(CAZA2).Y = Y - 26
NPC_Pos(CLERO1).X = X - 1
NPC_Pos(CLERO1).Y = Y
NPC_Pos(CLERO2).X = X + 2
NPC_Pos(CLERO2).Y = Y
NPC_Pos(PALA1).X = X - 3
NPC_Pos(PALA1).Y = Y + 2
NPC_Pos(PALA2).X = X + 4
NPC_Pos(PALA2).Y = Y + 2
NPC_Pos(GUERRE1).X = X - 1
NPC_Pos(GUERRE1).Y = Y + 3
NPC_Pos(GUERRE2).X = X + 2
NPC_Pos(GUERRE2).Y = Y + 3

NPC_Num(CAZA1) = NPC_CAZA
NPC_Num(CAZA2) = NPC_CAZA
NPC_Num(PALA1) = NPC_PALA
NPC_Num(PALA2) = NPC_PALA
NPC_Num(CLERO1) = NPC_CLERO
NPC_Num(CLERO2) = NPC_CLERO
NPC_Num(GUERRE1) = NPC_GUERRE
NPC_Num(GUERRE2) = NPC_GUERRE


For i = 1 To CANT_NPC
    'Set NPC_Path(i) = New clsTree
    NPC_Pos(i).map = 1
    NPC(i) = SpawnNpc(NPC_Num(i), NPC_Pos(i), False, True, 0)
    Npclist(NPC(i)).flags.Faccion = i
    'Call BuscarRuta(i, NPC(i))
Next i
Dim Pos As WorldPos
Pos.map = 1
Pos.X = X
Pos.Y = Y
ReyIndex = SpawnNpc(NPCRey, Pos, True, False, 0)
DescOriginal = Npclist(ReyIndex).desc
End Sub


Private Sub UsuarioContinuo(ByVal NpcIndex As Integer, ByRef tHeading As Byte, ByRef tmpIndex As Integer)
With Npclist(NpcIndex)
    If ChaseUser(MapData(.Pos.map, .Pos.X - 1, .Pos.Y).UserIndex, NpcIndex, True) Then
        tmpIndex = MapData(.Pos.map, .Pos.X - 1, .Pos.Y).UserIndex
        tHeading = eHeading.WEST
    ElseIf ChaseUser(MapData(.Pos.map, .Pos.X + 1, .Pos.Y).UserIndex, NpcIndex, True) Then
        tmpIndex = MapData(.Pos.map, .Pos.X + 1, .Pos.Y).UserIndex
        tHeading = eHeading.EAST
    ElseIf ChaseUser(MapData(.Pos.map, .Pos.X, .Pos.Y - 1).UserIndex, NpcIndex, True) Then
        tmpIndex = MapData(.Pos.map, .Pos.X, .Pos.Y - 1).UserIndex
        tHeading = eHeading.NORTH
    ElseIf ChaseUser(MapData(.Pos.map, .Pos.X, .Pos.Y + 1).UserIndex, NpcIndex, True) Then
        tmpIndex = MapData(.Pos.map, .Pos.X, .Pos.Y + 1).UserIndex
        tHeading = eHeading.SOUTH
    End If
End With
End Sub

Public Sub MoverNPC(ByVal NpcIndex As Integer)
Dim index As Integer
Dim Victima As Integer
Dim tmpIndex As Integer
Dim tHeading As Byte
Dim DifX As Integer, DifY As Integer
Dim nPos As WorldPos
Dim i As Integer
With Npclist(NpcIndex)
index = .flags.Faccion

Victima = SeleccionarVictima(NpcIndex)

If Victima > 0 Then

    Select Case .Numero
        Case NPC_GUERRE
            Call UsuarioContinuo(NpcIndex, tHeading, tmpIndex)
            If .flags.Paralizado = 1 And tHeading <> .Char.Heading Then tHeading = 0
            

            
            If tHeading > 0 Then
                'Si tengo a un boludo al lado mientras estoy inmo lo ataco.
                If tHeading <> .Char.Heading Then Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                Call NpcAtacaUser(NpcIndex, tmpIndex)
                Exit Sub
            ElseIf UserList(Victima).flags.invisible = 0 And UserList(Victima).flags.Oculto = 0 Then
                'Voy hasta el usuario
                If Destino(index).X <> UserList(Victima).Pos.X Or Destino(index).Y <> UserList(Victima).Pos.Y Then
                    Destino(index).X = UserList(Victima).Pos.X
                    Destino(index).Y = UserList(Victima).Pos.Y
                    Call BuscarRuta(index, NpcIndex)
                End If
            Else
                Call IrFormacion(index, NpcIndex)
            End If
        Case NPC_CAZA
            If UserList(Victima).flags.invisible = 0 And UserList(Victima).flags.Oculto = 0 And Abs(UserList(Victima).Pos.X - .Pos.X) <= RANGO_VISION_X And Abs(UserList(Victima).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                DifY = .Pos.Y - UserList(Victima).Pos.Y
                DifX = .Pos.X - UserList(Victima).Pos.X
                If Abs(DifY) >= Abs(DifX) Then
                    If DifY >= 0 Then
                        tHeading = eHeading.NORTH
                    Else
                        tHeading = eHeading.SOUTH
                    End If
                Else
                    If DifX <= 0 Then
                        tHeading = eHeading.EAST
                    Else
                        tHeading = eHeading.WEST
                    End If
                End If
                'Si el npc esta paralizado no puede girar y si no esta apuntando para el lado del usuario no dispara
                If tHeading <> .Char.Heading And .flags.Paralizado = 0 Then Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                If RandomNumber(1, 3) = 1 And tHeading = .Char.Heading Then  'Un poco de ayuda sino te parten :P
                    Call NpcLanzaSpellSobreUser(NpcIndex, Victima, .Spells(1))
                End If
            End If
        Case NPC_PALA
            Call UsuarioContinuo(NpcIndex, tHeading, tmpIndex)
            If .flags.Paralizado = 1 And tHeading <> .Char.Heading Then tHeading = 0
            If tHeading > 0 Then
                'Si tengo a un boludo al lado mientras estoy inmo lo ataco.
                If tHeading <> .Char.Heading Then Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, tHeading)
                Call NpcAtacaUser(NpcIndex, tmpIndex)
                Exit Sub
            ElseIf .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                tmpIndex = SeleccionarVictima(NpcIndex, True)
                If tmpIndex > 0 Then
                    If UserList(tmpIndex).flags.invisible = 1 Or UserList(tmpIndex).flags.Oculto = 1 Or Abs(UserList(tmpIndex).Pos.X - .Pos.X) >= RANGO_VISION_X Or Abs(UserList(tmpIndex).Pos.Y - .Pos.Y) >= RANGO_VISION_Y Then
                        tmpIndex = 0
                    End If
                End If
                If tmpIndex > 0 Then
                    'Voy hasta el usuario
                    Call AtacarConMagia(NpcIndex, tmpIndex)
                    If Destino(index).X <> UserList(tmpIndex).Pos.X Or Destino(index).Y <> UserList(tmpIndex).Pos.Y Then
                        Destino(index).X = UserList(tmpIndex).Pos.X
                        Destino(index).Y = UserList(tmpIndex).Pos.Y
                        Call BuscarRuta(index, NpcIndex)
                    End If
                Else
                    If UserList(Victima).flags.invisible = 0 And UserList(Victima).flags.Oculto = 0 And Abs(UserList(Victima).Pos.X - .Pos.X) <= RANGO_VISION_X And Abs(UserList(Victima).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        Call AtacarConMagia(NpcIndex, Victima)
                    End If
                    
                    tHeading = .Char.Heading 'Movimiento estilo mercader
                    If RandomNumber(1, 2) = 1 Then
                        tHeading = RandomNumber(1, 4)
                        If .Char.Heading = tHeading Then tHeading = RandomNumber(1, 4)
                    End If
                    
                    nPos = .Pos
                    Call HeadtoPos(tHeading, nPos)
                    
                    If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                        MoveNPCChar NpcIndex, tHeading
                    Else
                        For i = 1 To 4
                            If i <> tHeading Then
                                nPos = .Pos
                                Call HeadtoPos(i, nPos)
                                If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                                    MoveNPCChar NpcIndex, i
                                    Exit For
                                End If
                            End If
                        Next i
                    End If
                    Destino(index).X = nPos.X
                    Destino(index).Y = nPos.Y
                    Exit Sub
                End If
            End If
        Case NPC_CLERO
            Call DefenderConMagia(NpcIndex)
            If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                nPos.map = 1 'Mira la distancia al rey
                nPos.X = X
                nPos.Y = Y
                If Distancia(.Pos, nPos) < 5 Then
                    If .Pos.Y < Y Then 'Con este codigo el bicho gira en circulos sobre el rey
                        If .Pos.X < X + 2 Then
                            tHeading = eHeading.EAST
                        Else
                            tHeading = eHeading.SOUTH
                        End If
                    ElseIf .Pos.Y > Y Then
                        If .Pos.X > X - 1 Then
                            tHeading = eHeading.WEST
                        Else
                            tHeading = eHeading.NORTH
                        End If
                    Else
                        If .Pos.X < X Then
                            tHeading = eHeading.NORTH
                        ElseIf .Pos.X > X Then
                            tHeading = eHeading.SOUTH
                        End If
                    End If
                    For i = 1 To 4
                        nPos = .Pos
                        Call HeadtoPos(tHeading, nPos)
                        If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                            MoveNPCChar NpcIndex, tHeading
                            Destino(index).X = nPos.X
                            Destino(index).Y = nPos.Y
                            Exit Sub
                        End If
                        tHeading = RandomNumber(1, 4)
                    Next i
                Else
                    Call IrFormacion(index, NpcIndex)
                End If
            End If
    End Select

Else 'Vuelve a su posición
    Call IrFormacion(index, NpcIndex)
End If
If Not NPC_Path(index) Is Nothing And .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
    If Not SeguirRuta(index, NpcIndex) Then
        If Destino(index).X <> .Pos.X Or Destino(index).Y <> .Pos.Y Then
            If BuscarRuta(index, NpcIndex) Then
                'MoverNPC (NpcIndex)
            End If
        ElseIf .Char.Heading <> SOUTH Then
            .Char.Heading = SOUTH
            Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, SOUTH)
            Set NPC_Path(index) = Nothing
        End If
        
    End If
End If
End With
End Sub

Private Sub DefenderConMagia(ByVal NpcIndex As Integer)
Dim Mana As Integer
Dim XX As Integer, YY As Integer, NI As Integer
Dim UI As Integer
Dim Elemento
Dim Victima As Integer
Mana = 2000
With Npclist(NpcIndex)
    If .CanAttack = 1 And RandomNumber(1, 6) = 1 Then
    
    If RandomNumber(1, 5) = 1 Then
    For Each Elemento In .AreasInfo.Users.Items
        UI = Elemento
        If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Zonas(UserList(UI).Zona).Terreno = eTerreno.Fortaleza Then
            If ChaseUser(UI, NpcIndex, True) Then
                If UserList(UI).flags.invisible = 1 Or UserList(UI).flags.Oculto = 1 Then
                    If NpcLanzaSpellSobreUserComoUser(NpcIndex, UI, HECHI_REMOINVI, Mana) Then
                        .CanAttack = 0
                        .AttackTimer = TIMER_ATTACK * 3
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next Elemento
    End If
    
    For YY = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
        For XX = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
            If XX >= MinXBorder And XX <= MaxXBorder And YY >= MinYBorder And YY <= MaxYBorder Then
                NI = MapData(.Pos.map, XX, YY).NpcIndex
                If NI > 0 Then
                    If Npclist(NI).flags.Paralizado = 1 Or Npclist(NI).flags.Inmovilizado = 1 And RandomNumber(1, 4) = 1 Then
                        If NpcLanzaSpellSobreNpcComoUser(NpcIndex, NI, HECHI_REMO, Mana) Then
                            .CanAttack = 0
                            .AttackTimer = TIMER_ATTACK * 2
                            Exit Sub
                        End If
                    ElseIf Npclist(NI).Stats.MinHP < Npclist(NI).Stats.MaxHP Then
                        If NpcLanzaSpellSobreNpcComoUser(NpcIndex, NI, HECHI_CURAR, Mana) Then
                            .CanAttack = 0
                            .AttackTimer = TIMER_ATTACK * 2
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next XX
    Next YY
    
    End If
End With
End Sub
Private Sub AtacarConMagia(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
Dim Mana As Integer
Mana = 2000
With Npclist(NpcIndex)
    If .CanAttack = 1 Then
        If UserList(UserIndex).flags.Inmovilizado = 0 And UserList(UserIndex).flags.Paralizado = 0 And RandomNumber(1, 15) = 1 Then
            If NpcLanzaSpellSobreUserComoUser(NpcIndex, UserIndex, HECHI_INMO, Mana) Then
                .CanAttack = 0
                .AttackTimer = TIMER_ATTACK
            End If
        ElseIf RandomNumber(1, 10) = 1 Then
            If NpcLanzaSpellSobreUserComoUser(NpcIndex, UserIndex, IIf(RandomNumber(1, 3) = 1, HECHI_DESCARGA, HECHI_TORMENTA), Mana) Then
                .CanAttack = 0
                .AttackTimer = TIMER_ATTACK
            End If
        End If
    End If
End With
End Sub


Private Sub IrFormacion(ByVal index As Integer, ByVal NpcIndex As Integer)
    If Destino(index).X <> NPC_Pos(index).X Or Destino(index).Y <> NPC_Pos(index).Y Then
        Destino(index).X = NPC_Pos(index).X
        Destino(index).Y = NPC_Pos(index).Y
        Call BuscarRuta(index, NpcIndex)
    End If
End Sub

Function SeleccionarVictima(ByVal NpcIndex As Integer, Optional ByVal PrioriInmo As Boolean = False) As Integer
Dim UI As Integer
Dim Elemento
Dim Victima As Integer
Dim VictimaHP As Integer
VictimaHP = 2000
Dim nPos As WorldPos
With Npclist(NpcIndex)

    nPos.map = 1 'Mira la distancia al rey
    nPos.X = NPC_Pos(.flags.Faccion).X
    nPos.Y = NPC_Pos(.flags.Faccion).Y
    For Each Elemento In .AreasInfo.Users.Items
        UI = Elemento
        If Zonas(UserList(UI).Zona).Terreno = eTerreno.Fortaleza Then
            If ChaseUser(UI, NpcIndex, True) And ((PrioriInmo And (UserList(UI).flags.Inmovilizado = 1 Or UserList(UI).flags.Paralizado = 1)) Or PrioriInmo = False) And (.Numero <> NPC_GUERRE Or (Abs(UserList(UI).Pos.X - nPos.X) < 9 And Abs(UserList(UI).Pos.Y - nPos.Y) < 8)) Then
                If UserList(UI).Stats.MinHP < VictimaHP Then
                    VictimaHP = UserList(UI).Stats.MinHP
                    Victima = UI
                End If
                If Npclist(NpcIndex).flags.AttackedBy = UserList(UI).Name Then
                    Victima = UI
                    Exit For
                End If
            End If
        End If
    Next Elemento
End With
SeleccionarVictima = Victima
End Function

Private Function SeguirRuta(ByVal index As Integer, ByVal NpcIndex As Integer) As Boolean
Dim tHeading As Byte
Dim startX As Integer, startY As Integer
With Npclist(NpcIndex)
startX = .Pos.X
startY = .Pos.Y
If NPC_Path(index).PathStepNext(startX, startY) > 0 Then
    Dim tmpPos As WorldPos
    tmpPos.map = .Pos.map
    tmpPos.X = startX
    tmpPos.Y = startY
    tHeading = FindDirection(.Pos, tmpPos)
            
    If LegalPosNPC(tmpPos.map, tmpPos.X, tmpPos.Y, .flags.AguaValida, True) Then
        If HayAgua(tmpPos.map, tmpPos.X, tmpPos.Y) Then
            If .Char.Body <> iGalera Then 'El npc navega :P
                Call ChangeNPCChar(NpcIndex, iGalera, 0, tHeading)
            End If
        Else
            If .Char.Body = iGalera Then
                Call ChangeNPCChar(NpcIndex, .OrigBody, .OrigHead, tHeading)
            End If
        End If
        MoveNPCChar NpcIndex, tHeading
    ElseIf Destino(index).X <> tmpPos.X Or Destino(index).Y <> tmpPos.Y Then 'Si no puedo ir al lugar que tengo que ir no importa
        If BuscarRuta(index, NpcIndex) Then
            MoverNPC (NpcIndex)
        End If
    Else
        SeguirRuta = False
        Exit Function
    End If
    SeguirRuta = True
Else
    SeguirRuta = False
End If
End With
End Function
Private Function BuscarRuta(ByVal index As Integer, ByVal NpcIndex As Integer) As Boolean
Set NPC_Path(index) = New clsTree
BuscarRuta = NPC_Path(index).RunSearch(1, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y, Destino(index).X, Destino(index).Y, 1)
End Function


Public Property Get IdClan() As Integer
IdClan = pIdClan
End Property

Public Property Let IdClan(ByVal lNewValue As Integer)
pIdClan = lNewValue
End Property


Public Sub ReSpawn(ByVal Num As Byte)
NPC(Num) = 0
NPC_Respawn(Num) = Now

'Hay que comprobar cuantos bichos estan lejos del rey
If Conquistable Then
    'El castillo es conquistable
    Npclist(ReyIndex).desc = "¡Han logrado reducir todas mis defensas!, tu clan puede ser el proximo protector de esta fortaleza si así lo desean, escribe /PROTEGER y serán los nuevos guardianes de la fortaleza " & Nombre
    Call SendData(ToNPCArea, ReyIndex, PrepareMessageChatOverHead(Npclist(ReyIndex).desc, Npclist(ReyIndex).Char.CharIndex, vbWhite))
    EsConquistable = True
End If
End Sub
Public Sub CheckRespawns()
Dim i As Byte
For i = 1 To CANT_NPC
    If NPC(i) = 0 And DateDiff("s", NPC_Respawn(i), Now) >= TIEMPO_RESPAWN Then
        NPC(i) = SpawnNpc(NPC_Num(i), SpawnPos, True, True, 0)
        Npclist(NPC(i)).flags.Faccion = i
        Call BuscarRuta(i, NPC(i))
    End If
Next i

If EsConquistable Then
    If Not Conquistable Then
        EsConquistable = False
        Npclist(NPCRey).desc = DescOriginal
    End If
End If
End Sub

Private Function Conquistable() As Boolean
Dim i As Integer
Dim Lejos As Byte
For i = 1 To CANT_NPC - CANT_NPC_CONTRATABLES
    If NPC(i) = 0 Then
        Lejos = Lejos + 1
    Else
        If Npclist(NPC(i)).Pos.Y < Y - 34 Then 'Si estan yendo a la fortaleza pero no estan a una distancia cercana
            Lejos = Lejos + 1
        End If
    End If
Next i
Conquistable = Lejos = CANT_NPC - CANT_NPC_CONTRATABLES
End Function

Public Sub Proteger(ByVal UserIndex As Integer)
Dim Pos As WorldPos

    Pos.map = 1
    Pos.X = X
    Pos.Y = Y
    If Distancia(UserList(UserIndex).Pos, Pos) > 4 Then
        Call WriteConsoleMsg(UserIndex, "Debes acercarte más al rey.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).flags.Muerto = 1 Then
        Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).GuildIndex = 0 Then
        Call WriteConsoleMsg(UserIndex, "Solo los clanes pueden ser nombrados como protectores de la fortaleza.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If UserList(UserIndex).GuildIndex = IdClan Then
        Call WriteConsoleMsg(UserIndex, "La fortaleza ya esta siendo protegida por tu clan, si deseas ir hacia el muelle usa /PROTEGER OESTE o /PROTEGER ESTE", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    If EsConquistable Then
        EsConquistable = False
        Npclist(ReyIndex).desc = DescOriginal
        IdClan = UserList(UserIndex).GuildIndex
        Execute ("UPDATE fortalezas SET IdClan=" & IdClan & ", Fecha=NOW() WHERE Id=" & Id)
        Call SendData(ToNPCArea, ReyIndex, PrepareMessageChatOverHead("¡Felicitaciones! Han tomado la desición correcta. Ahora el clan " & modGuilds.GuildName(IdClan) & " es el nuevo protector de la fortaleza " & Nombre & ".", Npclist(ReyIndex).Char.CharIndex, vbWhite))
        
        Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("El clan " & modGuilds.GuildName(IdClan) & " es el nuevo protector de la fortaleza " & Nombre & ".", FontTypeNames.FONTTYPE_GUILD))
        Call SendData(SendTarget.ToAll, 0, PrepareMessagePlayWave(WAV_CONQUISTA, NO_3D_SOUND, NO_3D_SOUND))
    Else
        Call SendData(ToNPCArea, ReyIndex, PrepareMessageChatOverHead("No me doblegaré ante vosotros a menos que prueben ser un clan digno de un rey.", Npclist(ReyIndex).Char.CharIndex, vbWhite))
    End If

End Sub

Public Sub AvisarAtaque(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
If DateDiff("s", UltimoAviso, Now) >= TIEMPO_AVISOS Then
    UltimoAviso = Now
    If UserList(UserIndex).GuildIndex > 0 Then
        Call SendData(ToGuildMembers, IdClan, PrepareMessageShowMessageScroll("¡El clan " & modGuilds.GuildName(UserList(UserIndex).GuildIndex) & " está atacando la fortaleza " & Nombre & "!" & vbCrLf & "Todos los miembros del clan usen el portal del muelle" & vbCrLf & "y ayuden a protegerla.", 510, 1))
    Else
        Call SendData(ToGuildMembers, IdClan, PrepareMessageShowMessageScroll("¡" & UserList(UserIndex).Name & " está atacando la fortaleza " & Nombre & "!" & vbCrLf & "Todos los miembros del clan usen el portal del muelle" & vbCrLf & "y ayuden a protegerla.", 510, 1))
    End If
End If
End Sub

Public Sub SumUser(ByVal UserIndex As Integer)
If UserList(UserIndex).GuildIndex = IdClan Then
    If Zonas(UserList(UserIndex).Zona).Segura = 1 Then
        Call WarpUserChar(UserIndex, SpawnPos.map, SpawnPos.X, SpawnPos.Y, True)
    Else
        Call WriteConsoleMsg(UserIndex, "Debes estar en una zona segura para usar el portal del muelle.", FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "Solo los miembros del clan protector pueden utilizar el portal del muelle.", FontTypeNames.FONTTYPE_INFO)
End If
End Sub

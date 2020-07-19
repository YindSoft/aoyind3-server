Attribute VB_Name = "NPCs"
'AoYind 3.0.0
'Copyright (C) 2002 Márquez Pablo Ignacio
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez


'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'Contiene todas las rutinas necesarias para cotrolar los
'NPCs meno la rutina de AI que se encuentra en el modulo
'AI_NPCs para su mejor comprension.
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Option Explicit

Sub QuitarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
    Dim i As Integer
    
    For i = 1 To MAXMASCOTAS
      If UserList(UserIndex).MascotasIndex(i) = NpcIndex Then
         UserList(UserIndex).MascotasIndex(i) = 0
         UserList(UserIndex).MascotasType(i) = 0
         
         UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas - 1
         Exit For
      End If
    Next i
End Sub

Sub QuitarMascotaNpc(ByVal Maestro As Integer)
    Npclist(Maestro).Mascotas = Npclist(Maestro).Mascotas - 1
End Sub

Sub MuereNpc(ByVal NpcIndex As Integer, ByVal UserIndex As Integer, Optional ByVal NpcIndexAttacker As Integer = 0)
'********************************************************
'Author: Unknown
'Llamado cuando la vida de un NPC llega a cero.
'Last Modify Date: 24/01/2007
'22/06/06: (Nacho) Chequeamos si es pretoriano
'24/01/2007: Pablo (ToxicWaste): Agrego para actualización de tag si cambia de status.
'********************************************************
On Error GoTo errhandler
    Dim MiNPC As NPC
    MiNPC = Npclist(NpcIndex)
    Dim EraCriminal As Boolean
    Dim Area As Integer
   
    If (esPretoriano(NpcIndex) = 4) Then
        'Solo nos importa si fue matado en el mapa pretoriano.
        Area = Npclist(NpcIndex).Zona
        If Area = MAPA_PRETORIANO Then
            'seteamos todos estos 'flags' acorde para que cambien solos de alcoba
            Dim i As Integer
            Dim j As Integer
            Dim NPCI As Integer
        
            For i = Areas(Area).X1 To Areas(Area).X2
                For j = Areas(Area).Y1 To Areas(Area).Y2
                
                    NPCI = MapData(Npclist(NpcIndex).Pos.map, i, j).NpcIndex
                    If NPCI > 0 Then
                        If esPretoriano(NPCI) > 0 And NPCI <> NpcIndex Then
                            If Npclist(NpcIndex).Pos.X > Areas(Area).X1 + 50 Then
                                If Npclist(NPCI).Pos.X > Areas(Area).X1 + 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 1
                            Else
                                If Npclist(NPCI).Pos.X <= Areas(Area).X1 + 50 Then Npclist(NPCI).Invent.ArmourEqpSlot = 5
                            End If
                        End If
                    End If
                Next j
            Next i
            Call CrearClanPretoriano(Npclist(NpcIndex).Pos.X)
        End If
    ElseIf esPretoriano(NpcIndex) > 0 Then
        If Npclist(NpcIndex).Zona = MAPA_PRETORIANO Then
            Npclist(NpcIndex).Invent.ArmourEqpSlot = 0
            pretorianosVivos = pretorianosVivos - 1
        End If
    End If
   
    'Quitamos el npc
    Call QuitarNPC(NpcIndex)
    
    If UserIndex > 0 Then ' Lo mato un usuario?
        If MiNPC.flags.Snd3 > 0 Then
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MiNPC.flags.Snd3, MiNPC.Pos.X, MiNPC.Pos.Y))
        End If
        UserList(UserIndex).flags.TargetNPC = 0
        UserList(UserIndex).flags.TargetNpcTipo = eNPCType.Comun
        
        'El user que lo mato tiene mascotas?
        If UserList(UserIndex).NroMascotas > 0 Then
            Dim t As Integer
            For t = 1 To MAXMASCOTAS
                  If UserList(UserIndex).MascotasIndex(t) > 0 Then
                      If Npclist(UserList(UserIndex).MascotasIndex(t)).TargetNPC = NpcIndex Then
                              Call FollowAmo(UserList(UserIndex).MascotasIndex(t))
                      End If
                  End If
            Next t
        End If
        
        '[KEVIN]
        If MiNPC.flags.ExpCount > 0 Then
            If UserList(UserIndex).PartyIndex > 0 Then
                Call mdParty.ObtenerExito(UserIndex, MiNPC.flags.ExpCount, MiNPC.Pos.map, MiNPC.Pos.X, MiNPC.Pos.Y)
            Else
                UserList(UserIndex).Stats.Exp = UserList(UserIndex).Stats.Exp + MiNPC.flags.ExpCount
                If UserList(UserIndex).Stats.Exp > MAXEXP Then _
                    UserList(UserIndex).Stats.Exp = MAXEXP
                Call WriteConsoleMsg(UserIndex, "Has ganado " & MiNPC.flags.ExpCount & " puntos de experiencia.", FontTypeNames.FONTTYPE_EXP)
            End If
            MiNPC.flags.ExpCount = 0
        End If
        
        '[/KEVIN]
        Call WriteConsoleMsg(UserIndex, "¡Has matado a la criatura!", FontTypeNames.FONTTYPE_FIGHT)
        If UserList(UserIndex).Stats.NPCsMuertos < 32000 Then _
            UserList(UserIndex).Stats.NPCsMuertos = UserList(UserIndex).Stats.NPCsMuertos + 1
        
        EraCriminal = Criminal(UserIndex)
        
        If MiNPC.Stats.Alineacion = 0 Then
            If MiNPC.Numero = Guardias Then
                UserList(UserIndex).Reputacion.NobleRep = 0
                UserList(UserIndex).Reputacion.PlebeRep = 0
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + 500
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
            If MiNPC.MaestroUser = 0 Then
                UserList(UserIndex).Reputacion.AsesinoRep = UserList(UserIndex).Reputacion.AsesinoRep + vlASESINO
                If UserList(UserIndex).Reputacion.AsesinoRep > MAXREP Then _
                    UserList(UserIndex).Reputacion.AsesinoRep = MAXREP
            End If
        ElseIf MiNPC.Stats.Alineacion = 1 Then
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 2 Then
            UserList(UserIndex).Reputacion.NobleRep = UserList(UserIndex).Reputacion.NobleRep + vlASESINO / 2
            If UserList(UserIndex).Reputacion.NobleRep > MAXREP Then _
                UserList(UserIndex).Reputacion.NobleRep = MAXREP
        ElseIf MiNPC.Stats.Alineacion = 4 Then
            UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlCAZADOR
            If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
                UserList(UserIndex).Reputacion.PlebeRep = MAXREP
        End If
        If Criminal(UserIndex) And esArmada(UserIndex) Then Call ExpulsarFaccionReal(UserIndex)
        If Not Criminal(UserIndex) And esCaos(UserIndex) Then Call ExpulsarFaccionCaos(UserIndex)
        
        If EraCriminal And Not Criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        ElseIf Not EraCriminal And Criminal(UserIndex) Then
            Call RefreshCharStatus(UserIndex)
        End If
        
        Call CheckUserLevel(UserIndex)
    End If ' Userindex > 0
   
    If MiNPC.MaestroUser = 0 Then
    
        'Si lo mato un guardia no debe dropear items.
        If Npclist(NpcIndexAttacker).NPCtype <> 2 Then
            'Tiramos el oro
            Call NPCTirarOro(MiNPC)
            'Tiramos el inventario
            Call NPC_TIRAR_ITEMS(MiNPC)
        End If
        'ReSpawn o no
        Call ReSpawnNpc(MiNPC)
    End If
   
    
    
Exit Sub

errhandler:
    Call LogError("Error en MuereNpc - Error: " & Err.Number & " - Desc: " & Err.Description)
End Sub

Private Sub ResetNpcFlags(ByVal NpcIndex As Integer)
    'Clear the npc's flags
    
    With Npclist(NpcIndex).flags
        .AfectaParalisis = 0
        .AguaValida = 0
        .AttackedBy = vbNullString
        .AttackedFirstBy = vbNullString
        .BackUp = 0
        .Bendicion = 0
        .Domable = 0
        .Envenenado = 0
        .Faccion = 0
        .Follow = False
        .AtacaDoble = 0
        .LanzaSpells = 0
        .invisible = 0
        .Maldicion = 0
        .OldHostil = 0
        .OldMovement = 0
        .Paralizado = 0
        .Inmovilizado = 0
        .ReSpawn = 0
        .RespawnOrigPos = 0
        .Snd1 = 0
        .Snd2 = 0
        .Snd3 = 0
        .TierraInvalida = 0
    End With
End Sub

Private Sub ResetNpcCounters(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex).Contadores
        .Paralisis = 0
        .TiempoExistencia = 0
    End With
End Sub

Private Sub ResetNpcCharInfo(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex).Char
        .Body = 0
        .CascoAnim = 0
        .CharIndex = 0
        .FX = 0
        .Head = 0
        .Heading = 0
        .Loops = 0
        .ShieldAnim = 0
        .WeaponAnim = 0
    End With
End Sub

Private Sub ResetNpcCriatures(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroCriaturas
            .Criaturas(j).NpcIndex = 0
            .Criaturas(j).NpcName = vbNullString
        Next j
        
        .NroCriaturas = 0
    End With
End Sub

Sub ResetExpresiones(ByVal NpcIndex As Integer)
    Dim j As Long
    
    With Npclist(NpcIndex)
        For j = 1 To .NroExpresiones
            .Expresiones(j) = vbNullString
        Next j
        
        .NroExpresiones = 0
    End With
End Sub

Private Sub ResetNpcMainInfo(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        .Attackable = 0
        .CanAttack = 0
        .AttackTimer = 0
        .Comercia = 0
        .GiveEXP = 0
        .GiveGLDMin = 0
        .GiveGLDMax = 0
        .Hostile = 0
        .Zona = 0
        .InvReSpawn = 0
        .Mensaje = 0
        
        If .MaestroUser > 0 Then Call QuitarMascota(.MaestroUser, NpcIndex)
        If .MaestroNpc > 0 Then Call QuitarMascotaNpc(.MaestroNpc)
        
        .MaestroUser = 0
        .MaestroNpc = 0
        
        .PFINFO.curPos = 0
        .PFINFO.NoPath = True
        .PFINFO.PathLenght = 0
        .PFINFO.TargetNPC = 0
        .PFINFO.TargetUser = 0
        
        
        .Mascotas = 0
        .Movement = 0
        .Name = vbNullString
        .NPCtype = 0
        .Numero = 0
        .Orig.map = 0
        .Orig.X = 0
        .Orig.Y = 0
        .OrigHeading = 0
        .OrigBody = 0
        .OrigHead = 0
        .PoderAtaque = 0
        .PoderEvasion = 0
        .Pos.map = 0
        .Pos.X = 0
        .Pos.Y = 0
        .SkillDomar = 0
        .Target = 0
        .TargetNPC = 0
        .TipoItems = 0
        .Veneno = 0
        .desc = vbNullString
        
        
        Dim j As Long
        For j = 1 To .NroSpells
            .Spells(j) = 0
        Next j
    End With
    
    Call ResetNpcCharInfo(NpcIndex)
    Call ResetNpcCriatures(NpcIndex)
    Call ResetExpresiones(NpcIndex)
End Sub

Public Sub QuitarNPC(ByVal NpcIndex As Integer)

On Error GoTo errhandler

    With Npclist(NpcIndex)
        .flags.NPCActive = False
        
        If InMapBounds(.Pos.map, .Pos.X, .Pos.Y) Then
            Call EraseNPCChar(NpcIndex)
        End If
    End With
        
    'Nos aseguramos de que el inventario sea removido...
    'asi los lobos no volveran a tirar armaduras ;))
    Call ResetNpcInv(NpcIndex)
    Call ResetNpcFlags(NpcIndex)
    Call ResetNpcCounters(NpcIndex)
    
    Call ResetNpcMainInfo(NpcIndex)
    
    If NpcIndex = LastNPC Then
        Do Until Npclist(LastNPC).flags.NPCActive
            LastNPC = LastNPC - 1
            If LastNPC < 1 Then Exit Do
        Loop
    End If
        
      
    If NumNPCs <> 0 Then
        NumNPCs = NumNPCs - 1
    End If
Exit Sub

errhandler:
    Call LogError("Error en QuitarNPC")
End Sub

Private Function TestSpawnTrigger(Pos As WorldPos, Optional PuedeAgua As Boolean = False) As Boolean
    
    If LegalPos(Pos.map, Pos.X, Pos.Y, PuedeAgua) Then
        TestSpawnTrigger = _
        MapData(Pos.map, Pos.X, Pos.Y).Trigger <> 3 And _
        MapData(Pos.map, Pos.X, Pos.Y).Trigger <> 2 And _
        MapData(Pos.map, Pos.X, Pos.Y).Trigger <> 1
    End If

End Function

Sub CrearNPC(NroNPC As Integer, Area As Integer, OrigPos As WorldPos)
'Call LogTarea("Sub CrearNPC")
'Crea un NPC del tipo NRONPC

Dim Pos As WorldPos
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim Iteraciones As Long
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer

    nIndex = OpenNPC(NroNPC) 'Conseguimos un indice
    
    If nIndex > MAXNPCS Then Exit Sub
    PuedeAgua = Npclist(nIndex).flags.AguaValida
    PuedeTierra = IIf(Npclist(nIndex).flags.TierraInvalida = 1, False, True)
    
    'Necesita ser respawned en un lugar especifico
    If InMapBounds(OrigPos.map, OrigPos.X, OrigPos.Y) Then
        
        map = OrigPos.map
        X = OrigPos.X
        Y = OrigPos.Y
        Npclist(nIndex).Orig = OrigPos
        Npclist(nIndex).Pos = OrigPos
        Call CambiarOrigHeading(nIndex, MapData(map, X, Y).Trigger)
       
    ElseIf Area = 0 Then
       Call QuitarNPC(nIndex)
       Call LogError("Se intento CrearNpc con Area 0 NroNpc:" & NroNPC)
       Exit Sub
    Else
        
        Npclist(nIndex).Area = Area
        
        Pos.map = Areas(Area).mapa
        altpos.map = Areas(Area).mapa
        
        Do While Not PosicionValida
            Pos.X = RandomNumber(Areas(Area).X1, Areas(Area).X2)      'Obtenemos posicion al azar en x
            Pos.Y = RandomNumber(Areas(Area).Y1, Areas(Area).Y2)    'Obtenemos posicion al azar en y
            
            Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
            If newpos.X <> 0 And newpos.Y <> 0 Then
                altpos.X = newpos.X
                altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn, pero intentando qeu si tenía que ser en el agua, sea en el agua.)
            Else
                Call ClosestLegalPos(Pos, newpos, PuedeAgua)
                If newpos.X <> 0 And newpos.Y <> 0 Then
                    altpos.X = newpos.X
                    altpos.Y = newpos.Y     'posicion alternativa (para evitar el anti respawn)
                End If
            End If
            'Si X e Y son iguales a 0 significa que no se encontro posicion valida
            If LegalPosNPC(newpos.map, newpos.X, newpos.Y, PuedeAgua) And _
               Not HayPCarea(newpos) And TestSpawnTrigger(newpos, PuedeAgua) Then
                'Asignamos las nuevas coordenas solo si son validas
                Npclist(nIndex).Pos.map = newpos.map
                Npclist(nIndex).Pos.X = newpos.X
                Npclist(nIndex).Pos.Y = newpos.Y
                Call CambiarOrigHeading(nIndex, MapData(newpos.map, newpos.X, newpos.Y).Trigger)
                PosicionValida = True
            Else
                newpos.X = 0
                newpos.Y = 0
            
            End If
                
            'for debug
            Iteraciones = Iteraciones + 1
            If Iteraciones > MAXSPAWNATTEMPS Then
                If altpos.X <> 0 And altpos.Y <> 0 Then
                    map = altpos.map
                    X = altpos.X
                    Y = altpos.Y
                    Npclist(nIndex).Pos.map = map
                    Npclist(nIndex).Pos.X = X
                    Npclist(nIndex).Pos.Y = Y
                    Call CambiarOrigHeading(nIndex, MapData(map, X, Y).Trigger)
                    Call MakeNPCChar(True, map, nIndex, map, X, Y)
                    Exit Sub
                Else
                    altpos.X = RandomNumber(Areas(Area).X1, Areas(Area).X2)      'Obtenemos posicion al azar en x
                    altpos.Y = RandomNumber(Areas(Area).Y1, Areas(Area).Y2)    'Obtenemos posicion al azar en y
                    Call ClosestLegalPos(altpos, newpos)
                    If newpos.X <> 0 And newpos.Y <> 0 Then
                        Npclist(nIndex).Pos.map = newpos.map
                        Npclist(nIndex).Pos.X = newpos.X
                        Npclist(nIndex).Pos.Y = newpos.Y
                        Call CambiarOrigHeading(nIndex, MapData(newpos.map, newpos.X, newpos.Y).Trigger)
                        Call MakeNPCChar(True, newpos.map, nIndex, newpos.map, newpos.X, newpos.Y)
                        Exit Sub
                    Else
                        Call QuitarNPC(nIndex)
                        Call LogError(MAXSPAWNATTEMPS & " iteraciones en CrearNpc Area:" & Area & " NroNpc:" & NroNPC)
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        'asignamos las nuevas coordenas
        map = newpos.map
        X = Npclist(nIndex).Pos.X
        Y = Npclist(nIndex).Pos.Y
        
    End If
    CheckZonaNPC (nIndex)
    'Crea el NPC
    Call MakeNPCChar(True, map, nIndex, map, X, Y)

End Sub

Public Sub MakeNPCChar(ByVal toMap As Boolean, sndIndex As Integer, NpcIndex As Integer, ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer)
Dim CharIndex As Integer
Dim Nombre As String
Dim Criminal As Byte
    If Npclist(NpcIndex).Char.CharIndex = 0 Then
        CharIndex = NextOpenCharIndex
        Npclist(NpcIndex).Char.CharIndex = CharIndex
        CharList(CharIndex) = NpcIndex
    End If
    
    MapData(map, X, Y).NpcIndex = NpcIndex
    
    If Npclist(NpcIndex).MostrarNombre Then
        Nombre = "!" & Npclist(NpcIndex).Name
    Else
        Nombre = vbNullString
    End If
    Criminal = Npclist(NpcIndex).Stats.Alineacion
    
    If Not toMap Then
        Call WriteCharacterCreate(sndIndex, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.Heading, Npclist(NpcIndex).Char.CharIndex, X, Y, Npclist(NpcIndex).Char.WeaponAnim, Npclist(NpcIndex).Char.ShieldAnim, 0, 0, Npclist(NpcIndex).Char.CascoAnim, Nombre, Criminal, 0)
        Call FlushBuffer(sndIndex)
    Else
        Call AgregarNpc(NpcIndex)
    End If
End Sub

Public Sub ChangeNPCChar(ByVal NpcIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As eHeading)
    If NpcIndex > 0 Then
        With Npclist(NpcIndex).Char
            .Body = Body
            .Head = Head
            .Heading = Heading
            
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterChange(Body, Head, Heading, .CharIndex, 0, 0, 0, 0, 0))
        End With
    End If
End Sub

Private Sub EraseNPCChar(ByVal NpcIndex As Integer)

If Npclist(NpcIndex).Char.CharIndex <> 0 Then CharList(Npclist(NpcIndex).Char.CharIndex) = 0

If Npclist(NpcIndex).Char.CharIndex = LastChar Then
    Do Until CharList(LastChar) > 0
        LastChar = LastChar - 1
        If LastChar <= 1 Then Exit Do
    Loop
End If

'Quitamos del mapa
MapData(Npclist(NpcIndex).Pos.map, Npclist(NpcIndex).Pos.X, Npclist(NpcIndex).Pos.Y).NpcIndex = 0

'Actualizamos los clientes
Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterRemove(Npclist(NpcIndex).Char.CharIndex))

'Update la lista npc
Npclist(NpcIndex).Char.CharIndex = 0


'update NumChars
NumChars = NumChars - 1


End Sub

Public Sub MoveNPCChar(ByVal NpcIndex As Integer, ByVal nHeading As Byte)
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/04/2009
'06/04/2009: ZaMa - Now npcs can force to change position with dead character
'***************************************************

On Error GoTo errh
    Dim nPos As WorldPos
    Dim CasperPos As WorldPos
    Dim UserIndex As Integer
    Dim CasperHeading As Byte
    
    With Npclist(NpcIndex)
        nPos = .Pos
        Call HeadtoPos(nHeading, nPos)
        ' es una posicion legal
        If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, .flags.AguaValida = 1, .MaestroUser <> 0) Then
            
            If .flags.AguaValida = 0 And HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
            If .flags.TierraInvalida = 1 And Not HayAgua(.Pos.map, nPos.X, nPos.Y) Then Exit Sub
            

            UserIndex = MapData(.Pos.map, nPos.X, nPos.Y).UserIndex
            ' Si hay un usuario a dodne se mueve el npc, entonces esta muerto
            If UserIndex > 0 Then
                    CasperHeading = InvertHeading(nHeading)
                    CasperPos = UserList(UserIndex).Pos
                    MapData(.Pos.map, CasperPos.X, CasperPos.Y).UserIndex = 0
                    Call HeadtoPos(CasperHeading, CasperPos)
            

                    ' Avisamos a los usuarios del area, y al propio usuario lo forzamos a moverse
                    'Call SendData(SendTarget.ToPCAreaButIndex, UserIndex, PrepareMessageCharacterMove(UserList(UserIndex).Char.CharIndex, CasperPos.X, CasperPos.Y))
                    'Call WriteForceCharMove(UserIndex, CasperHeading)
                    
                    
                    'Update map and user pos
                    UserList(UserIndex).Pos = CasperPos
                    UserList(UserIndex).Char.Heading = CasperHeading
                    MapData(.Pos.map, CasperPos.X, CasperPos.Y).UserIndex = UserIndex
            


            End If
            Call SendData(SendTarget.ToNPCArea, NpcIndex, PrepareMessageCharacterMove(.Char.CharIndex, nPos.X, nPos.Y))

            'Update map and user pos
            MapData(.Pos.map, .Pos.X, .Pos.Y).NpcIndex = 0
            .Pos = nPos
            .Char.Heading = nHeading
            MapData(.Pos.map, nPos.X, nPos.Y).NpcIndex = NpcIndex
            
            'Actualizamos las áreas de ser necesario
            If UserIndex > 0 Then Call ModAreas.CheckUpdateNeededUser(UserIndex, CasperHeading)
            Call CheckUpdateNeededNpc(NpcIndex, nHeading)
        
        ElseIf .MaestroUser = 0 Then
            If .Movement = TipoAI.NpcPathfinding Then
                'Someone has blocked the npc's way, we must to seek a new path!
                .PFINFO.PathLenght = 0
            End If
        End If
        CheckZonaNPC (NpcIndex)
    End With
Exit Sub

errh:
    LogError ("Error en move npc " & NpcIndex)
End Sub

Function NextOpenNPC() As Integer
'Call LogTarea("Sub NextOpenNPC")

On Error GoTo errhandler
    Dim LoopC As Long
      
    For LoopC = 1 To MAXNPCS + 1
        If LoopC > MAXNPCS Then Exit For
        If Not Npclist(LoopC).flags.NPCActive Then Exit For
    Next LoopC
      
    NextOpenNPC = LoopC
Exit Function

errhandler:
    Call LogError("Error en NextOpenNPC")
End Function

Sub NpcEnvenenarUser(ByVal UserIndex As Integer)

Dim N As Integer
N = RandomNumber(1, 100)
If N < 30 Then
    UserList(UserIndex).flags.Envenenado = 1
    Call WriteConsoleMsg(UserIndex, "¡¡La criatura te ha envenenado!!", FontTypeNames.FONTTYPE_FIGHT)
End If

End Sub

Function SpawnNpc(ByVal NpcIndex As Integer, Pos As WorldPos, ByVal FX As Boolean, ByVal ReSpawn As Boolean, ByVal Zona As Integer) As Integer
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 06/15/2008
'23/01/2007 -> Pablo (ToxicWaste): Creates an NPC of the type Npcindex
'06/15/2008 -> Optimizé el codigo. (NicoNZ)
'***************************************************
Dim newpos As WorldPos
Dim altpos As WorldPos
Dim nIndex As Integer
Dim PosicionValida As Boolean
Dim PuedeAgua As Boolean
Dim PuedeTierra As Boolean


Dim map As Integer
Dim X As Integer
Dim Y As Integer

nIndex = OpenNPC(NpcIndex, ReSpawn)    'Conseguimos un indice

If nIndex > MAXNPCS Then
    SpawnNpc = 0
    Exit Function
End If

PuedeAgua = Npclist(nIndex).flags.AguaValida
PuedeTierra = Not Npclist(nIndex).flags.TierraInvalida = 1
        
Call ClosestLegalPos(Pos, newpos, PuedeAgua, PuedeTierra)  'Nos devuelve la posicion valida mas cercana
Call ClosestLegalPos(Pos, altpos, PuedeAgua)
'Si X e Y son iguales a 0 significa que no se encontro posicion valida

If newpos.X <> 0 And newpos.Y <> 0 Then
    'Asignamos las nuevas coordenas solo si son validas
    Npclist(nIndex).Pos.map = newpos.map
    Npclist(nIndex).Pos.X = newpos.X
    Npclist(nIndex).Pos.Y = newpos.Y
    Call CambiarOrigHeading(nIndex, MapData(newpos.map, newpos.X, newpos.Y).Trigger)
    If ReSpawn Then Npclist(nIndex).Orig = newpos
    PosicionValida = True
Else
    If altpos.X <> 0 And altpos.Y <> 0 Then
        Npclist(nIndex).Pos.map = altpos.map
        Npclist(nIndex).Pos.X = altpos.X
        Npclist(nIndex).Pos.Y = altpos.Y
        Call CambiarOrigHeading(nIndex, MapData(altpos.map, altpos.X, altpos.Y).Trigger)
        If ReSpawn Then Npclist(nIndex).Orig = altpos
        PosicionValida = True
    Else
        PosicionValida = False
    End If
End If

If Not PosicionValida Then
    Call QuitarNPC(nIndex)
    SpawnNpc = 0
    Exit Function
End If

'asignamos las nuevas coordenas
map = newpos.map
X = Npclist(nIndex).Pos.X
Y = Npclist(nIndex).Pos.Y

Npclist(nIndex).Zona = Zona

'Crea el NPC
Call MakeNPCChar(True, map, nIndex, map, X, Y)

Call CheckZonaNPC(nIndex)

If FX Then
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessagePlayWave(SND_WARP, X, Y))
    Call SendData(SendTarget.ToNPCArea, nIndex, PrepareMessageCreateFX(Npclist(nIndex).Char.CharIndex, FXIDs.FXWARP, 0))
End If

SpawnNpc = nIndex

End Function

Sub ReSpawnNpc(MiNPC As NPC)

If (MiNPC.flags.ReSpawn = 0) Then
    If MiNPC.NPCtype = Mercader Then
        Call ReSpawnMercader(MiNPC.Numero)
    ElseIf MiNPC.NPCtype = eNPCType.Fortaleza Then
        Call ReSpawnFortaleza(MiNPC.flags.Faccion, MiNPC.Pos.X < 550)
    Else
        Call CrearNPC(MiNPC.Numero, MiNPC.Area, MiNPC.Orig)
    End If
End If
End Sub

Private Sub NPCTirarOro(ByRef MiNPC As NPC)
'SI EL NPC TIENE ORO LO TIRAMOS
    If MiNPC.GiveGLDMin > 0 Then
        Dim MiObj As Obj
        Dim MiAux As Long
        MiAux = RandomNumber(MiNPC.GiveGLDMin, MiNPC.GiveGLDMax)
        Do While MiAux > MAX_INVENTORY_OBJS
            MiObj.Amount = MAX_INVENTORY_OBJS
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
            MiAux = MiAux - MAX_INVENTORY_OBJS
        Loop
        If MiAux > 0 Then
            MiObj.Amount = MiAux
            MiObj.ObjIndex = iORO
            Call TirarItemAlPiso(MiNPC.Pos, MiObj)
        End If
    End If
End Sub

Public Function OpenNPC(ByVal NpcNumber As Integer, Optional ByVal ReSpawn = True) As Integer

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'    ¡¡¡¡ NO USAR GetVar PARA LEER LOS NPCS !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'conmigo. Para leer los NPCS se deberá usar la
'nueva clase clsIniReader.
'
'Alejo
'
'###################################################
    Dim NpcIndex As Integer
    Dim Leer As clsIniReader
    Dim LoopC As Long
    Dim ln As String
    Dim aux As String
    
    Set Leer = LeerNPCs
    
    'If requested index is invalid, abort
    If Not Leer.KeyExists("NPC" & NpcNumber) Then
        OpenNPC = MAXNPCS + 1
        Exit Function
    End If
    
    NpcIndex = NextOpenNPC
    
    If NpcIndex > MAXNPCS Then 'Limite de npcs
        OpenNPC = NpcIndex
        Exit Function
    End If
    
    With Npclist(NpcIndex)
        .Numero = NpcNumber
        .Name = Leer.GetValue("NPC" & NpcNumber, "Name")
        .MostrarNombre = False
        If .Name <> "" Then
            If left(.Name, 1) = "!" Then
                .Name = Right(.Name, Len(.Name) - 1)
                .MostrarNombre = True
            End If
        End If
        
        .desc = Leer.GetValue("NPC" & NpcNumber, "Desc")
        
        .Movement = val(Leer.GetValue("NPC" & NpcNumber, "Movement"))
        .flags.OldMovement = .Movement
        
        .flags.AguaValida = val(Leer.GetValue("NPC" & NpcNumber, "AguaValida"))
        .flags.TierraInvalida = val(Leer.GetValue("NPC" & NpcNumber, "TierraInValida"))
        .flags.Faccion = val(Leer.GetValue("NPC" & NpcNumber, "Faccion"))
        .flags.AtacaDoble = val(Leer.GetValue("NPC" & NpcNumber, "AtacaDoble"))
        
        .NPCtype = val(Leer.GetValue("NPC" & NpcNumber, "NpcType"))
        
        .Char.Body = val(Leer.GetValue("NPC" & NpcNumber, "Body"))
        .Char.Head = val(Leer.GetValue("NPC" & NpcNumber, "Head"))
        If .Char.Head < 0 Then .Char.Head = DarCabeza(.Char.Head)
        .OrigBody = .Char.Body
        .OrigHead = .Char.Head
        
        .Char.WeaponAnim = val(Leer.GetValue("NPC" & NpcNumber, "Arma"))
        .Char.ShieldAnim = val(Leer.GetValue("NPC" & NpcNumber, "Escudo"))
        .Char.CascoAnim = val(Leer.GetValue("NPC" & NpcNumber, "Casco"))
        .OrigHeading = val(Leer.GetValue("NPC" & NpcNumber, "Heading"))
        .Char.Heading = .OrigHeading
        
        .Attackable = val(Leer.GetValue("NPC" & NpcNumber, "Attackable"))
        .Comercia = val(Leer.GetValue("NPC" & NpcNumber, "Comercia"))
        .Hostile = val(Leer.GetValue("NPC" & NpcNumber, "Hostile"))
        .flags.OldHostil = .Hostile
        
        .GiveEXP = val(Leer.GetValue("NPC" & NpcNumber, "GiveEXP")) * MultiplicadorEXP
        
        .flags.ExpCount = .GiveEXP
        
        .Veneno = val(Leer.GetValue("NPC" & NpcNumber, "Veneno"))
        
        .flags.Domable = val(Leer.GetValue("NPC" & NpcNumber, "Domable"))
        
        .GiveGLDMin = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLDMin")) * MultiplicadorORO
        .GiveGLDMax = val(Leer.GetValue("NPC" & NpcNumber, "GiveGLDMax")) * MultiplicadorORO
        
        .PoderAtaque = val(Leer.GetValue("NPC" & NpcNumber, "PoderAtaque"))
        .PoderEvasion = val(Leer.GetValue("NPC" & NpcNumber, "PoderEvasion"))
        
        .InvReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "InvReSpawn"))
        
        With .Stats
            .MaxHP = val(Leer.GetValue("NPC" & NpcNumber, "MaxHP"))
            .MinHP = val(Leer.GetValue("NPC" & NpcNumber, "MinHP"))
            .MaxHIT = val(Leer.GetValue("NPC" & NpcNumber, "MaxHIT"))
            .MinHIT = val(Leer.GetValue("NPC" & NpcNumber, "MinHIT"))
            .def = val(Leer.GetValue("NPC" & NpcNumber, "DEF"))
            .defM = val(Leer.GetValue("NPC" & NpcNumber, "DEFm"))
            .Alineacion = val(Leer.GetValue("NPC" & NpcNumber, "Alineacion"))
        End With
        
        .Invent.NroItems = val(Leer.GetValue("NPC" & NpcNumber, "NROITEMS"))
        For LoopC = 1 To .Invent.NroItems
            ln = Leer.GetValue("NPC" & NpcNumber, "Obj" & LoopC)
            .Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
            .Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
        Next LoopC
        
        .flags.LanzaSpells = val(Leer.GetValue("NPC" & NpcNumber, "LanzaSpells"))
        If .flags.LanzaSpells > 0 Then ReDim .Spells(1 To .flags.LanzaSpells)
        For LoopC = 1 To .flags.LanzaSpells
            .Spells(LoopC) = val(Leer.GetValue("NPC" & NpcNumber, "Sp" & LoopC))
        Next LoopC
        
        If .NPCtype = eNPCType.Entrenador Then
            .NroCriaturas = val(Leer.GetValue("NPC" & NpcNumber, "NroCriaturas"))
            ReDim .Criaturas(1 To .NroCriaturas) As tCriaturasEntrenador
            For LoopC = 1 To .NroCriaturas
                .Criaturas(LoopC).NpcIndex = Leer.GetValue("NPC" & NpcNumber, "CI" & LoopC)
                .Criaturas(LoopC).NpcName = Leer.GetValue("NPC" & NpcNumber, "CN" & LoopC)
            Next LoopC
        End If
        
        With .flags
            .NPCActive = True
            
            If ReSpawn Then
                .ReSpawn = val(Leer.GetValue("NPC" & NpcNumber, "ReSpawn"))
            Else
                .ReSpawn = 1
            End If
            
            .BackUp = val(Leer.GetValue("NPC" & NpcNumber, "BackUp"))
            .RespawnOrigPos = val(Leer.GetValue("NPC" & NpcNumber, "OrigPos"))
            .AfectaParalisis = val(Leer.GetValue("NPC" & NpcNumber, "AfectaParalisis"))
            
            .Snd1 = val(Leer.GetValue("NPC" & NpcNumber, "Snd1"))
            .Snd2 = val(Leer.GetValue("NPC" & NpcNumber, "Snd2"))
            .Snd3 = val(Leer.GetValue("NPC" & NpcNumber, "Snd3"))
        End With
        
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        .NroExpresiones = val(Leer.GetValue("NPC" & NpcNumber, "NROEXP"))
        If .NroExpresiones > 0 Then ReDim .Expresiones(1 To .NroExpresiones) As String
        For LoopC = 1 To .NroExpresiones
            .Expresiones(LoopC) = Leer.GetValue("NPC" & NpcNumber, "Exp" & LoopC)
        Next LoopC
        '<<<<<<<<<<<<<< Expresiones >>>>>>>>>>>>>>>>
        
        'Tipo de items con los que comercia
        .TipoItems = val(Leer.GetValue("NPC" & NpcNumber, "TipoItems"))
        
        .Ciudad = val(Leer.GetValue("NPC" & NpcNumber, "Ciudad"))
    End With
    
    'Update contadores de NPCs
    If NpcIndex > LastNPC Then LastNPC = NpcIndex
    NumNPCs = NumNPCs + 1
    
    'Devuelve el nuevo Indice
    OpenNPC = NpcIndex
End Function

Public Sub DoFollow(ByVal NpcIndex As Integer, ByVal UserName As String)
    With Npclist(NpcIndex)
        If .flags.Follow Then
            .flags.AttackedBy = vbNullString
            .flags.Follow = False
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
        Else
            .flags.AttackedBy = UserName
            .flags.Follow = True
            .Movement = TipoAI.NPCDEFENSA
            .Hostile = 0
        End If
    End With
End Sub

Public Sub FollowAmo(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        .flags.Follow = True
        .Movement = TipoAI.SigueAmo
        .Hostile = 0
        .Target = 0
        .TargetNPC = 0
    End With
End Sub

Attribute VB_Name = "PathFinding"
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

'#######################################################
'PathFinding Module
'Coded By Gulfas Morgolock
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'
'Ore is an excellent engine for introducing you not only
'to online game programming but also to general
'game programming. I am convinced that Aaron Perkings, creator
'of ORE, did a great work. He made possible that a lot of
'people enjoy for no fee games made with his engine, and
'for me, this is something great.
'
'I'd really like to contribute to this work, and all the
'projects of free ore-based MMORPGs that are on the net.
'
'I did some basic improvements on the AI of the NPCs, I
'added pathfinding, so now, the npcs are able to avoid
'obstacles. I believe that this improvement was essential
'for the engine.
'
'I'd like to see this as my contribution to ORE project,
'I hope that someone finds this source code useful.
'So, please feel free to do whatever you want with my
'pathfinging module.
'
'I'd really appreciate that if you find this source code
'useful you mention my nickname on the credits of your
'program. But there is no obligation ;).
'
'.........................................................
'Note:
'There is a little problem, ORE refers to map arrays in a
'different manner that my pathfinding routines. When I wrote
'these routines, I did it without thinking in ORE, so in my
'program I refer to maps in the usual way I do it.
'
'For example, suppose we have:
'Map(1 to Y,1 to X) as MapBlock
'I usually use the first coordinate as Y, and
'the second one as X.
'
'ORE refers to maps in converse way, for example:
'Map(1 to X,1 to Y) as MapBlock. As you can see the
'roles of first and second coordinates are different
'that my routines
'
'#######################################################


Option Explicit

Private Const ROWS As Integer = YMaxMapSize
Private Const COLUMS As Integer = XMaxMapSize
Private Const MAXINT As Integer = 10000

Private Type tIntermidiateWork
    Known As Boolean
    DistV As Integer
    PrevV As tVertice
End Type

Dim TmpArray(1 To COLUMS, 1 To ROWS) As tIntermidiateWork

Dim TilePosY As Integer

Private Function Limites(ByVal vfila As Integer, ByVal vcolu As Integer)
Limites = vcolu >= 1 And vcolu <= ROWS And vfila >= 1 And vfila <= COLUMS
End Function

Public Function IsWalkable(ByVal map As Integer, ByVal row As Integer, ByVal Col As Integer, ByVal NpcIndex As Integer) As Boolean
IsWalkable = MapData(map, row, Col).Blocked = 0 And Not HayAgua(map, row, Col)
With Npclist(NpcIndex)
If MapData(map, row, Col).UserIndex <> 0 Then
     If MapData(map, row, Col).UserIndex <> .PFINFO.TargetUser Then
        IsWalkable = False
        If .PFINFO.TargetUser = 0 And .PFINFO.TargetNPC = 0 Then
            'Un pequeño detalle ;)
            If .Mensaje = 0 And row = .Orig.X And Col = .Orig.Y Then
                Call SendData(ToNPCArea, NpcIndex, PrepareMessageChatOverHead("Disculpe señor, usted está parado en mi puesto de vigilancia.", Npclist(NpcIndex).Char.CharIndex, vbWhite))
                Npclist(NpcIndex).Mensaje = 1
            End If
        End If
    End If
ElseIf MapData(map, row, Col).NpcIndex <> 0 Then
    If MapData(map, row, Col).NpcIndex <> .PFINFO.TargetNPC Then IsWalkable = False
End If
End With
End Function

Private Sub ProcessAdjacents(ByVal MapIndex As Integer, ByRef T() As tIntermidiateWork, ByRef vfila As Integer, ByRef vcolu As Integer, ByVal NpcIndex As Integer)
    Dim v As tVertice
    Dim j As Integer
    'Look to North
    j = vfila - 1
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                    'Nos aseguramos que no hay un camino más corto
                    If T(j, vcolu).DistV = MAXINT Then
                        'Actualizamos la tabla de calculos intermedios
                        T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                        T(j, vcolu).PrevV.X = vcolu
                        T(j, vcolu).PrevV.Y = vfila
                        'Mete el vertice en la cola
                        v.X = vcolu
                        v.Y = j
                        Call Push(v)
                    End If
            End If
    End If
    j = vfila + 1
    'look to south
    If Limites(j, vcolu) Then
            If IsWalkable(MapIndex, j, vcolu, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(j, vcolu).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(j, vcolu).DistV = T(vfila, vcolu).DistV + 1
                    T(j, vcolu).PrevV.X = vcolu
                    T(j, vcolu).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    v.X = vcolu
                    v.Y = j
                    Call Push(v)
                End If
            End If
    End If
    'look to west
    If Limites(vfila, vcolu - 1) Then
            If IsWalkable(MapIndex, vfila, vcolu - 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(vfila, vcolu - 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(vfila, vcolu - 1).DistV = T(vfila, vcolu).DistV + 1
                    T(vfila, vcolu - 1).PrevV.X = vcolu
                    T(vfila, vcolu - 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    v.X = vcolu - 1
                    v.Y = vfila
                    Call Push(v)
                End If
            End If
    End If
    'look to east
    If Limites(vfila, vcolu + 1) Then
            If IsWalkable(MapIndex, vfila, vcolu + 1, NpcIndex) Then
                'Nos aseguramos que no hay un camino más corto
                If T(vfila, vcolu + 1).DistV = MAXINT Then
                    'Actualizamos la tabla de calculos intermedios
                    T(vfila, vcolu + 1).DistV = T(vfila, vcolu).DistV + 1
                    T(vfila, vcolu + 1).PrevV.X = vcolu
                    T(vfila, vcolu + 1).PrevV.Y = vfila
                    'Mete el vertice en la cola
                    v.X = vcolu + 1
                    v.Y = vfila
                    Call Push(v)
                End If
            End If
    End If
   
   
End Sub


Public Sub SeekPath(ByVal NpcIndex As Integer, Optional ByVal MaxSteps As Integer = 30)
'############################################################
'This Sub seeks a path from the npclist(npcindex).pos
'to the location NPCList(NpcIndex).PFINFO.Target.
'The optional parameter MaxSteps is the maximum of steps
'allowed for the path.
'############################################################

Dim cur_npc_pos As tVertice
Dim tar_npc_pos As tVertice
Dim v As tVertice
Dim NpcMap As Integer
Dim steps As Integer

NpcMap = Npclist(NpcIndex).Pos.map

steps = 0

cur_npc_pos.X = Npclist(NpcIndex).Pos.Y
cur_npc_pos.Y = Npclist(NpcIndex).Pos.X

tar_npc_pos.X = Npclist(NpcIndex).PFINFO.Target.X '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.X
tar_npc_pos.Y = Npclist(NpcIndex).PFINFO.Target.Y '  UserList(NPCList(NpcIndex).PFINFO.TargetUser).Pos.Y

Call InitializeTable(TmpArray, cur_npc_pos)
Call InitQueue

'We add the first vertex to the Queue
Call Push(cur_npc_pos)

Do While (Not IsEmpty)
    If steps > MaxSteps Then Exit Do
    v = Pop
    If v.X = tar_npc_pos.X And v.Y = tar_npc_pos.Y Then Exit Do
    Call ProcessAdjacents(NpcMap, TmpArray, v.Y, v.X, NpcIndex)
Loop

Call MakePath(NpcIndex)

End Sub

Private Sub MakePath(ByVal NpcIndex As Integer)
'#######################################################
'Builds the path previously calculated
'#######################################################

Dim Pasos As Integer
Dim miV As tVertice
Dim I As Integer

Pasos = TmpArray(Npclist(NpcIndex).PFINFO.Target.Y, Npclist(NpcIndex).PFINFO.Target.X).DistV
Npclist(NpcIndex).PFINFO.PathLenght = Pasos


If Pasos = MAXINT Then
    'MsgBox "There is no path."
    Npclist(NpcIndex).PFINFO.NoPath = True
    Npclist(NpcIndex).PFINFO.PathLenght = 0
    Exit Sub
End If

ReDim Npclist(NpcIndex).PFINFO.Path(0 To Pasos) As tVertice

miV.X = Npclist(NpcIndex).PFINFO.Target.X
miV.Y = Npclist(NpcIndex).PFINFO.Target.Y

For I = Pasos To 1 Step -1
    Npclist(NpcIndex).PFINFO.Path(I) = miV
    miV = TmpArray(miV.Y, miV.X).PrevV
Next I

Npclist(NpcIndex).PFINFO.curPos = 1
Npclist(NpcIndex).PFINFO.NoPath = False
   
End Sub

Private Sub InitializeTable(ByRef T() As tIntermidiateWork, ByRef S As tVertice, Optional ByVal MaxSteps As Integer = 30)
'#########################################################
'Initialize the array where we calculate the path
'#########################################################

Dim j As Integer, k As Integer
Const anymap = 1
For j = S.Y - MaxSteps To S.Y + MaxSteps
    For k = S.X - MaxSteps To S.X + MaxSteps
        If InMapBounds(anymap, j, k) Then
            T(j, k).Known = False
            T(j, k).DistV = MAXINT
            T(j, k).PrevV.X = 0
            T(j, k).PrevV.Y = 0
        End If
    Next
Next
T(S.Y, S.X).Known = False
T(S.Y, S.X).DistV = 0

End Sub


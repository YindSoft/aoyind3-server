Attribute VB_Name = "AI"
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

Option Explicit

Public Enum TipoAI
    ESTATICO = 1
    MueveAlAzar = 2
    NpcMaloAtacaUsersBuenos = 3
    NPCDEFENSA = 4
    GuardiasAtacanCriminales = 5
    SigueAmo = 8
    NpcAtacaNpc = 9
    NpcPathfinding = 10
    Personalizado = 11
End Enum

Public Const ELEMENTALFUEGO As Integer = 93
Public Const ELEMENTALTIERRA As Integer = 94
Public Const ELEMENTALAGUA As Integer = 92

'Damos a los NPCs el mismo rango de visión que un PJ
Public Const RANGO_VISION_X As Byte = 8
Public Const RANGO_VISION_Y As Byte = 6

'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'                        Modulo AI_NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'AI de los NPC
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿
'?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿

Private Function GuardiasAI(ByVal NpcIndex As Integer, ByVal Alineacion As Byte) As Boolean
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NI As Integer
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If (.flags.Inmovilizado = 0 And .flags.Paralizado = 0) Or headingloop = .Char.heading Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If (Alineacion = 0 And Criminal(UI)) Or (Alineacion = 1 And Not Criminal(UI) And Not EsNewbie(UI)) Or (Alineacion = 3 And Not FortalezaDelClan(UserList(UI).GuildIndex, NpcIndex)) Then
                                If .Char.Head <> headingloop Then
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                                End If
                                Call NpcAtacaUser(NpcIndex, UI)
                                GuardiasAI = True
                                Exit Function
                            ElseIf .flags.AttackedBy = UserList(UI).Name And Not .flags.Follow Then
                                If .Char.Head <> headingloop Then
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                                End If
                                Call NpcAtacaUser(NpcIndex, UI)
                                GuardiasAI = True
                                Exit Function
                            End If
                        End If
                    End If
                    NI = MapData(nPos.map, nPos.X, nPos.Y).NpcIndex
                    If NI > 0 Then
                        If ChaseNPC(NI, NpcIndex) Then
                            If .Char.Head <> headingloop Then
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                            End If
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                            GuardiasAI = True
                            Exit Function
                        End If
                    End If
                End If
            End If  'not inmovil
        Next headingloop
    End With
    GuardiasAI = False
    
    'Call RestoreOldMovement(NpcIndex)
End Function

''
' Handles the evil npcs' artificial intelligency.
'
' @param NpcIndex Specifies reference to the npc
Private Sub HostilMalvadoAI(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify Date: 28/04/2009
'28/04/2009: ZaMa - Now those NPCs who doble attack, have 50% of posibility of casting a spell on user.
'**************************************************************
    Dim nPos As WorldPos
    Dim headingloop As Byte
    Dim UI As Integer
    Dim NPCI As Integer
    Dim atacoPJ As Boolean
    
    atacoPJ = False
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If (.flags.Inmovilizado = 0 And .flags.Paralizado = 0) Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
                    NPCI = MapData(nPos.map, nPos.X, nPos.Y).NpcIndex
                    If UI > 0 And Not atacoPJ Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible And .flags.Paralizado = 0 Then
                            atacoPJ = True
                            If .flags.LanzaSpells Then
                                If .flags.AtacaDoble Then
                                    If (RandomNumber(0, 1)) Then
                                        If NpcAtacaUser(NpcIndex, UI) Then
                                            Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                                        End If
                                        Exit Sub
                                    End If
                                End If
                                
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            If NpcAtacaUser(NpcIndex, UI) Then
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                            End If
                            Exit Sub
                        End If
                    ElseIf NPCI > 0 Then
                        If Npclist(NPCI).MaestroUser > 0 Then
                            Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                            Call SistemaCombate.NpcAtacaNpc(NpcIndex, NPCI, False)
                            Exit Sub
                        End If
                    End If
                End If
            End If  'inmo
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub HostilBuenoAI(ByVal NpcIndex As Integer)
    Dim nPos As WorldPos
    Dim headingloop As eHeading
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        For headingloop = eHeading.NORTH To eHeading.WEST
            nPos = .Pos
            If (.flags.Inmovilizado = 0 And .flags.Paralizado = 0) Or .Char.heading = headingloop Then
                Call HeadtoPos(headingloop, nPos)
                If InMapBounds(nPos.map, nPos.X, nPos.Y) Then
                    UI = MapData(nPos.map, nPos.X, nPos.Y).UserIndex
                    If UI > 0 Then
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                
                                If NpcAtacaUser(NpcIndex, UI) Then
                                    Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, headingloop)
                                End If
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        Next headingloop
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub IrUsuarioCercano(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim i As Long
    Dim Elemento
    
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Then Exit Sub
        If .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If UserList(UI).flags.Muerto = 0 Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            Exit Sub
                        End If
                        
                    End If
                End If
            Next Elemento
        Else
            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                            If .flags.LanzaSpells <> 0 Then Call NpcLanzaUnSpell(NpcIndex, UI)
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                        
                    End If
                End If
            Next Elemento
            
            'Si llega aca es que no había ningún usuario cercano vivo.
            'A bailar. Pablo (ToxicWaste)
            If RandomNumber(0, 10) = 0 Then
                Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

''
' Makes a Pet / Summoned Npc to Follow an enemy
'
' @param NpcIndex Specifies reference to the npc
Private Sub SeguirAgresor(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Unknown
'Last Modify by: Marco Vanotti (MarKoxX)
'Last Modify Date: 08/16/2008
'08/16/2008: MarKoxX - Now pets that do melé attacks have to be near the enemy to attack.
'**************************************************************
    Dim tHeading As Byte
    Dim UI As Integer
    
    Dim i As Long
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim Elemento
    With Npclist(NpcIndex)
        If .flags.Paralizado = 1 Or .flags.Inmovilizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select

            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then

                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not Criminal(.MaestroUser) And Not Criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Exit Sub
                                End If
                            End If

                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NpcIndex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, UI)
                                        End If
                                    End If
                                 End If
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next Elemento
        Else
            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If UserList(UI).Name = .flags.AttackedBy Then
                            If .MaestroUser > 0 Then
                                If Not Criminal(.MaestroUser) And Not Criminal(UI) And (UserList(.MaestroUser).flags.Seguro Or UserList(.MaestroUser).Faccion.ArmadaReal = 1) Then
                                    Call WriteConsoleMsg(.MaestroUser, "La mascota no atacará a ciudadanos si eres miembro de la Armada Real o tienes el seguro activado", FontTypeNames.FONTTYPE_INFO)
                                    Call FlushBuffer(.MaestroUser)
                                    .flags.AttackedBy = vbNullString
                                    Call FollowAmo(NpcIndex)
                                    Exit Sub
                                End If
                            End If
                            
                            If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                                 If .flags.LanzaSpells > 0 Then
                                        Call NpcLanzaUnSpell(NpcIndex, UI)
                                 Else
                                    If Distancia(UserList(UI).Pos, Npclist(NpcIndex).Pos) <= 1 Then
                                        ' TODO : Set this a separate AI for Elementals and Druid's pets
                                        If Npclist(NpcIndex).Numero <> 92 Then
                                            Call NpcAtacaUser(NpcIndex, UI)
                                        End If
                                    End If
                                 End If
                                 
                                 tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 
                                 Exit Sub
                            End If
                        End If
                        
                    End If
                End If
                
            Next Elemento
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub RestoreOldMovement(ByVal NpcIndex As Integer)
    With Npclist(NpcIndex)
        If .MaestroUser = 0 Then
            .Movement = .flags.OldMovement
            .Hostile = .flags.OldHostil
            .flags.AttackedBy = vbNullString
        End If
    End With
End Sub

Private Sub PersigueCiudadano(ByVal NpcIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim Elemento
    With Npclist(NpcIndex)
        For Each Elemento In .AreasInfo.Users.Items
            UI = Elemento
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    If Not Criminal(UI) Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                    
               End If
            End If
            
        Next Elemento
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub


Private Sub PersigueClan(ByVal NpcIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim Elemento
    With Npclist(NpcIndex)
        For Each Elemento In .AreasInfo.Users.Items
            UI = Elemento
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    
                    If Not FortalezaDelClan(UserList(UI).GuildIndex, NpcIndex) Then
                        If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 Then
                            If .flags.LanzaSpells > 0 Then
                                Call NpcLanzaUnSpell(NpcIndex, UI)
                            End If
                            tHeading = FindDirection(.Pos, UserList(UI).Pos)
                            Call MoveNPCChar(NpcIndex, tHeading)
                            Exit Sub
                        End If
                    End If
                    
               End If
            End If
            
        Next Elemento
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub


Private Sub PersigueCriminal(ByVal NpcIndex As Integer)
    Dim UI As Integer
    Dim tHeading As Byte
    Dim i As Long
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    Dim Elemento
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X And Sgn(UserList(UI).Pos.X - .Pos.X) = SignoEO Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y And Sgn(UserList(UI).Pos.Y - .Pos.Y) = SignoNS Then
                        
                        If Criminal(UI) Then
                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                      Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                Exit Sub
                           End If
                        End If
                        
                   End If
                End If
                    
            Next Elemento
        Else
            For Each Elemento In .AreasInfo.Users.Items
                UI = Elemento
                'Is it in it's range of vision??
                If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                    If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                        If Criminal(UI) Then
                           If UserList(UI).flags.Muerto = 0 And UserList(UI).flags.invisible = 0 And UserList(UI).flags.Oculto = 0 And UserList(UI).flags.AdminPerseguible Then
                                If .flags.LanzaSpells > 0 Then
                                    Call NpcLanzaUnSpell(NpcIndex, UI)
                                End If
                                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                                tHeading = FindDirection(.Pos, UserList(UI).Pos)
                                Call MoveNPCChar(NpcIndex, tHeading)
                                Exit Sub
                           End If
                        End If
                        
                   End If
                End If
                
            Next Elemento
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub SeguirAmo(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim UI As Integer
    
    With Npclist(NpcIndex)
        If .Target = 0 And .TargetNPC = 0 Then
            UI = .MaestroUser
            
            'Is it in it's range of vision??
            If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
                If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                    If UserList(UI).flags.Muerto = 0 _
                            And UserList(UI).flags.invisible = 0 _
                            And UserList(UI).flags.Oculto = 0 _
                            And Distancia(.Pos, UserList(UI).Pos) > 3 Then
                        tHeading = FindDirection(.Pos, UserList(UI).Pos)
                        Call MoveNPCChar(NpcIndex, tHeading)
                        Exit Sub
                    End If
                End If
            End If
        End If
    End With
    
    Call RestoreOldMovement(NpcIndex)
End Sub

Private Sub AiNpcAtacaNpc(ByVal NpcIndex As Integer)
    Dim tHeading As Byte
    Dim X As Long
    Dim Y As Long
    Dim NI As Integer
    Dim bNoEsta As Boolean
    
    Dim SignoNS As Integer
    Dim SignoEO As Integer
    
    With Npclist(NpcIndex)
        If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
            Select Case .Char.heading
                Case eHeading.NORTH
                    SignoNS = -1
                    SignoEO = 0
                
                Case eHeading.EAST
                    SignoNS = 0
                    SignoEO = 1
                
                Case eHeading.SOUTH
                    SignoNS = 1
                    SignoEO = 0
                
                Case eHeading.WEST
                    SignoEO = -1
                    SignoNS = 0
            End Select
            
            For Y = .Pos.Y To .Pos.Y + SignoNS * RANGO_VISION_Y Step IIf(SignoNS = 0, 1, SignoNS)
                For X = .Pos.X To .Pos.X + SignoEO * RANGO_VISION_X Step IIf(SignoEO = 0, 1, SignoEO)
                    If X >= 1 And X <= XMaxMapSize And Y >= 1 And Y <= YMaxMapSize Then
                        NI = MapData(.Pos.map, X, Y).NpcIndex
                        If NI > 0 Then
                            If .TargetNPC = NI Then
                                bNoEsta = True
                                If .Numero = ELEMENTALFUEGO Then
                                    Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                    If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                 End If
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        Else
            For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y
                For X = .Pos.X - RANGO_VISION_Y To .Pos.X + RANGO_VISION_Y
                    If X >= 1 And X <= XMaxMapSize And Y >= 1 And Y <= YMaxMapSize Then
                       NI = MapData(.Pos.map, X, Y).NpcIndex
                       If NI > 0 Then
                            If .TargetNPC = NI Then
                                 bNoEsta = True
                                 If .Numero = ELEMENTALFUEGO Then
                                     Call NpcLanzaUnSpellSobreNpc(NpcIndex, NI)
                                     If Npclist(NI).NPCtype = DRAGON Then
                                        Npclist(NI).CanAttack = 1
                                        Call NpcLanzaUnSpellSobreNpc(NI, NpcIndex)
                                     End If
                                 Else
                                    'aca verificamosss la distancia de ataque
                                    If Distancia(.Pos, Npclist(NI).Pos) <= 1 Then
                                        Call SistemaCombate.NpcAtacaNpc(NpcIndex, NI)
                                    End If
                                 End If
                                 If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                                 If .TargetNPC = 0 Then Exit Sub
                                 tHeading = FindDirection(.Pos, Npclist(MapData(.Pos.map, X, Y).NpcIndex).Pos)
                                 Call MoveNPCChar(NpcIndex, tHeading)
                                 Exit Sub
                            End If
                       End If
                    End If
                Next X
            Next Y
        End If
        
        If Not bNoEsta Then
            If .MaestroUser > 0 Then
                Call FollowAmo(NpcIndex)
            Else
                .Movement = .flags.OldMovement
                .Hostile = .flags.OldHostil
            End If
        End If
    End With
End Sub

Sub NPCAI(ByVal NpcIndex As Integer)
On Error GoTo ErrorHandler
Dim Ataco As Boolean
    With Npclist(NpcIndex)
        '<<<<<<<<<<< Ataques >>>>>>>>>>>>>>>>
        If .MaestroUser = 0 Then
            'Busca a alguien para atacar
            '¿Es un guardia?
            If .NPCtype = eNPCType.Guardia Then
                Ataco = GuardiasAI(NpcIndex, .Stats.Alineacion)
            ElseIf .Hostile And .Stats.Alineacion <> 0 Then
                Call HostilMalvadoAI(NpcIndex)
            ElseIf .Hostile And .Stats.Alineacion = 0 Then
                Call HostilBuenoAI(NpcIndex)
            End If
        Else
            'Evitamos que ataque a su amo, a menos
            'que el amo lo ataque.
            'Call HostilBuenoAI(NpcIndex)
        End If
        
        
        '<<<<<<<<<<<Movimiento>>>>>>>>>>>>>>>>
        Select Case .Movement
            Case TipoAI.MueveAlAzar
                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
                If .NPCtype = eNPCType.Guardia Then
                    If EsGuardiaReal(NpcIndex) Then
                        Call PersigueCriminal(NpcIndex)
                    ElseIf EsGuardiaCaos(NpcIndex) Then
                        Call PersigueCiudadano(NpcIndex)
                    ElseIf EsGuardiaClan(NpcIndex) Then
                        Call PersigueClan(NpcIndex)
                    Else
                        Call SeguirAgresor(NpcIndex)
                    End If
                End If
            
            'Va hacia el usuario cercano
            Case TipoAI.NpcMaloAtacaUsersBuenos
                If .flags.Paralizado = 1 Then Exit Sub
                Call IrUsuarioCercano(NpcIndex)
            
            'Va hacia el usuario que lo ataco(FOLLOW)
            Case TipoAI.NPCDEFENSA
                If .flags.Paralizado = 1 Then Exit Sub
                Call SeguirAgresor(NpcIndex)
            
            'Persigue criminales
            Case TipoAI.GuardiasAtacanCriminales
                If .flags.Paralizado = 1 Then Exit Sub
                Call PersigueCriminal(NpcIndex)
            
            Case TipoAI.SigueAmo
                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                Call SeguirAmo(NpcIndex)
                If RandomNumber(1, 12) = 3 Then
                    Call MoveNPCChar(NpcIndex, CByte(RandomNumber(eHeading.NORTH, eHeading.WEST)))
                End If
            
            Case TipoAI.NpcAtacaNpc
                If .flags.Paralizado = 1 Then Exit Sub
                Call AiNpcAtacaNpc(NpcIndex)
            
            Case TipoAI.NpcPathfinding
                If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then Exit Sub
                If ReCalculatePath(NpcIndex) Then
                    Call PathFindingAI(NpcIndex)
                    'Existe el camino?
                    If .PFINFO.NoPath Then 'Si no existe nos movemos al azar
                        'Move randomly
                        If RandomNumber(1, 4) = 1 Then Call MoveNPCChar(NpcIndex, RandomNumber(eHeading.NORTH, eHeading.WEST))
                    Else
                        .Mensaje = 0
                    End If
                End If
                If Not .PFINFO.NoPath Then
                    If Not PathEnd(NpcIndex) Then
                        Call FollowPath(NpcIndex)
                    Else
                        .PFINFO.PathLenght = 0
                        If Not PathFindingAI(NpcIndex) Then
                            If .Char.heading <> .OrigHeading And Not Ataco Then
                                Call ChangeNPCChar(NpcIndex, .Char.Body, .Char.Head, .OrigHeading)
                            End If
                        End If
                    End If
                End If
            Case TipoAI.Personalizado
                Select Case .NPCtype
                    Case eNPCType.Mercader
                        'Call MoverMercader(NpcIndex)
                End Select
        End Select
    End With
Exit Sub

ErrorHandler:
    Call LogError("NPCAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As NPC
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)
End Sub

Function EsGuardiaReal(NpcIndex) As Boolean
EsGuardiaReal = (Npclist(NpcIndex).NPCtype = eNPCType.Guardia Or Npclist(NpcIndex).NPCtype = eNPCType.Mercader) And Npclist(NpcIndex).Stats.Alineacion = 0
End Function
Function EsGuardiaCaos(NpcIndex) As Boolean
EsGuardiaCaos = (Npclist(NpcIndex).NPCtype = eNPCType.Guardia Or Npclist(NpcIndex).NPCtype = eNPCType.Mercader) And Npclist(NpcIndex).Stats.Alineacion = 1
End Function
Function EsGuardiaNeutral(NpcIndex) As Boolean
EsGuardiaNeutral = Npclist(NpcIndex).NPCtype = eNPCType.Guardia And Npclist(NpcIndex).Stats.Alineacion = 2
End Function
Function EsGuardiaClan(NpcIndex) As Boolean
EsGuardiaClan = (Npclist(NpcIndex).NPCtype = eNPCType.Guardia Or Npclist(NpcIndex).NPCtype = eNPCType.Fortaleza) And Npclist(NpcIndex).Stats.Alineacion = 3
End Function

Function UserNear(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns True if there is an user adjacent to the npc position.
'#################################################################
If Npclist(NpcIndex).PFINFO.TargetUser > 0 Then
    UserNear = Abs(Npclist(NpcIndex).Pos.X - UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.X) <= RANGO_VISION_X + 3 And Abs(Npclist(NpcIndex).Pos.Y - UserList(Npclist(NpcIndex).PFINFO.TargetUser).Pos.Y) <= RANGO_VISION_Y + 3 And UserList(Npclist(NpcIndex).PFINFO.TargetUser).flags.Muerto = 0
ElseIf Npclist(NpcIndex).PFINFO.TargetNPC > 0 Then
    UserNear = Abs(Npclist(NpcIndex).Pos.X - Npclist(Npclist(NpcIndex).PFINFO.TargetNPC).Pos.X) <= RANGO_VISION_X + 3 And Abs(Npclist(NpcIndex).Pos.Y - Npclist(Npclist(NpcIndex).PFINFO.TargetNPC).Pos.Y) <= RANGO_VISION_Y + 3
Else
    UserNear = False
End If
End Function
Function TargetMal(ByVal NpcIndex As Integer) As Boolean
Dim tmpInt As Integer
'Sirve para saber si el target esta mas adelante o mas atraz, si es asi hay que recalcular
If Npclist(NpcIndex).PFINFO.TargetUser > 0 Then
    tmpInt = Npclist(NpcIndex).PFINFO.TargetUser
    TargetMal = (UserList(tmpInt).Pos.X > Npclist(NpcIndex).Pos.X And UserList(tmpInt).Pos.X < Npclist(NpcIndex).PFINFO.Target.Y) Or _
                (UserList(tmpInt).Pos.X < Npclist(NpcIndex).Pos.X And UserList(tmpInt).Pos.X > Npclist(NpcIndex).PFINFO.Target.Y) Or _
                (UserList(tmpInt).Pos.Y > Npclist(NpcIndex).Pos.X And UserList(tmpInt).Pos.Y < Npclist(NpcIndex).PFINFO.Target.X) Or _
                (UserList(tmpInt).Pos.Y < Npclist(NpcIndex).Pos.X And UserList(tmpInt).Pos.Y > Npclist(NpcIndex).PFINFO.Target.X)
    
    
ElseIf Npclist(NpcIndex).PFINFO.TargetNPC > 0 Then
    tmpInt = Npclist(NpcIndex).PFINFO.TargetNPC
    TargetMal = (Npclist(tmpInt).Pos.X > Npclist(NpcIndex).Pos.X And Npclist(tmpInt).Pos.X < Npclist(NpcIndex).PFINFO.Target.Y) Or _
                (Npclist(tmpInt).Pos.X < Npclist(NpcIndex).Pos.X And Npclist(tmpInt).Pos.X > Npclist(NpcIndex).PFINFO.Target.Y) Or _
                (Npclist(tmpInt).Pos.Y > Npclist(NpcIndex).Pos.X And Npclist(tmpInt).Pos.Y < Npclist(NpcIndex).PFINFO.Target.X) Or _
                (Npclist(tmpInt).Pos.Y < Npclist(NpcIndex).Pos.X And Npclist(tmpInt).Pos.Y > Npclist(NpcIndex).PFINFO.Target.X)
    
Else
    TargetMal = False
End If
End Function
Function ReCalculatePath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Returns true if we have to seek a new path
'#################################################################
    
    If Npclist(NpcIndex).PFINFO.NoPath Then
        ReCalculatePath = True
    ElseIf Npclist(NpcIndex).PFINFO.PathLenght = 0 Then
        ReCalculatePath = True
    ElseIf Not UserNear(NpcIndex) Then
        ReCalculatePath = True
    ElseIf TargetMal(NpcIndex) Then
        ReCalculatePath = True
    End If
End Function

Function PathEnd(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Returns if the npc has arrived to the end of its path
'#################################################################
If Npclist(NpcIndex).PFINFO.TargetUser > 0 Or Npclist(NpcIndex).PFINFO.TargetNPC > 0 Then
    'Si tiene un user tiene que ir al lado
    PathEnd = Npclist(NpcIndex).PFINFO.curPos = Npclist(NpcIndex).PFINFO.PathLenght
Else
    If Npclist(NpcIndex).PFINFO.PathLenght = 0 Or Npclist(NpcIndex).PFINFO.curPos > Npclist(NpcIndex).PFINFO.PathLenght Then
        'Npclist(NpcIndex).Mensaje = 0
        PathEnd = True
    Else
        PathEnd = False
    End If
End If
End Function

Function FollowPath(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock
'Moves the npc.
'#################################################################
    Dim tmpPos As WorldPos
    Dim tHeading As Byte
    
    tmpPos.map = Npclist(NpcIndex).Pos.map
    tmpPos.X = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.curPos).Y ' invertí las coordenadas
    tmpPos.Y = Npclist(NpcIndex).PFINFO.Path(Npclist(NpcIndex).PFINFO.curPos).X
    
    'Debug.Print "(" & tmpPos.X & "," & tmpPos.Y & ")"
    
    tHeading = FindDirection(Npclist(NpcIndex).Pos, tmpPos)
    
    MoveNPCChar NpcIndex, tHeading
    
    Npclist(NpcIndex).PFINFO.curPos = Npclist(NpcIndex).PFINFO.curPos + 1
    If Npclist(NpcIndex).PFINFO.TargetUser > 0 Then
        If Npclist(NpcIndex).flags.LanzaSpells > 0 And RandomNumber(1, 10) = 1 Then
            Call NpcLanzaUnSpell(NpcIndex, Npclist(NpcIndex).PFINFO.TargetUser)
        ElseIf RandomNumber(1, 10) = 1 Then
            'Mensajito
            Call SendData(ToNPCArea, NpcIndex, PrepareMessageChatOverHead("¡Ven aquí maldito bastardo!", Npclist(NpcIndex).Char.CharIndex, vbWhite))
        End If
    ElseIf Npclist(NpcIndex).PFINFO.TargetNPC > 0 Then
        If Npclist(NpcIndex).flags.LanzaSpells > 0 And RandomNumber(1, 5) = 1 Then
            Call NpcLanzaUnSpellSobreNpc(NpcIndex, Npclist(NpcIndex).PFINFO.TargetNPC)
        End If
    End If
End Function

Function PathFindingAI(ByVal NpcIndex As Integer) As Boolean
'#################################################################
'Coded By Gulfas Morgolock / 11-07-02
'www.geocities.com/gmorgolock
'morgolock@speedy.com.ar
'This function seeks the shortest path from the Npc
'to the user's location.
'#################################################################
Dim Y As Integer
Dim X As Integer
Dim NI As Integer
Dim UI As Integer
Dim Elemento
Dim Primero As Boolean
On Error GoTo ErrorHandler
With Npclist(NpcIndex)
    For Each Elemento In .AreasInfo.Users.Items
        UI = Elemento
                'Is it in it's range of vision??
        If Abs(UserList(UI).Pos.X - .Pos.X) <= RANGO_VISION_X Then
            If Abs(UserList(UI).Pos.Y - .Pos.Y) <= RANGO_VISION_Y Then
                        
                If ChaseUser(UI, NpcIndex) Then
                        .PFINFO.Target.X = UserList(UI).Pos.Y
                        .PFINFO.Target.Y = UserList(UI).Pos.X 'ops!
                        .PFINFO.TargetUser = UI
                        Call SeekPath(NpcIndex)
                        PathFindingAI = True
                        Exit Function
                End If
                        
            End If
        End If
    Next Elemento
    

    'Si hay algun usuario cerca miramos si hay NPC hostiles.
    For Y = .Pos.Y - RANGO_VISION_Y To .Pos.Y + RANGO_VISION_Y    'Makes a loop that looks at
         For X = .Pos.X - RANGO_VISION_X To .Pos.X + RANGO_VISION_X   '5 tiles in every direction
            
             'Make sure tile is legal
             If X > 1 And X < XMaxMapSize And Y > 1 And Y < YMaxMapSize Then
                
                 'look for a npc
                 NI = MapData(.Pos.map, X, Y).NpcIndex
                 If NI > 0 Then
                    If ChaseNPC(NI, NpcIndex) Then
                            .PFINFO.Target.X = Npclist(NI).Pos.Y
                            .PFINFO.Target.Y = Npclist(NI).Pos.X 'ops!
                            .PFINFO.TargetNPC = NI
                            Call SeekPath(NpcIndex)
                            PathFindingAI = True
                            Exit Function
                    End If
                End If
            End If
        Next X
    Next Y

    Call VolverOrigPos(NpcIndex)
End With
Exit Function
ErrorHandler:
    Call LogError("PathFindingAI " & Npclist(NpcIndex).Name & " " & Npclist(NpcIndex).MaestroUser & " " & Npclist(NpcIndex).MaestroNpc & " mapa:" & Npclist(NpcIndex).Pos.map & " x:" & Npclist(NpcIndex).Pos.X & " y:" & Npclist(NpcIndex).Pos.Y & " Mov:" & Npclist(NpcIndex).Movement & " TargU:" & Npclist(NpcIndex).Target & " TargN:" & Npclist(NpcIndex).TargetNPC)
    Dim MiNPC As NPC
    MiNPC = Npclist(NpcIndex)
    Call QuitarNPC(NpcIndex)
    Call ReSpawnNpc(MiNPC)

End Function
Sub VolverOrigPos(ByVal NpcIndex As Integer)
With Npclist(NpcIndex)
    'El Yind / Quiero que si no encontro a nadie vuelva para su cucha.
    If .Orig.map > 0 Then
        If .Pos.X <> .Orig.X Or .Pos.Y <> .Orig.Y Then
            If .PFINFO.Target.X <> .Orig.Y Or .PFINFO.Target.Y <> .Orig.X Or .PFINFO.PathLenght = 0 Then
                .PFINFO.Target.X = .Orig.Y
                .PFINFO.Target.Y = .Orig.X 'ops!
                .PFINFO.TargetUser = 0
                .PFINFO.TargetNPC = 0
                Call SeekPath(NpcIndex, 50)
            End If
        End If
    End If
End With
End Sub
Function ChaseUser(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal VerInvi As Boolean = False) As Boolean
'[El Yind]
'Funcion para saber si un NPC tiene que perseguir determinado usuario...
ChaseUser = False
If UserIndex = 0 Then
    ChaseUser = False
ElseIf UserList(UserIndex).flags.Muerto = 1 Or UserList(UserIndex).flags.AdminPerseguible = False Or (Not VerInvi And (UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1)) Then
    ChaseUser = False
Else
    With Npclist(NpcIndex)
        If .Hostile And .Stats.Alineacion <> 0 Then
            ChaseUser = True
        ElseIf .Hostile And .Stats.Alineacion = 0 Then
            If UserList(UserIndex).Name = .flags.AttackedBy Then
                ChaseUser = True
            End If
        ElseIf EsGuardiaReal(NpcIndex) Then
            If Criminal(UserIndex) And Zonas(UserList(UserIndex).Zona).Segura = 1 Then
                ChaseUser = True
            End If
        ElseIf EsGuardiaCaos(NpcIndex) Then
            If ((Not Criminal(UserIndex) And Not EsNewbie(UserIndex)) Or UserList(UserIndex).Name = .flags.AttackedBy) And Zonas(UserList(UserIndex).Zona).Segura = 1 Then
                ChaseUser = True
            End If
        ElseIf EsGuardiaNeutral(NpcIndex) Then
            If UserList(UserIndex).Name = .flags.AttackedBy And Zonas(UserList(UserIndex).Zona).Segura = 1 Then
                ChaseUser = True
            End If
        ElseIf EsGuardiaClan(NpcIndex) Then
            If Not FortalezaDelClan(UserList(UserIndex).GuildIndex, NpcIndex) And Zonas(UserList(UserIndex).Zona).Terreno = eTerreno.Fortaleza Then
                ChaseUser = True
            End If
        End If
    End With
End If
End Function
Function ChaseNPC(ByVal VictimaIndex As Integer, ByVal AtacanteIndex As Integer) As Boolean
'[El Yind]
'Funcion para saber si un NPC tiene que perseguir determinado npc...
If Npclist(VictimaIndex).Hostile And Npclist(VictimaIndex).Stats.Alineacion <> 0 And Npclist(VictimaIndex).Zona > 0 Then
    'Si el bicho se metio en la ciudad lo perseguimos :)
    If Zonas(Npclist(VictimaIndex).Zona).Segura = 1 Then
        ChaseNPC = True
    End If
Else
    ChaseNPC = False
End If
End Function
Sub NpcLanzaUnSpell(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
    If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then Exit Sub
    
    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreUser(NpcIndex, UserIndex, Npclist(NpcIndex).Spells(k))
End Sub

Sub NpcLanzaUnSpellSobreNpc(ByVal NpcIndex As Integer, ByVal TargetNPC As Integer)
    Dim k As Integer
    k = RandomNumber(1, Npclist(NpcIndex).flags.LanzaSpells)
    Call NpcLanzaSpellSobreNpc(NpcIndex, TargetNPC, Npclist(NpcIndex).Spells(k))
End Sub

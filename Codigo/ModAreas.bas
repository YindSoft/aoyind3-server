Attribute VB_Name = "ModAreas"
'Modulo rehecho para posibilitar las dimensiones del mapa

'El sistema consiste en tener siempre presente los usuarios y npc que el usuario ve, al mover, verifica si alguno de esos no esta mas en su rango
'si es asi actualiza, sino no hace nada, lo mismo pasa con los objetos, de esta forma es sistema es muy efectivo al moverse en cualquier lugar
'donde no haya demasiados usuarios aglomerados, que al ser en general el caso es conveniente.

Option Explicit

'>>>>>>AREAS>>>>>AREAS>>>>>>>>AREAS>>>>>>>AREAS>>>>>>>>>>




Public Type AreaInfo
    Users As New Dictionary
    NPCs As New Dictionary
    Barco(0 To 1) As Byte
End Type


Public Const USER_NUEVO As Byte = 255

Public Const MargenX As Integer = 12
Public Const MargenY As Integer = 10

'Cuidado:
' ¡¡¡LAS AREAS ESTÁN HARDCODEADAS!!!
Private CurDay As Byte
Private CurHour As Byte


Public Sub CheckUpdateNeededUser(ByVal UserIndex As Integer, ByVal Head As Byte)
'**************************************************************
'Author: Javier Podavini (El Yind)
'Last Modify Date: 31/10/2011
'Es la función clave del sistema de areas... Es llamada al mover un user
'**************************************************************
    Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer
    Dim BMinX As Integer, BMaxX As Integer, BMinY As Integer, BMaxY As Integer
    Dim TempInt As Integer, map As Integer
    Dim Elemento
    Dim QuitoAlgo As Boolean
    Dim Barco As clsBarco
On Error GoTo errh:

    With UserList(UserIndex)
    
        '.AreasInfo.Pasos = .AreasInfo.Pasos + 1
        'Chequeo que haya hecho 3 pasos para no estar comprobando todo el tiempo
        'If .AreasInfo.Pasos = 3 Or Head = USER_NUEVO Then
        '.AreasInfo.Pasos = 0
        
        MinX = .Pos.X
        MaxX = MinX
        MinY = .Pos.Y
        MaxY = MinY

        If Head = eHeading.NORTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MinY
            
            BMinX = MinX
            BMaxX = MaxX
            BMinY = MinY + MargenY * 2
            BMaxY = BMinY
            
        ElseIf Head = eHeading.SOUTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY + MargenY
            MaxY = MinY
            
            BMinX = MinX
            BMaxX = MaxX
            BMinY = MinY - MargenY * 2
            BMaxY = BMinY
        
        ElseIf Head = eHeading.WEST Then
            MinX = MinX - MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
            BMinX = MinX + MargenX * 2
            BMaxX = BMinX
            BMinY = MinY
            BMaxY = MaxY
        
        ElseIf Head = eHeading.EAST Then
            MinX = MinX + MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
            BMinX = MinX - MargenX * 2
            BMaxX = BMinX
            BMinY = MinY
            BMaxY = MaxY
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
            
            .AreasInfo.Users.RemoveAll
            .AreasInfo.NPCs.RemoveAll
            
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
        If MaxX > XMaxMapSize Then MaxX = XMaxMapSize
        
        If BMinY < 1 Then BMinY = 1
        If BMinX < 1 Then BMinX = 1
        If BMaxY > YMaxMapSize Then BMaxY = YMaxMapSize
        If BMaxX > XMaxMapSize Then BMaxX = XMaxMapSize
        
        map = .Pos.map
        
        'Esto es para ke el cliente elimine lo "fuera de area..."
        'Con el nuevo sistema de areas el cliente lo limpia solo al moverse
        
        
        For Each Elemento In .AreasInfo.Users.Items
            TempInt = Elemento
            If UserList(TempInt).Pos.map <> .Pos.map Or Abs(UserList(TempInt).Pos.X - .Pos.X) > MargenX Or Abs(UserList(TempInt).Pos.Y - .Pos.Y) > MargenY Then
                .AreasInfo.Users.Remove (TempInt)
                UserList(TempInt).AreasInfo.Users.Remove (UserIndex)
                'Les aviso a todos los users q no veo mas que sali de su pantalla
                Call WriteCharacterRemove(TempInt, UserList(UserIndex).Char.CharIndex)
                QuitoAlgo = True
            End If
        Next
        For Each Elemento In .AreasInfo.NPCs.Items
            TempInt = Elemento
            If Npclist(TempInt).Pos.map <> .Pos.map Or Abs(Npclist(TempInt).Pos.X - .Pos.X) > MargenX Or Abs(Npclist(TempInt).Pos.Y - .Pos.Y) > MargenY Then
                .AreasInfo.NPCs.Remove (TempInt)
                Npclist(TempInt).AreasInfo.Users.Remove (UserIndex)
                QuitoAlgo = True
            End If
        Next
        
        For X = 0 To 1
            If .AreasInfo.Barco(X) > 0 Then
                Call Barcos(.AreasInfo.Barco(X)).CheckUser(UserIndex)
            End If
        Next X
        
        If Head <> USER_NUEVO Then
            If Not QuitoAlgo Then
                For X = BMinX To BMaxX
                    For Y = BMinY To BMaxY
                        TempInt = MapData(map, X, Y).ObjInfo.ObjIndex
                        If TempInt > 0 Then
                            If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                                QuitoAlgo = True
                                Exit For
                            End If
                        End If
                    Next Y
                    If QuitoAlgo Then Exit For
                Next X
            End If
            If QuitoAlgo Then
                'Aviso al cliente que limpie toda la linea que desplazo
                Call WriteAreaChanged(UserIndex)
            End If
        End If
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                
                '<<< User >>>
                TempInt = MapData(map, X, Y).UserIndex
                If TempInt Then
                    Call .AreasInfo.Users.Add(TempInt, TempInt)
                    If UserIndex <> TempInt Then
                        Call UserList(TempInt).AreasInfo.Users.Add(UserIndex, UserIndex)
                        Call MakeUserChar(False, UserIndex, TempInt, map, X, Y)
                        Call MakeUserChar(False, TempInt, UserIndex, map, .Pos.X, .Pos.Y)
                        
                        'Si el user estaba invisible le avisamos al nuevo cliente de eso
                        If UserList(TempInt).flags.invisible Or UserList(TempInt).flags.Oculto Then
                            Call WriteSetInvisible(UserIndex, UserList(TempInt).Char.CharIndex, True)
                        End If
                        If UserList(UserIndex).flags.invisible Or UserList(UserIndex).flags.Oculto Then
                            Call WriteSetInvisible(TempInt, UserList(UserIndex).Char.CharIndex, True)
                        End If
                                
                        Call FlushBuffer(TempInt)
                            
                    ElseIf Head = USER_NUEVO Then
                        Call MakeUserChar(False, UserIndex, UserIndex, map, X, Y)
                    End If
                End If
                
                '<<< Npc >>>
                TempInt = MapData(map, X, Y).NpcIndex
                If TempInt Then
                    Call Npclist(TempInt).AreasInfo.Users.Add(UserIndex, UserIndex)
                    Call .AreasInfo.NPCs.Add(TempInt, TempInt)
                    Call MakeNPCChar(False, UserIndex, TempInt, map, X, Y)
                    'Call WriteCharacterCreate(UserIndex, Npclist(TempInt).Char.Body, Npclist(TempInt).Char.Head, Npclist(TempInt).Char.heading, Npclist(TempInt).Char.CharIndex, X, Y, 0, 0, 0, 0, 0, vbNullString, 0, 0)
                End If
                 
                '<<< Item >>>
                TempInt = MapData(map, X, Y).ObjInfo.ObjIndex
                If TempInt Then
                    If Not EsObjetoFijo(ObjData(TempInt).OBJType) Then
                        Call WriteObjectCreate(UserIndex, ObjData(TempInt).GrhIndex, X, Y)
                        
                        If ObjData(TempInt).OBJType = eOBJType.otPuertas Then
                            Call Bloquear(False, UserIndex, X, Y, MapData(map, X, Y).Blocked)
                            Call Bloquear(False, UserIndex, X - 1, Y, MapData(map, X - 1, Y).Blocked)
                        End If
                    End If
                End If
                
                Set Barco = BarcoEn(X, Y)
                If Not Barco Is Nothing Then
                    Call Barco.AgregarVisible(UserIndex)
                End If
            Next Y
        Next X
        
        'End If 'Chequeo que haya hecho 3 pasos
    End With
Call FlushBuffer(UserIndex)
Exit Sub
errh:
LimpiarAreasUser (UserIndex)
LogError ("Error en CheckUpdateNeededUser. Número " & Err.Number & " Descripción: " & Err.Description)
End Sub

Public Sub CheckUpdateNeededNpc(ByVal NpcIndex As Integer, ByVal Head As Byte)
    Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer
    Dim TempInt As Integer
    Dim Elemento
On Error GoTo errh:
    With Npclist(NpcIndex)
        MinX = .Pos.X
        MaxX = MinX
        MinY = .Pos.Y
        MaxY = MinY
        
        If Head = eHeading.NORTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MinY
        ElseIf Head = eHeading.SOUTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY + MargenY
            MaxY = MinY
        
        ElseIf Head = eHeading.WEST Then
            MinX = MinX - MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
        
        ElseIf Head = eHeading.EAST Then
            MinX = MinX + MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
        If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

        
        For Each Elemento In .AreasInfo.Users.Items
            TempInt = Elemento
            If UserList(TempInt).Pos.map <> .Pos.map Or Abs(UserList(TempInt).Pos.X - .Pos.X) > MargenX Or Abs(UserList(TempInt).Pos.Y - .Pos.Y) > MargenY Then
                .AreasInfo.Users.Remove (TempInt)
                UserList(TempInt).AreasInfo.NPCs.Remove (NpcIndex)
                'Les aviso a todos los users q no veo mas que sali de su pantalla
                Call WriteCharacterRemove(TempInt, Npclist(NpcIndex).Char.CharIndex)
                Call FlushBuffer(TempInt)
            End If
        Next
        
        
        'Actualizamos!!!
            For X = MinX To MaxX
                For Y = MinY To MaxY
                    TempInt = MapData(.Pos.map, X, Y).UserIndex
                    If TempInt Then
                        Call UserList(TempInt).AreasInfo.NPCs.Add(NpcIndex, NpcIndex)
                        Call .AreasInfo.Users.Add(TempInt, TempInt)
                        Call MakeNPCChar(False, TempInt, NpcIndex, .Pos.map, .Pos.X, .Pos.Y)
                        'Call WriteCharacterCreate(TempInt, Npclist(NpcIndex).Char.Body, Npclist(NpcIndex).Char.Head, Npclist(NpcIndex).Char.heading, Npclist(NpcIndex).Char.CharIndex, .Pos.X, .Pos.Y, 0, 0, 0, 0, 0, vbNullString, 0, 0)

                        Call FlushBuffer(TempInt)
                    End If
                Next Y
            Next X
    End With
Exit Sub
errh:
LimpiarAreasNpc (NpcIndex)
LogError ("Error en CheckUpdateNeededNpc. Número " & Err.Number & " Descripción: " & Err.Description)
End Sub

Public Sub CheckUpdateNeededBarco(ByRef Barco As clsBarco, ByVal Head As Byte)
    Dim MinX As Integer, MaxX As Integer, MinY As Integer, MaxY As Integer, X As Integer, Y As Integer
    Dim TempInt As Integer
    Dim i As Integer
    Dim Elemento
On Error GoTo errh:
    With Barco
        MinX = .X
        MaxX = MinX
        MinY = .Y
        MaxY = MinY
        
        If Head = eHeading.NORTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MinY
        ElseIf Head = eHeading.SOUTH Then
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY + MargenY
            MaxY = MinY
        
        ElseIf Head = eHeading.WEST Then
            MinX = MinX - MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
        
        ElseIf Head = eHeading.EAST Then
            MinX = MinX + MargenX
            MaxX = MinX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        
           
        ElseIf Head = USER_NUEVO Then
            'Esto pasa por cuando cambiamos de mapa o logeamos...
            MinX = MinX - MargenX
            MaxX = MaxX + MargenX
            MinY = MinY - MargenY
            MaxY = MaxY + MargenY
        End If
        
        If MinY < 1 Then MinY = 1
        If MinX < 1 Then MinX = 1
        If MaxY > YMaxMapSize Then MaxY = YMaxMapSize
        If MaxX > XMaxMapSize Then MaxX = XMaxMapSize

        
        For Each Elemento In .UsersVisibles
            TempInt = Elemento
            If UserList(TempInt).Pos.map <> 1 Or Abs(UserList(TempInt).Pos.X - .X) > MargenX Or Abs(UserList(TempInt).Pos.Y - .Y) > MargenY And UserList(TempInt).flags.Embarcado <> Barco.index Then
                .UsersVisibles.Remove (TempInt)
            End If
        Next
        
        Dim OtroBarco As clsBarco
        
        'Actualizamos!!!
        For X = MinX To MaxX
            For Y = MinY To MaxY
                TempInt = MapData(1, X, Y).UserIndex
                If TempInt Then
                    .AgregarVisible (TempInt)
                End If
                Set OtroBarco = BarcoEn(X, Y)
                If Not OtroBarco Is Nothing Then
                    If OtroBarco.index <> Barco.index Then
                        For i = 1 To 4
                            TempInt = Barco.GetPasajero(i)
                            If TempInt Then
                                Call OtroBarco.AgregarVisible(TempInt)
                            End If
                        Next i
                    End If
                End If
            Next Y
        Next X
        
        
        
    End With
Exit Sub
errh:
LogError ("Error en CheckUpdateNeededBarco. Número " & Err.Number & " Descripción: " & Err.Description)
End Sub

Public Sub AgregarUser(ByVal UserIndex As Integer, ByVal map As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: 04/01/2007
'Modified by Juan Martín Sotuyo Dodero (Maraxus)
'   - Now the method checks for repetead users instead of trusting parameters.
'   - If the character is new to the map, update it
'**************************************************************
    Dim TempVal As Long
    Dim i As Long
    
    If Not MapaValido(map) Then Exit Sub
    
    LimpiarAreasUser (UserIndex)
    
    Call CheckUpdateNeededUser(UserIndex, USER_NUEVO)
End Sub

Public Sub AgregarNpc(ByVal NpcIndex As Integer)
'**************************************************************
'Author: Lucio N. Tourrilhes (DuNga)
'Last Modify Date: Unknow
'
'**************************************************************
    LimpiarAreasNpc (NpcIndex)
    
    Call CheckUpdateNeededNpc(NpcIndex, USER_NUEVO)
End Sub

Public Sub LimpiarAreasUser(ByVal UserIndex As Integer)
Dim Elemento
Dim i As Integer

For Each Elemento In UserList(UserIndex).AreasInfo.Users
    i = Elemento
    UserList(i).AreasInfo.Users.Remove (UserIndex)
Next Elemento
For Each Elemento In UserList(UserIndex).AreasInfo.NPCs
    i = Elemento
    Npclist(i).AreasInfo.Users.Remove (UserIndex)
Next Elemento

UserList(UserIndex).AreasInfo.Users.RemoveAll
UserList(UserIndex).AreasInfo.NPCs.RemoveAll
End Sub
Public Sub LimpiarAreasNpc(ByVal NpcIndex As Integer)
Dim Elemento
Dim i As Integer

For Each Elemento In Npclist(NpcIndex).AreasInfo.Users
    i = Elemento
    UserList(i).AreasInfo.NPCs.Remove (NpcIndex)
Next Elemento

Npclist(NpcIndex).AreasInfo.Users.RemoveAll
End Sub

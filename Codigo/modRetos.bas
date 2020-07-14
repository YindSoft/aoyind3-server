Attribute VB_Name = "modRetos"
Option Explicit
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%% Hecho por El Yind 21/01/2011 %%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%          Correcciones            %%
'%% Fecha:      Correción:           %%
'%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
Public Const MAX_SOLRETOS = 250
Public Const VALOR_RETOS = 5000
Public Const MAX_APUESTA_RETO = 500000
Public Const MIN_APUESTA_RETO = 5000
Public Const NUM_SALAS = 25
Private Const SALA1_X = 347
Private Const SALA1_Y = 525
Private Const SALA_ANCHO = 20
Private Const SALA_ALTO = 19
Private Const ANCHOSALA = 30
Private Const ATRIL1_X = 5
Private Const ATRIL1_Y = 3
Private Const ATRIL2_X = 15
Private Const ATRIL2_Y = 3
Private Const EQUIPO1_X = 1
Private Const EQUIPO1_Y = 6
Private Const EQUIPO2_X = 19
Private Const EQUIPO2_Y = 18
Private Const SEGUNDOS_RETO_ITEMS As Integer = 30
Public Type tReto
    Active As Boolean
    Vs As Byte
    PorItems As Boolean
    Oro As Long
    Pj(1 To 4) As Integer
    APj(1 To 4) As Boolean
End Type
Public Type tSalaReto
    Active As Boolean
    RetoIndex As Integer
    PrimerRonda As Byte
    SegundaRonda As Byte
    X As Integer
    Y As Integer
    Segundos As Integer
End Type

Public Retos(1 To MAX_SOLRETOS) As tReto
Public LastRetoIndex As Integer
Public SalasRetos(1 To NUM_SALAS) As tSalaReto


Public Sub RetosDecide(UserIndex As Integer, Decide As Boolean)
Dim RI As Integer
Dim PjsIndex(1 To 4) As Integer
Dim i As Integer
Dim Validados As Byte
Dim Cantidad As Byte
Dim MyIndex As Integer

RI = UserList(UserIndex).RetoIndex
If RI > 0 And UserList(UserIndex).SalaIndex = 0 Then
    With Retos(RI)
        Cantidad = IIf(.Vs = 1, 2, 4)
        Validados = 1
        If Decide Then
            For i = 2 To Cantidad
                If .APj(i) Then Validados = Validados + 1
                If .Pj(i) = UserIndex Then MyIndex = i
            Next i
            If AptoParaReto(RI, UserIndex, True) Then
                .APj(MyIndex) = True
                Validados = Validados + 1
                If Validados = Cantidad Then
                    'Revalido a todos los participantes para evitar problemas
                    Validados = 0
                    For i = 1 To Cantidad
                        If AptoParaReto(RI, .Pj(i), False) Then Validados = Validados + 1
                    Next i
                    'Estan todos aptos
                    If Validados = Cantidad Then
                        Call WriteRetosRespuesta(UserIndex, 1)
                        Call WriteConsoleMsg(UserIndex, "Has aceptado el reto de " & UserList(.Pj(1)).name, FontTypeNames.FONTTYPE_RETOS)
                        Call ComenzarReto(RI)
                    Else
                        Call WriteRetosRespuesta(UserIndex, 15)
                    End If
                Else
                    Call WriteRetosRespuesta(UserIndex, 1)
                    Call WriteConsoleMsg(UserIndex, "Has aceptado el reto de " & UserList(.Pj(1)).name, FontTypeNames.FONTTYPE_RETOS)
                End If
            End If
        Else
            For i = 1 To Cantidad
                UserList(.Pj(i)).RetoIndex = 0
                If .Pj(i) = UserIndex Then
                    Call WriteConsoleMsg(UserIndex, "Has rechazado la invitación al reto. El reto ha sido cancelado.", FontTypeNames.FONTTYPE_RETOS)
                Else
                    Call WriteConsoleMsg(.Pj(i), UserList(UserIndex).name & " ha rechazado la invitación al reto. El reto ha sido cancelado.", FontTypeNames.FONTTYPE_RETOS)
                End If
            Next i
            .Active = False
            Call WriteRetosRespuesta(UserIndex, 1)
        End If
    End With
End If
End Sub
Public Function AptoParaReto(RI As Integer, UserIndex As Integer, Informar As Boolean) As Boolean
If UserList(UserIndex).flags.Muerto = 1 Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 11)
    AptoParaReto = False
ElseIf UserList(UserIndex).Stats.GLD < Retos(RI).Oro + VALOR_RETOS Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 2)
    AptoParaReto = False
ElseIf UserList(UserIndex).Counters.Pena > 0 Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 12)
    AptoParaReto = False
ElseIf Zonas(UserList(UserIndex).Zona).Segura = 0 Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 3)
    AptoParaReto = False
ElseIf UserList(UserIndex).flags.Comerciando Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 13)
    AptoParaReto = False
ElseIf UserList(UserIndex).flags.Navegando = 1 Then
    If Informar Then Call WriteRetosRespuesta(UserIndex, 14)
    AptoParaReto = False
Else
    AptoParaReto = True
End If
End Function
Public Sub CancelarReto(UserIndex As Integer)
Dim RI As Integer
Dim SI As Integer
Dim i As Integer
Dim Ganador As Byte

RI = UserList(UserIndex).RetoIndex
SI = UserList(UserIndex).SalaIndex
With Retos(RI)
    If Retos(RI).Active Then
        For i = 1 To IIf(.Vs = 1, 2, 4)
            UserList(.Pj(i)).RetoIndex = 0
            If .Pj(i) <> UserIndex Then
                Call WriteConsoleMsg(.Pj(i), UserList(UserIndex).name & " ha salido del juego. El reto ha sido cancelado.", FontTypeNames.FONTTYPE_RETOS)
            Else
                Ganador = i Mod 2
            End If
        Next i
    End If
End With
If SI > 0 Then
    'Si un user cierra en medio de un reto, gana el equipo contrario automaticamente
    'Se podria hacer que le de un minuto para volver a reconectar...
    Call FinalizarReto(SI, IIf(Ganador = 1, 2, 1))
Else
    Call ResetReto(RI)
End If
End Sub

Public Function ComenzarReto(RI As Integer) As Boolean
Dim i As Integer
Dim X As Integer
Dim index As Integer
For i = 1 To NUM_SALAS
    If SalasRetos(i).Active = False Then Exit For
Next i

If i > NUM_SALAS Then
    ComenzarReto = False
Else
    With SalasRetos(i)
        .Active = True
        .X = SALA1_X + ((i - 1) * ANCHOSALA)
        .Y = SALA1_Y
        .RetoIndex = RI
        .PrimerRonda = 0
        .SegundaRonda = 0
        
        'Guardo donde estaban los pjs antes de comenzar el reto
        For X = 1 To IIf(Retos(RI).Vs = 1, 2, 4)
            index = Retos(RI).Pj(X)
            UserList(index).RetoAntPos = UserList(index).Pos
            'Les quito el oro de la apuesta
            UserList(index).Stats.GLD = UserList(index).Stats.GLD - Retos(RI).Oro - VALOR_RETOS
            Call WriteUpdateGold(index)
            
        Next X
        
        Call ComenzarRonda(i)
    End With
End If
End Function

Public Sub MuereEnSala(UserIndex As Integer)
Dim SI As Integer
Dim RI As Integer
Dim ParejaIndex As Integer
Dim Pareja As Integer
Dim i As Integer
SI = UserList(UserIndex).SalaIndex
RI = UserList(UserIndex).RetoIndex

If Retos(RI).Vs = 1 Then
    Call FinalizarRonda(SI, IIf(Retos(RI).Pj(1) = UserIndex, 2, 1))
Else
    For i = 1 To 4
        If Retos(RI).Pj(i) = UserIndex Then
            ParejaIndex = Retos(RI).Pj(IIf(i = 2, 4, (i + 2) Mod 4))
            Exit For
        End If
    Next i
    Pareja = i Mod 2
    If UserList(ParejaIndex).flags.Muerto = 1 Then
        Call FinalizarRonda(SI, IIf(Pareja = 0, 1, 2))
    Else
        If Pareja = 1 Then
            'Atril de la paraja 1
            Call WarpUserChar(UserIndex, 2, SalasRetos(SI).X + ATRIL1_X, SalasRetos(SI).Y + ATRIL1_Y, False)
        Else
            'Atril de la pareja 2
            Call WarpUserChar(UserIndex, 2, SalasRetos(SI).X + ATRIL2_X, SalasRetos(SI).Y + ATRIL2_Y, False)
        End If
    End If
End If
End Sub

Public Sub ComenzarRonda(SI As Integer)
Dim X As Integer
Dim RI As Integer
RI = SalasRetos(SI).RetoIndex
For X = 1 To IIf(Retos(RI).Vs = 1, 2, 4)
    UserList(Retos(RI).Pj(X)).SalaIndex = SI
    Call WriteRetosRespuesta(Retos(RI).Pj(X), 10)
Next X
Call RevivirUsuarioReto(Retos(RI).Pj(1))
Call RevivirUsuarioReto(Retos(RI).Pj(2))
Call WarpUserChar(Retos(RI).Pj(1), 2, SalasRetos(SI).X + EQUIPO1_X, SalasRetos(SI).Y + EQUIPO1_Y, False)
Call WarpUserChar(Retos(RI).Pj(2), 2, SalasRetos(SI).X + EQUIPO2_X, SalasRetos(SI).Y + EQUIPO2_Y, False)
If Retos(RI).Vs = 2 Then
    Call RevivirUsuarioReto(Retos(RI).Pj(3))
    Call RevivirUsuarioReto(Retos(RI).Pj(4))
    Call WarpUserChar(Retos(RI).Pj(3), 2, SalasRetos(SI).X + EQUIPO1_X, SalasRetos(SI).Y + EQUIPO1_Y + 1, False)
    Call WarpUserChar(Retos(RI).Pj(4), 2, SalasRetos(SI).X + EQUIPO2_X, SalasRetos(SI).Y + EQUIPO2_Y - 1, False)
End If
End Sub
Public Sub FinalizarRonda(SI As Integer, Ganador As Byte)
Dim RI As Integer
RI = SalasRetos(SI).RetoIndex
Dim G1 As Integer, G2 As Integer
Dim Mensaje As String
Dim i As Integer
If Ganador = 1 Then
    G1 = Retos(RI).Pj(1)
    G2 = Retos(RI).Pj(3)
Else
    G1 = Retos(RI).Pj(2)
    G2 = Retos(RI).Pj(4)
End If

If SalasRetos(SI).PrimerRonda = 0 Then
    'Primer Round
    SalasRetos(SI).PrimerRonda = Ganador
    Call ComenzarRonda(SI)
    If Retos(RI).Vs = 1 Then
        Mensaje = "¡" & UserList(G1).name & " gana la primer ronda!"
    Else
        Mensaje = "¡" & UserList(G1).name & " y " & UserList(G2).name & " ganan la primer ronda!"
    End If
ElseIf SalasRetos(SI).PrimerRonda = Ganador Then
    'Gano el equipo
    Call FinalizarReto(SI, Ganador)
ElseIf SalasRetos(SI).SegundaRonda = 0 Then
    'Segundo Round
    SalasRetos(SI).SegundaRonda = Ganador
    Call ComenzarRonda(SI)
    If Retos(RI).Vs = 1 Then
        Mensaje = "¡" & UserList(G1).name & " gana la segunda ronda!"
    Else
        Mensaje = "¡" & UserList(G1).name & " y " & UserList(G2).name & " ganan la segunda ronda!"
    End If
Else
    'Gano el equipo
    Call FinalizarReto(SI, Ganador)
End If
If Mensaje <> "" Then
    For i = 1 To IIf(Retos(RI).Vs = 1, 2, 4)
        Call WriteConsoleMsg(Retos(RI).Pj(i), Mensaje, FontTypeNames.FONTTYPE_RETOS)
        Call FlushBuffer(Retos(RI).Pj(i))
    Next i
End If
End Sub
Public Sub FinalizarReto(SI As Integer, Ganador As Byte)
'Entrego los premios y devuelvo cada quien a donde estaba.
Dim RI As Integer
Dim G1 As Integer, G2 As Integer
Dim P1 As Integer, P2 As Integer
RI = SalasRetos(SI).RetoIndex
If Ganador = 1 Then
    G1 = Retos(RI).Pj(1)
    G2 = Retos(RI).Pj(3)
    P1 = Retos(RI).Pj(2)
    P2 = Retos(RI).Pj(4)
Else
    G1 = Retos(RI).Pj(2)
    G2 = Retos(RI).Pj(4)
    P1 = Retos(RI).Pj(1)
    P2 = Retos(RI).Pj(3)
End If

UserList(G1).Stats.GLD = UserList(G1).Stats.GLD + Retos(RI).Oro * 2
Call WriteUpdateGold(G1)

Call WarpUserChar(P1, UserList(P1).RetoAntPos.map, UserList(P1).RetoAntPos.X, UserList(P1).RetoAntPos.Y, True)
UserList(G1).RetoIndex = 0
UserList(G1).SalaIndex = 0
'UserList(G1).RetosGanados = UserList(G1).RetosGanados + 1
UserList(P1).RetoIndex = 0
UserList(P1).SalaIndex = 0
'UserList(P1).RetosPerdidos = UserList(P1).RetosPerdidos - 1
    
If Retos(RI).Vs = 2 Then
    UserList(G2).Stats.GLD = UserList(G2).Stats.GLD + Retos(RI).Oro * 2
    Call WriteUpdateGold(G2)
    Call WarpUserChar(P2, UserList(P2).RetoAntPos.map, UserList(P2).RetoAntPos.X, UserList(P2).RetoAntPos.Y, True)

    'Libero el buffer porque sino quedan mal los mensajes :S
    Call FlushBuffer(G1)
    Call FlushBuffer(G2)

    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(G1).name & " y " & UserList(G2).name & " le han ganado el reto a " & UserList(P1).name & " y " & UserList(P2).name & " por " & Retos(RI).Oro & " monedas de oro" & IIf(Retos(RI).PorItems, " y los items", ""), FontTypeNames.FONTTYPE_RETOS))
    UserList(G2).RetoIndex = 0
    UserList(G2).SalaIndex = 0
    'UserList(G2).RetosGanados = UserList(G2).RetosGanados + 1
    UserList(P2).RetoIndex = 0
    UserList(P2).SalaIndex = 0
    'UserList(P2).RetosPerdidos = UserList(P2).RetosPerdidos - 1
Else
    'Libero el buffer porque sino quedan mal los mensajes :S
    Call FlushBuffer(G1)
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg(UserList(G1).name & " le ha ganado el reto a " & UserList(P1).name & " por " & Retos(RI).Oro & " monedas de oro" & IIf(Retos(RI).PorItems, " y los items", ""), FontTypeNames.FONTTYPE_RETOS))
End If

If Retos(RI).PorItems Then
    If UserList(G1).flags.Muerto = 1 Then
        'Si lo habian matado antes
        Call RevivirUsuarioReto(G1)
        Call WarpUserChar(G1, 2, UserList(G1).Pos.X, UserList(G1).Pos.Y + 7, False)
    End If

    Call TirarTodoReto(P1, G1)

    'Si es por los items les dejo 30 segundos para agarrar las cosas y depositar
    Call WriteConsoleMsg(G1, "Tienes " & SEGUNDOS_RETO_ITEMS & " segundos para recoger los items ganados.", FontTypeNames.FONTTYPE_RETOS)
    If Retos(RI).Vs = 2 Then
        If UserList(G2).flags.Muerto = 1 Then
            'Si lo habian matado antes
            Call RevivirUsuarioReto(G2)
            Call WarpUserChar(G2, 2, UserList(G2).Pos.X, UserList(G2).Pos.Y + 7, False)
        End If
    
        Call TirarTodoReto(P2, G2)
        Call WriteConsoleMsg(G2, "Tienes " & SEGUNDOS_RETO_ITEMS & " segundos para recoger los items ganados.", FontTypeNames.FONTTYPE_RETOS)
    End If
    SalasRetos(SI).Segundos = SEGUNDOS_RETO_ITEMS
Else
    'Limpiamos la sala y transportamos a los ganadores
    Call LimpiarSala(SI)
End If
Call ResetReto(RI)
End Sub
Sub TirarTodoReto(ByVal UserIndex As Integer, ByVal GanadorIndex As Integer)
    Dim i As Byte
    Dim NuevaPos As WorldPos
    Dim MiObj As Obj
    Dim ItemIndex As Integer
    
    For i = 1 To MAX_INVENTORY_SLOTS
        ItemIndex = UserList(UserIndex).Invent.Object(i).ObjIndex
        If ItemIndex > 0 Then
             If ItemSeCae(ItemIndex) Then
                NuevaPos.X = 0
                NuevaPos.Y = 0
                
                'Creo el Obj
                MiObj.Amount = UserList(UserIndex).Invent.Object(i).Amount
                MiObj.ObjIndex = ItemIndex

                'Busco un tile cerca del ganador
                Tilelibre UserList(GanadorIndex).Pos, NuevaPos, MiObj, True, True
                
                If NuevaPos.X <> 0 And NuevaPos.Y <> 0 Then
                    Call DropObj(UserIndex, i, MAX_INVENTORY_OBJS, NuevaPos.map, NuevaPos.X, NuevaPos.Y)
                End If
             End If
        End If
    Next i
End Sub
Public Sub LimpiarSala(SI As Integer)
Dim X As Integer
Dim Y As Integer
Dim tmpInt As Integer
With SalasRetos(SI)
    .Active = False
    .PrimerRonda = 0
    .RetoIndex = 0
    .SegundaRonda = 0
    .Segundos = 0
    For X = .X To .X + SALA_ANCHO
        For Y = .Y To .Y + SALA_ALTO
            'Borramos los objetos que haya en el piso
            tmpInt = MapData(2, X, Y).ObjInfo.ObjIndex
            If tmpInt > 0 Then
                MapData(2, X, Y).ObjInfo.ObjIndex = 0
                MapData(2, X, Y).ObjInfo.Amount = 0
            End If
            'Teletransportamos a los usuarios que esten dentro
            tmpInt = MapData(2, X, Y).UserIndex
            If tmpInt > 0 Then
                If UserList(tmpInt).RetoAntPos.map = 0 Or UserList(tmpInt).RetoAntPos.X = 0 Or UserList(tmpInt).RetoAntPos.Y = 0 Then
                    'Si no tenia una posicion anterior lo llevamos a Ullathorpe :/
                    UserList(tmpInt).RetoAntPos = Ullathorpe
                End If
                Call WarpUserChar(tmpInt, UserList(tmpInt).RetoAntPos.map, UserList(tmpInt).RetoAntPos.X, UserList(tmpInt).RetoAntPos.Y, True)
            End If
        Next Y
    Next X
End With
End Sub
Public Sub ResetReto(RI As Integer)
Dim i As Integer
With Retos(RI)
    .Active = False
    .Oro = 0
    .PorItems = False
    .Vs = 1
    For i = 1 To 4
        .Pj(i) = 0
        .APj(i) = False
    Next i
End With
End Sub


Public Sub RevivirUsuarioReto(tUser As Integer)
If UserList(tUser).flags.Muerto = 1 Then
    UserList(tUser).flags.Muerto = 0
                        
    Call DarCuerpoDesnudo(tUser)
    UserList(tUser).Char.Head = UserList(tUser).OrigChar.Head
Else
    '<<<< Paralisis >>>>
    If UserList(tUser).flags.Paralizado = 1 Then
        UserList(tUser).flags.Paralizado = 0
        Call WriteParalizeOK(tUser)
    End If
        
    '<<< Estupidez >>>
    If UserList(tUser).flags.Estupidez = 1 Then
        UserList(tUser).flags.Estupidez = 0
        Call WriteDumbNoMore(tUser)
    End If
End If
UserList(tUser).Stats.MinHP = UserList(tUser).Stats.MaxHP
UserList(tUser).Stats.MinSta = UserList(tUser).Stats.MaxSta
                
Call WriteUpdateHP(tUser)
Call WriteUpdateSta(tUser)
End Sub

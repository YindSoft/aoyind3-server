VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMercader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
'%%%%%   Creado por El Yind   %%%%%
'%%%%%       23/01/2012       %%%%%
'%%%%% Build: 2               %%%%%
'%%%%% Modificado: 24/02/2012 %%%%%
'%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Option Explicit
Private Const MAX_MANA As Integer = 2000
Private Const ORO_INICIAL As Integer = 5000
Private Const INCREMENTO_ORO As Integer = 300
Private Const MAX_ORO As Integer = 15000
Private Const TIEMPO_ESPERA As Integer = 1 'Minutos
Private Const NUM_ROJAS As Integer = 500
Private Const NUM_AZULES As Integer = 2500

Private Camino() As Position
Private pCantPasos As Byte
Private pPaso As Byte
Private pNpcIndex As Integer
Private pNroNPC As Integer
Private pRuta As clsTree
Private startX As Integer, startY As Integer
Private Yendo As Boolean
Private Tiempo As Date
Private Agresores As Dictionary
Private Direccion As Byte
Private Recalcular As Boolean
Private DestinoX As Integer
Private DestinoY As Integer
Private LastAgresor As Integer
Private estado As Byte
Private Destino(1 To 2) As String

'INVENTARIO
Private potasRojas As Integer
Private potasAzules As Integer
Private Oro As Long


'DEFENSA
Private CPRemo As Integer
Private CPDialogo As Integer
Private CPAtaca As Integer
Private BardeoInvis As Boolean
Private Mana As Integer
Private Estrategia As Byte 'Se usa para definir la manera en que el npc va a jugar
Private Combear As Boolean 'Se usa cuando dio un golpe para priorizar un ataque magico
Private Escapar As Boolean 'Se usa para que si le esta pegando cuerpo a cuerpo y el random dio, se corra para no quedarse siempre ahi
Private Emergencia As Boolean 'Se usa por si el npc le queda poca vida que no ataque hasta tomar una pocion
Private PasosEnemigo As Integer 'Cuenta cuantos pasos hace con alguien cerca que rompa las bolas

Private LanzoAtaque As Integer

Private Const HECHI_REMO As Integer = 10
Private Const HECHI_INMO As Integer = 24
Private Const HECHI_APOCA As Integer = 25
Private Const HECHI_DESCARGA As Integer = 23
Private Const HECHI_REMOINVI As Integer = 38
Private CalcularInmo As Boolean
Public Property Get NpcIndex() As Integer
NpcIndex = pNpcIndex
End Property

Public Property Let NpcIndex(ByVal lNewValue As Integer)
pNpcIndex = lNewValue
End Property

Public Property Get NpcNum() As Integer
NpcNum = pNroNPC
End Property

Public Property Let NpcNum(ByVal lNewValue As Integer)
pNroNPC = lNewValue
End Property

Public Sub Init(ByVal NroNPC As Integer, Ruta As String, Ciudad1 As String, Ciudad2 As String)
pNroNPC = NroNPC

Dim Rutas() As String
Rutas = Split(Ruta, ";")
Dim i As Integer
pCantPasos = UBound(Rutas)
ReDim Camino(pCantPasos)
For i = 0 To pCantPasos
    Camino(i).X = val(ReadField(1, Rutas(i), 44))
    Camino(i).Y = val(ReadField(2, Rutas(i), 44))
Next i

Destino(1) = Ciudad1
Destino(2) = Ciudad2

Set Agresores = New Dictionary

ReSpawn

End Sub

Private Function SiguientePaso() As Boolean
If Yendo Then
    If pPaso < pCantPasos Then
        pPaso = pPaso + 1
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        BuscarCamino
        SiguientePaso = True
    Else
        Yendo = False
        ResetMercader
        SiguientePaso = False
    End If
Else
    
    If pPaso > 0 Then
        pPaso = pPaso - 1
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        BuscarCamino
        SiguientePaso = True
    Else
        Yendo = True
        ResetMercader
        SiguientePaso = False
    End If
End If
End Function

Private Sub ResetMercader()
potasRojas = NUM_ROJAS
potasAzules = NUM_AZULES

Mana = MAX_MANA
Npclist(pNpcIndex).Stats.MinHP = Npclist(pNpcIndex).Stats.MaxHP

Oro = Oro + INCREMENTO_ORO * MultiplicadorORO
If Oro > MAX_ORO * MultiplicadorORO Then Oro = MAX_ORO * MultiplicadorORO
Npclist(pNpcIndex).GiveGLDMin = Oro
Npclist(pNpcIndex).GiveGLDMax = Oro

Npclist(pNpcIndex).Invent.Object(1).Amount = potasRojas
Npclist(pNpcIndex).Invent.Object(2).Amount = potasAzules

Emergencia = False
Recalcular = False
Combear = False
Estrategia = 0

Set Agresores = New Dictionary
Set pRuta = Nothing

LastAgresor = -1
Tiempo = Now
estado = 0
End Sub

Private Function BuscarCamino() As Boolean
Set pRuta = New clsTree
BuscarCamino = pRuta.RunSearch(1, Npclist(pNpcIndex).Pos.X, Npclist(pNpcIndex).Pos.Y, DestinoX, DestinoY)
End Function


Public Sub MoverMercader()
On Error GoTo Errhandler:
With Npclist(pNpcIndex)
Dim agresorIndex As Integer
Dim tmpIndex As Integer
Dim tHeading As Byte
Dim X As Integer, Y As Integer
Dim nPos As WorldPos
Dim i As Integer

    'Si tiene poca vida que priorize curarse.
    If .Stats.MinHP <= 210 Then
        Emergencia = True
    End If

    'Si tiene potas rojas y esta herido se cura solo :D
    If .Stats.MinHP < .Stats.MaxHP And (LanzoAtaque <= 3 Or CPAtaca >= 3) Then
        If potasRojas > 0 Then
            potasRojas = potasRojas - 1
            Npclist(pNpcIndex).Invent.Object(1).Amount = potasRojas
            .Stats.MinHP = .Stats.MinHP + 30
            If .Stats.MinHP > .Stats.MaxHP Then
                .Stats.MinHP = .Stats.MaxHP
            End If
            Emergencia = False
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
    End If
    'Si tiene potas azules que restaure mana
    If Mana < MAX_MANA And LanzoAtaque <= 4 Then
        If potasAzules > 0 Then
            potasAzules = potasAzules - 1
            Npclist(pNpcIndex).Invent.Object(2).Amount = potasAzules
            Mana = Mana + MAX_MANA * 0.035
            If Mana > MAX_MANA Then
                Mana = MAX_MANA
            End If
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessagePlayWave(SND_BEBER, .Pos.X, .Pos.Y))
        End If
    End If
    
    If .flags.Inmovilizado = 1 Or .flags.Paralizado = 1 Then
        CPRemo = CPRemo + 1
        If CPRemo > 4 And LanzoAtaque < 2 And RandomNumber(1, 3) = 1 Then
            If NpcLanzaSpellSobreNpcComoUser(pNpcIndex, pNpcIndex, HECHI_REMO, Mana) Then
                CPRemo = 0
                LanzoAtaque = 9
            End If
        End If
    End If

    agresorIndex = SeleccionarAgresor
    
    If agresorIndex <> LastAgresor Then
        'Si cambia de agresor le borro la ruta que tenia
        LastAgresor = agresorIndex
        CalcularInmo = False
        Set pRuta = Nothing
    End If
    If agresorIndex > 0 Then
        If estado = 1 Then
            LanzoAtaque = 12
        End If
        estado = 2
        If LanzoAtaque > 0 Then LanzoAtaque = LanzoAtaque - 1
        'Busco users atacantes que esten al lado
        If Agresores.Exists(MapData(.Pos.map, .Pos.X - 1, .Pos.Y).UserIndex) Then
            tmpIndex = MapData(.Pos.map, .Pos.X - 1, .Pos.Y).UserIndex
            tHeading = eHeading.WEST
        ElseIf Agresores.Exists(MapData(.Pos.map, .Pos.X + 1, .Pos.Y).UserIndex) Then
            tmpIndex = MapData(.Pos.map, .Pos.X + 1, .Pos.Y).UserIndex
            tHeading = eHeading.EAST
        ElseIf Agresores.Exists(MapData(.Pos.map, .Pos.X, .Pos.Y - 1).UserIndex) Then
            tmpIndex = MapData(.Pos.map, .Pos.X, .Pos.Y - 1).UserIndex
            tHeading = eHeading.NORTH
        ElseIf Agresores.Exists(MapData(.Pos.map, .Pos.X, .Pos.Y + 1).UserIndex) Then
            tmpIndex = MapData(.Pos.map, .Pos.X, .Pos.Y + 1).UserIndex
            tHeading = eHeading.SOUTH
        End If
        If .flags.Paralizado = 1 And tHeading <> .Char.Heading Then tHeading = 0
        If tHeading > 0 And LanzoAtaque = 0 And Not Emergencia Then
            'Si tengo a un boludo al lado mientras estoy inmo lo ataco.
            If tHeading <> .Char.Heading Then Call ChangeNPCChar(pNpcIndex, .Char.Body, .Char.Head, tHeading)
            Call NpcAtacaUser(pNpcIndex, tmpIndex)
            LanzoAtaque = 2
            Combear = True
            If RandomNumber(1, 4) <> 1 Then
                Exit Sub
            End If
        ElseIf tHeading > 0 And LanzoAtaque > 0 And Not Emergencia Then
            'Giro pero no ataco
            Call ChangeNPCChar(pNpcIndex, .Char.Body, .Char.Head, tHeading)
            Exit Sub
        Else
            Call AtacaConMagia(agresorIndex)
        End If
    
        If UserList(agresorIndex).flags.Inmovilizado = 1 And Not Escapar And Not Emergencia Then
            'Si lo inmovilize voy al lado
            If Not CalcularInmo Then
                Set pRuta = New clsTree
                DestinoX = UserList(agresorIndex).Pos.X
                DestinoY = UserList(agresorIndex).Pos.Y
                If pRuta.RunSearch(1, Npclist(pNpcIndex).Pos.X, Npclist(pNpcIndex).Pos.Y, DestinoX, DestinoY) Then
                    CalcularInmo = True
                End If
            ElseIf .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                If Not SeguirRuta Then
                    Escapar = True
                End If
            End If
        Else
            If Not Emergencia Then 'No quiero que deje de hacer lo que estaba haciendo, solo que tome una pocion y siga
                CalcularInmo = False
                Set pRuta = Nothing
            End If
            If UserList(agresorIndex).flags.Inmovilizado = 0 Then Escapar = False
            
            If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 Then
                Recalcular = True
                If Direccion = 0 Then
                    Direccion = .Char.Heading
                ElseIf RandomNumber(1, 2) = 1 Then
                    Direccion = RandomNumber(1, 4)
                    If .Char.Heading = Direccion Then Direccion = RandomNumber(1, 4)
                End If
                nPos = .Pos
                Call HeadtoPos(Direccion, nPos)
                If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                    MoveNPCChar pNpcIndex, Direccion
                Else
                    For i = 1 To 4
                        If i <> Direccion Then
                            nPos = .Pos
                            Call HeadtoPos(i, nPos)
                            If LegalPosNPC(.Pos.map, nPos.X, nPos.Y, 0, True) Then
                                MoveNPCChar pNpcIndex, i
                                Exit For
                            End If
                        End If
                    Next i
                End If
                CPDialogo = CPDialogo + 1
                If CPDialogo > 25 And RandomNumber(1, 10) = 1 Then
                    CPDialogo = 0
                    'Mensajes que puede tirar el bicho al azar
                    Select Case RandomNumber(1, 3)
                        Case 1
                            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead("¡Bastardo! No lograrás robarme las mercancías mientras siga con vida.", .Char.CharIndex, vbWhite))
                        Case 2
                            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead("Te has cruzado con la persona equivocada.", .Char.CharIndex, vbWhite))
                        Case 3
                            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead("No te dejaré salir con vida.", .Char.CharIndex, vbWhite))
                    End Select
                End If
            End If
        
        
        End If
    ElseIf Recalcular Then
        Recalcular = False
        BardeoInvis = False
        DestinoX = Camino(pPaso).X
        DestinoY = Camino(pPaso).Y
        If BuscarCamino Then MoverMercader
    ElseIf Not pRuta Is Nothing Then
        If .flags.Inmovilizado = 0 And .flags.Paralizado = 0 And agresorIndex = 0 Then
            estado = 1
            'Debug.Print ("Paso: " & startX & "  " & startY & " - " & PasoActual & " - " & Camino(PasoActual).X & "  " & Camino(PasoActual).Y)
            If Not SeguirRuta Then
                If SiguientePaso Then
                    MoverMercader
                End If
            End If
            
            'Si alguien me sigue verifico cuanto viene siguiendo y si estuvo por mas de 20 pasos lo ataco.
            Dim Enemigo As Integer
            Enemigo = SeleccionarEnemigo
            If Enemigo > 0 Then
                PasosEnemigo = PasosEnemigo + 1
                If PasosEnemigo = 20 Then
                    If .Stats.Alineacion = 0 Then
                        Call WriteChatOverHead(Enemigo, "¡Si continúas siguiendome me veré obligado a atacarte!", .Char.CharIndex, vbWhite)
                    Else
                        Call WriteChatOverHead(Enemigo, "¡Aléjate patán o sufrirás mi ira!", .Char.CharIndex, vbWhite)
                    End If
                ElseIf PasosEnemigo >= 30 Then
                    AgregarAgresor (Enemigo)
                    PasosEnemigo = 22
                End If
            ElseIf PasosEnemigo > 0 Then
                PasosEnemigo = PasosEnemigo - 1
            End If
        End If
    ElseIf DateDiff("s", Tiempo, Now) >= TIEMPO_ESPERA * 60 Then
        estado = 1
        SiguientePaso
    End If
End With

Exit Sub
Errhandler:
Exit Sub
End Sub

Private Function SeguirRuta() As Boolean
Dim tHeading As Byte
startX = Npclist(pNpcIndex).Pos.X
startY = Npclist(pNpcIndex).Pos.Y
If pRuta.PathStepNext(startX, startY) > 0 Then
    Dim tmpPos As WorldPos
    tmpPos.map = Npclist(pNpcIndex).Pos.map
    tmpPos.X = startX
    tmpPos.Y = startY
    tHeading = FindDirection(Npclist(pNpcIndex).Pos, tmpPos)
            
    If LegalPosNPC(tmpPos.map, tmpPos.X, tmpPos.Y, 0, True) Then
        MoveNPCChar pNpcIndex, tHeading
    ElseIf DestinoX <> tmpPos.X Or DestinoY <> tmpPos.Y Then 'Si no puedo ir al lugar que tengo que ir no importa
        If BuscarCamino Then
            MoverMercader
        End If
    Else
        SeguirRuta = False
        Exit Function
    End If
    SeguirRuta = True
Else
    SeguirRuta = False
End If
End Function

Public Sub ReSpawn()
pPaso = 0
Yendo = True
Oro = ORO_INICIAL * MultiplicadorORO
Dim Pos As WorldPos
Pos.map = 1
Pos.X = Camino(0).X
Pos.Y = Camino(0).Y
pNpcIndex = SpawnNpc(pNroNPC, Pos, True, True, 1)
Npclist(pNpcIndex).GiveGLDMin = Oro
Npclist(pNpcIndex).GiveGLDMax = Oro
Call CheckZonaNPC(pNpcIndex)
ResetMercader
End Sub

Public Sub AgregarAgresor(ByVal UserIndex As Integer)
If Not Agresores.Exists(UserIndex) Then
    Call Agresores.Add(UserIndex, UserIndex)
    Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead("¡Lamentarás el momento en que te cruzaste conmigo rufián!", Npclist(pNpcIndex).Char.CharIndex, vbWhite))
End If
End Sub
Public Sub QuitarAgresor(ByVal UserIndex As Integer)
If Agresores.Exists(UserIndex) Then
    Call Agresores.Remove(UserIndex)
End If
End Sub


Private Function SeleccionarAgresor() As Integer
Dim i As Integer, UserIndex As Integer
Dim Candidato As Integer, CandidatoHP As Integer
Dim FactorMovimiento As Single
Candidato = 0
CandidatoHP = 2000
For i = 0 To Agresores.Count - 1
    UserIndex = Agresores.Items(i)
    If Abs(Npclist(pNpcIndex).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y And _
       Abs(Npclist(pNpcIndex).Pos.X - UserList(UserIndex).Pos.X) <= RANGO_VISION_X And _
       UserList(UserIndex).flags.Muerto = 0 And _
       Zonas(UserList(UserIndex).Zona).Segura = 0 And _
       Zonas(Npclist(pNpcIndex).Zona).Segura = 0 Then
        If UserList(UserIndex).flags.invisible = 1 And LastAgresor = UserIndex And Not BardeoInvis Then
            BardeoInvis = True
            LanzoAtaque = 7
            Call SendData(SendTarget.ToNPCArea, pNpcIndex, PrepareMessageChatOverHead("jaja, ¿piensas que te puedes ocultar de mi?", Npclist(pNpcIndex).Char.CharIndex, vbWhite))
        End If
        'Este calculo sirve para priorizar los pjs que no se estan moviendo sobre los que tienen baja vida. (Revisar)
        FactorMovimiento = (1 + (UserList(UserIndex).flags.Movimiento + 20) / 200)
        If UserList(UserIndex).Stats.MinHP * FactorMovimiento < CandidatoHP Then
            Candidato = UserIndex
            CandidatoHP = UserList(UserIndex).Stats.MinHP * FactorMovimiento
        End If
    End If
Next i
SeleccionarAgresor = Candidato
End Function

Private Function SeleccionarEnemigo() As Integer
Dim i As Integer, UserIndex As Integer
SeleccionarEnemigo = 0
For i = 0 To Npclist(pNpcIndex).AreasInfo.Users.Count - 1
    UserIndex = Npclist(pNpcIndex).AreasInfo.Users.Items(i)
    If Abs(Npclist(pNpcIndex).Pos.Y - UserList(UserIndex).Pos.Y) <= RANGO_VISION_Y And _
       Abs(Npclist(pNpcIndex).Pos.X - UserList(UserIndex).Pos.X) <= RANGO_VISION_X And _
       UserList(UserIndex).flags.Muerto = 0 And _
       Zonas(UserList(UserIndex).Zona).Segura = 0 And _
       Zonas(Npclist(pNpcIndex).Zona).Segura = 0 And _
       Not Agresores.Exists(UserIndex) And ((Npclist(pNpcIndex).Stats.Alineacion = 0 And Criminal(UserIndex)) Or (Npclist(pNpcIndex).Stats.Alineacion = 1 And UserList(UserIndex).Faccion.FuerzasCaos = 0)) Then
        SeleccionarEnemigo = UserIndex
        Exit For
    End If
Next i
End Function

Private Sub AtacaConMagia(UserIndex As Integer)
CPAtaca = CPAtaca + 1

If Emergencia Then Exit Sub

Debug.Print UserList(UserIndex).flags.Movimiento & "  " & Mana & "  " & Estrategia & "  -  " & LanzoAtaque

If RandomNumber(1, 100) = 1 Then
    Estrategia = Estrategia + 1
    If Estrategia = 3 Then Estrategia = 0
End If

If LanzoAtaque = 0 Then 'Un intervalo de pasos minimo por cada ataque

    If UserList(UserIndex).flags.invisible = 1 Then 'Si esta invisible lo unico que hago es quitarle el invi primero
        If RandomNumber(1, UserList(UserIndex).flags.Movimiento + 3) <= 2 Then
            If Mana >= Hechizos(HECHI_REMOINVI).ManaRequerido Then
                If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_REMOINVI, Mana) Then
                    CPAtaca = 0
                    LanzoAtaque = 8
                End If
            End If
        End If
    Else
        If UserList(UserIndex).flags.Inmovilizado = 0 And (Estrategia = 0 Or Estrategia = 2 Or RandomNumber(1, 5) = 1) And RandomNumber(1, 10) <> 1 And Combear = False Then
            If RandomNumber(1, UserList(UserIndex).flags.Movimiento + 3) = 1 Then
                If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_INMO, Mana) Then
                    LanzoAtaque = 6
                    CPAtaca = 0
                End If
            End If
        Else
            If RandomNumber(1, IIf(UserList(UserIndex).flags.Inmovilizado Or Combear, 2, UserList(UserIndex).flags.Movimiento + 3)) = 1 Then
                Combear = False 'Si pego un golpe priozo tirar un hechizo
                If Mana >= Hechizos(HECHI_APOCA).ManaRequerido Then
                    If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_APOCA, Mana) Then
                        CPAtaca = 0
                        LanzoAtaque = 6
                    End If
                Else 'Si no tiene mana para apoca prueba con descarga
                    If NpcLanzaSpellSobreUserComoUser(pNpcIndex, UserIndex, HECHI_DESCARGA, Mana) Then
                        CPAtaca = 0
                        LanzoAtaque = 5
                    End If
                End If
            End If
        End If
    End If
End If
End Sub
Public Sub Clicked(ByVal UserIndex As Integer)
Dim Status As String
Dim ft As FontTypeNames
If Npclist(pNpcIndex).Stats.Alineacion = 0 Then
    Status = "Ves a " & Npclist(pNpcIndex).Name & " <Ejército Real> <Mercader Real> [CONSEJO DE BANDERBILL]"
    ft = FontTypeNames.FONTTYPE_CONSEJOVesA
Else
    Status = "Ves a " & Npclist(pNpcIndex).Name & " <Legión Oscura> <Mercader Oscuro> [CONCILIO DE LAS SOMBRAS]"
    ft = FontTypeNames.FONTTYPE_CONSEJOCAOSVesA
End If
Call WriteConsoleMsg(UserIndex, Status, ft)

If Npclist(pNpcIndex).Stats.Alineacion = 0 Then
    If estado = 0 Then
        Call WriteChatOverHead(UserIndex, "Cuando termine de negociar partiré hacia " & Destino(IIf(Yendo, 2, 1)) & ", tendré que tener cuidado, dicen que esas rutas son peligrosas.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
    ElseIf estado = 1 Then
        Call WriteChatOverHead(UserIndex, "No tengo tiempo de charlar debo entregar estas mercancías.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
    End If
Else
    If estado = 0 Then
        Call WriteChatOverHead(UserIndex, "Cuando termine aquí iré hacia " & Destino(IIf(Yendo, 2, 1)) & ", quien ose estorbarme en el camino morirá.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
    ElseIf estado = 1 Then
        Call WriteChatOverHead(UserIndex, "No me estorbes patán o sufrirás mi ira.", Npclist(pNpcIndex).Char.CharIndex, vbWhite)
    End If

End If
End Sub

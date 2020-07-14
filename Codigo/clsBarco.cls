VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBarco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const ACELERACION As Byte = 60
Private Const VELOCIDAD As Single = 7
Private Const ScrollPixelsPerFrameX As Single = 8
Private Const EngineBase As Single = 0.018

Dim RutaBarco() As Position
Public Paso As Byte
Dim Pos As Position
Dim Direccion As Byte
Public UPasos As Byte
Dim FactorBarco As Single
Dim LastTick As Long
Dim BOffSetX As Single
Dim BOffSetY As Single
Dim Sentido As Byte
Public index As Byte

Dim EnMovimiento As Boolean
Public TickPuerto As Long
Dim Puerto As Byte

Dim Pasajeros(0 To 4) As Integer

Public Pasos As Long

Public UsersVisibles As Dictionary
Dim Updating As Boolean


Public Property Get X() As Integer
    X = Pos.X
End Property
Public Property Get Y() As Integer
    Y = Pos.Y
End Property

Public Function GetPasajero(ByVal Num As Byte) As Integer
GetPasajero = Pasajeros(Num)
End Function

Public Sub Init(ByVal Ruta As String, ByVal pPaso As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal TiempoPuerto As Long, ByVal mSentido As Byte, ByVal mIndex As Byte)
Dim Rutas() As String
Rutas = Split(Ruta, ";")
Dim i As Integer
UPasos = UBound(Rutas)
ReDim RutaBarco(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(i).X = val(ReadField(1, Rutas(i), 44))
    RutaBarco(i).Y = val(ReadField(2, Rutas(i), 44))
Next i

Set UsersVisibles = New Dictionary

Sentido = mSentido
index = mIndex

Paso = pPaso
Pos.X = X
Pos.Y = Y

For i = IIf(Sentido = 0, 1, NUM_PUERTOS) To IIf(Sentido = 0, NUM_PUERTOS, 1) Step IIf(Sentido = 0, 1, -1)
    If Paso > Puertos(i).Paso(Sentido) Then
        Puerto = i
    End If
Next i
Puerto = NextPuerto

If TiempoPuerto > 0 Then
    TickPuerto = (GetTickCount() And &H7FFFFFFF) - TiempoPuerto
    EnMovimiento = False
Else
    EnMovimiento = True
End If

Pasajeros(0) = 0

If Pos.X > RutaBarco(Paso).X Then
    Direccion = WEST
ElseIf Pos.X < RutaBarco(Paso).X Then
    Direccion = EAST
ElseIf Pos.Y > RutaBarco(Paso).Y Then
    Direccion = NORTH
ElseIf Pos.Y < RutaBarco(Paso).Y Then
    Direccion = SOUTH
End If
Call CalcularFactorBarco
End Sub
Private Function NextPuerto() As Integer
Dim P As Integer
If Sentido = 0 Then
    P = Puerto + 1
    If P > NUM_PUERTOS Then P = 1
Else
    P = Puerto - 1
    If P < 1 Then P = NUM_PUERTOS
End If
NextPuerto = P
End Function
Private Function BackPuerto() As Integer
Dim P As Integer
If Sentido = 0 Then
    P = Puerto - 1
    If P < 1 Then P = NUM_PUERTOS
Else
    P = Puerto + 1
    If P > NUM_PUERTOS Then P = 1
End If
BackPuerto = P
End Function
Public Function AgregarPasajero(ByVal UserIndex As Integer) As Boolean
Dim i As Integer
Dim Encontro As Boolean
For i = 1 To 4
    If Pasajeros(i) = 0 Then
        Pasajeros(i) = UserIndex
        Encontro = True
        Exit For
    End If
Next i

If Encontro Then
    With UserList(UserIndex)
        MapData(.Pos.map, .Pos.X, .Pos.Y).UserIndex = 0
        .flags.Embarcado = index
        .Pos.X = Pos.X
        .Pos.Y = Pos.Y
        Call CheckUpdateNeededUser(UserIndex, 0)
        Call SendDataVisibles(PrepareMessageAgregarPasajero(UserIndex, Sentido, i))
        'Le mandamos el mensaje para que borre el personaje a los clientes que estén cerca
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCharacterRemove(.Char.CharIndex))
        
        UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - VALOR_BILLETE
        Call WriteUpdateGold(UserIndex)
    End With
End If
AgregarPasajero = Encontro
End Function
Public Sub QuitarPasajero(ByVal UserIndex As Integer, Optional ByVal Warp As Boolean = True)
Dim i As Integer
Dim Encontro As Boolean
For i = 1 To 4
    If Pasajeros(i) = UserIndex Then
        Pasajeros(i) = 0
        Encontro = True
        Exit For
    End If
Next i

If Encontro Then
    UserList(UserIndex).flags.Embarcado = 0
    Call SendDataVisibles(PrepareMessageQuitarPasajero(UserIndex, Sentido, i))
    
    Dim bPos As WorldPos
    Dim nPos As WorldPos
    
    If Not UserList(UserIndex).flags.UserLogged Then
        bPos.map = 1
        bPos.X = RutaBarco(Puertos(BackPuerto).Paso(Sentido)).X
        bPos.Y = RutaBarco(Puertos(BackPuerto).Paso(Sentido)).Y
    Else
        bPos.map = 1
        bPos.X = Pos.X
        bPos.Y = Pos.Y
    End If
    If Warp Then
        Call ClosestLegalPos(bPos, nPos)
    
        Call WarpUserChar(UserIndex, nPos.map, nPos.X, nPos.Y, False)
    End If
End If
End Sub

Private Sub SendDataVisibles(ByVal sdData As String)
    Dim tempIndex As Integer
    Dim Elemento
    For Each Elemento In UsersVisibles
        tempIndex = Elemento
        If UserList(tempIndex).ConnIDValida Then
            If Abs(UserList(tempIndex).Pos.X - Pos.X) <= MargenX And Abs(UserList(tempIndex).Pos.Y - Pos.Y) <= MargenY Then
                Call EnviarDatosASlot(tempIndex, sdData)
            Else
                UsersVisibles.Remove (tempIndex)
            End If
        End If
    Next
End Sub

Public Sub AgregarVisible(ByVal UserIndex As Integer)
If Not UsersVisibles.Exists(UserIndex) Then
     Call UsersVisibles.Add(UserIndex, UserIndex)
     UserList(UserIndex).AreasInfo.Barco(Sentido) = index
     Call WriteShowBarco(UserIndex, Sentido, Paso, Pos.X, Pos.Y, IIf(TickPuerto > 0, (GetTickCount() And &H7FFFFFFF) - TickPuerto, 0), Pasajeros)
End If
End Sub

Public Sub CheckUser(ByVal UserIndex As Integer)
If Abs(UserList(UserIndex).Pos.X - Pos.X) > MargenX And Abs(UserList(UserIndex).Pos.Y - Pos.Y) > MargenY Then
    QuitarVisible (UserIndex)
    Call WriteQuitarBarco(UserIndex, Sentido)
End If
End Sub
Public Sub QuitarVisible(ByVal UserIndex As Integer)
If UsersVisibles.Exists(UserIndex) Then
    UsersVisibles.Remove (UserIndex)
End If
UserList(UserIndex).AreasInfo.Barco(Sentido) = 0

Call QuitarPasajero(UserIndex)
End Sub

Private Sub CalcularFactorBarco()
Dim NP As Byte
Dim Dist As Integer, Dist2 As Integer

NP = Puertos(Puerto).Paso(Sentido)
Dist = Abs(RutaBarco(NP).X - Pos.X) + Abs(RutaBarco(NP).Y - Pos.Y)
If Puerto - 1 = 0 Then
    NP = Puertos(NUM_PUERTOS).Paso(Sentido)
Else
    NP = Puertos(Puerto - 1).Paso(Sentido)
End If

'Sistema de aceleracion, era muy brusco sino cuando el barco llegaba ;)
Dist2 = Abs(RutaBarco(NP).X - Pos.X) + Abs(RutaBarco(NP).Y - Pos.Y)
If Dist <= ACELERACION Then
    FactorBarco = (Dist / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
ElseIf Dist2 <= ACELERACION Then
    FactorBarco = (Dist2 / ACELERACION) * (VELOCIDAD - 0.5) + 0.5
Else
    FactorBarco = 7
End If
End Sub

Public Sub Avanzar()
Dim NP As Byte
Dim i As Integer
Dim e As Integer

frmMain.pBarcos.ForeColor = vbWhite
frmMain.pBarcos.PSet (Pos.X / 5, Pos.Y / 5)

Select Case Direccion
    Case NORTH
        Pos.Y = Pos.Y - 1
        For i = Pos.X - 1 To Pos.X + 1
            e = MapData(1, i, Pos.Y).UserIndex
            If e > 0 Then
                If Not EsGM(e) Then
                    Call UserDie(e)
                End If
            End If
        Next i
    Case SOUTH
        Pos.Y = Pos.Y + 1
        For i = Pos.X - 1 To Pos.X + 1
            e = MapData(1, i, Pos.Y).UserIndex
            If e > 0 Then
                If Not EsGM(e) Then
                    Call UserDie(e)
                End If
            End If
        Next i
    Case WEST
        Pos.X = Pos.X - 1
        For i = Pos.Y - 1 To Pos.Y + 1
            e = MapData(1, Pos.X, i).UserIndex
            If e > 0 Then
                If Not EsGM(e) Then
                    Call UserDie(e)
                End If
            End If
        Next i
    Case EAST
        Pos.X = Pos.X + 1
        For i = Pos.Y - 1 To Pos.Y + 1
            e = MapData(1, Pos.X, i).UserIndex
            If e > 0 Then
                If Not EsGM(e) Then
                    Call UserDie(e)
                End If
            End If
        Next i
End Select



Call CheckUpdateNeededBarco(Me, Direccion)

Pasos = Pasos + 1
Call CalcularFactorBarco

For i = 1 To 4
    If Pasajeros(i) > 0 Then
        UserList(Pasajeros(i)).Pos.X = Pos.X
        UserList(Pasajeros(i)).Pos.Y = Pos.Y
        If FactorBarco < VELOCIDAD / 2 Then
            Updating = True
            Call CheckUpdateNeededUser(Pasajeros(i), Direccion)
        ElseIf FactorBarco >= VELOCIDAD / 2 And Updating Then
            Call LimpiarAreasUser(Pasajeros(i))
            For e = 1 To 4
                If Pasajeros(e) > 0 And e <> i Then
                    Call UserList(Pasajeros(i)).AreasInfo.Users.Add(Pasajeros(e), Pasajeros(e))
                End If
            Next e
        End If
    End If
Next i

If FactorBarco >= VELOCIDAD / 2 Then Updating = False

If Pos.X = RutaBarco(Paso).X And Pos.Y = RutaBarco(Paso).Y Then
    If EsPuerto(Paso) Then
        TickPuerto = (GetTickCount() And &H7FFFFFFF)
        Puerto = NextPuerto
    End If
    Paso = NextPaso
    
    If Pos.X > RutaBarco(Paso).X Then
        Direccion = WEST
    ElseIf Pos.X < RutaBarco(Paso).X Then
        Direccion = EAST
    ElseIf Pos.Y > RutaBarco(Paso).Y Then
        Direccion = NORTH
    ElseIf Pos.Y < RutaBarco(Paso).Y Then
        Direccion = SOUTH
    End If
End If
If Sentido = 0 Then
    frmMain.pBarcos.ForeColor = vbRed
Else
    frmMain.pBarcos.ForeColor = vbBlue
End If
frmMain.pBarcos.PSet (Pos.X / 5, Pos.Y / 5)
End Sub
Private Function NextPaso() As Byte
If Paso + 1 > UPasos Then
    NextPaso = 0
Else
    NextPaso = Paso + 1
End If
End Function
Private Function EsPuerto(ByVal Paso As Byte) As Boolean
Dim i As Integer
For i = 1 To NUM_PUERTOS
    If Puertos(i).Paso(Sentido) = Paso Then
        EsPuerto = True
        Exit For
    End If
Next i
End Function
Public Sub MoverBarco()
    'MoveScreen (Direccion)
    EnMovimiento = True
End Sub

Public Sub Calcular(ByVal ElapsedTime As Single)
Dim Factor As Double
Dim timerTicksPerFrame As Single
Dim i As Integer
timerTicksPerFrame = EngineBase * ElapsedTime
   

If TickPuerto > 0 Then
    If (GetTickCount() And &H7FFFFFFF) - TickPuerto < TIEMPO_EN_PUERTO * 1000 Then
        EnMovimiento = False
    Else
        TickPuerto = 0
        SendDataVisibles (PrepareMessageShowBarco(Sentido, Paso, Pos.X, Pos.Y, 0, Pasajeros))
        EnMovimiento = True

    End If
End If

If EnMovimiento Then
    Factor = timerTicksPerFrame * FactorBarco * ScrollPixelsPerFrameX
    
    If Abs(BOffSetX) >= 32 Then
        BOffSetX = BOffSetX - Sgn(BOffSetX) * 32
        BOffSetY = 0
        Avanzar
    ElseIf Abs(BOffSetY) >= 32 Then
        BOffSetY = BOffSetY - Sgn(BOffSetY) * 32
        BOffSetX = 0
        Avanzar
    End If
    Select Case Direccion
        Case NORTH
            BOffSetY = BOffSetY - Factor
        Case SOUTH
            BOffSetY = BOffSetY + Factor
        Case WEST
            BOffSetX = BOffSetX - Factor
        Case EAST
            BOffSetX = BOffSetX + Factor
    End Select
End If
End Sub

Public Function EstimarTiempo() As Integer
Dim Tiempo As Long
If TickPuerto > 0 Then
    Tiempo = TIEMPO_EN_PUERTO - ((GetTickCount() And &H7FFFFFFF) - TickPuerto) / 1000
End If
Tiempo = Tiempo + (Abs(Pos.X - RutaBarco(Puertos(Puerto).Paso(Sentido)).X) + Abs(Pos.Y - RutaBarco(Puertos(Puerto).Paso(Sentido)).Y) + 25) / 27.5

EstimarTiempo = Tiempo
End Function

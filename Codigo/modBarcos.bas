Attribute VB_Name = "modBarcos"
Option Explicit

Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Public Const TIEMPO_EN_PUERTO As Long = 20

Public Const NUM_PUERTOS As Byte = 5
Public Const NUM_BARCOS As Byte = 10

Public Const PUERTO_NIX As Byte = 1
Public Const PUERTO_BANDER As Byte = 2
Public Const PUERTO_ARGHAL As Byte = 3
Public Const PUERTO_LINDOS As Byte = 4
Public Const PUERTO_ARKHEIN As Byte = 5

Public Const VALOR_BILLETE As Integer = 500


Public Type tPuerto
    Paso(0 To 1) As Byte
    Nombre As String
End Type

Public Type tPositions
    Ruta() As Position
End Type

Public Puertos(1 To NUM_PUERTOS) As tPuerto

Public Barcos(1 To NUM_BARCOS) As clsBarco

Dim RutaBarco(0 To 1) As tPositions

Dim InicioBarcos(0 To 1) As Byte
Dim FinBarcos(0 To 1) As Byte

Dim end_time As Currency
Dim timer_freq As Currency

Dim ttt As Long

Public Sub InitBarcos()
Dim sRutaHoraria As String
Dim sRutaAntihoraria As String
Dim i As Integer

Puertos(PUERTO_NIX).Nombre = "Nix"
Puertos(PUERTO_NIX).Paso(0) = 0
Puertos(PUERTO_NIX).Paso(1) = 21

Puertos(PUERTO_BANDER).Nombre = "Banderbill"
Puertos(PUERTO_BANDER).Paso(0) = 4
Puertos(PUERTO_BANDER).Paso(1) = 16

Puertos(PUERTO_ARGHAL).Nombre = "Arghal"
Puertos(PUERTO_ARGHAL).Paso(0) = 12
Puertos(PUERTO_ARGHAL).Paso(1) = 9

Puertos(PUERTO_LINDOS).Nombre = "Lindos"
Puertos(PUERTO_LINDOS).Paso(0) = 15
Puertos(PUERTO_LINDOS).Paso(1) = 5

Puertos(PUERTO_ARKHEIN).Nombre = "Arkhein"
Puertos(PUERTO_ARKHEIN).Paso(0) = 19
Puertos(PUERTO_ARKHEIN).Paso(1) = 0

sRutaHoraria = "161,1247;35,1247;35,22;302,22;302,55;303,55;566,55;566,65;635,65;635,54;800,54;800,307;801,307;870,307;870,999;887,999;870,999;870,1224;643,1224;643,1371;643,1472;195,1472;195,1266;169,1266;169,1247"
sRutaAntihoraria = "639,1383;647,1383;647,1228;874,1228;874,995;887,995;874,995;874,303;804,303;804,311;804,50;631,50;631,61;570,61;570,51;306,51;306,59;306,18;31,18;31,1251;165,1251;165,1243;165,1270;191,1270;191,1476;647,1476;647,1383"


Dim Rutas() As String
Dim UPasos As Integer

Rutas = Split(sRutaHoraria, ";")
UPasos = UBound(Rutas)
ReDim RutaBarco(0).Ruta(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(0).Ruta(i).X = val(ReadField(1, Rutas(i), 44))
    RutaBarco(0).Ruta(i).Y = val(ReadField(2, Rutas(i), 44))
Next i

Rutas = Split(sRutaAntihoraria, ";")
UPasos = UBound(Rutas)
ReDim RutaBarco(1).Ruta(0 To UPasos) As Position
For i = 0 To UPasos
    RutaBarco(1).Ruta(i).X = val(ReadField(1, Rutas(i), 44))
    RutaBarco(1).Ruta(i).Y = val(ReadField(2, Rutas(i), 44))
Next i


For i = 1 To NUM_BARCOS
    Set Barcos(i) = New clsBarco
Next i

Call Barcos(1).Init(sRutaHoraria, 1, 161, 1247, 1, 0, 1)
Call Barcos(2).Init(sRutaHoraria, 2, 35, 352, 0, 0, 2)
Call Barcos(3).Init(sRutaHoraria, 8, 580, 65, 0, 0, 3)
Call Barcos(4).Init(sRutaHoraria, 14, 870, 638, 0, 0, 4)
Call Barcos(5).Init(sRutaHoraria, 19, 643, 1343, 0, 0, 5)

Call Barcos(6).Init(sRutaAntihoraria, 1, 639, 1383, 1, 1, 6)
Call Barcos(7).Init(sRutaAntihoraria, 6, 889, 995, 9703, 1, 7)
Call Barcos(8).Init(sRutaAntihoraria, 10, 804, 311, 14742, 1, 8)
Call Barcos(9).Init(sRutaAntihoraria, 17, 306, 59, 19671, 1, 9)
Call Barcos(10).Init(sRutaAntihoraria, 20, 140, 1251, 0, 1, 10)

InicioBarcos(0) = 1
InicioBarcos(1) = 6
FinBarcos(0) = 5
FinBarcos(1) = 10

Call QueryPerformanceCounter(end_time)

ttt = (GetTickCount() And &H7FFFFFFF)
End Sub

Public Sub CalcularBarcos()

Dim i As Integer
Dim ElapsedTime As Single
If Barcos(1) Is Nothing Then Exit Sub
'DoEvents
'frmMain.Show
'frmMain.pBarcos.Cls

Dim start_time As Currency

    'Get the timer frequency
If timer_freq = 0 Then
    QueryPerformanceFrequency timer_freq
End If
    
    'Get current time
Call QueryPerformanceCounter(start_time)
    
    'Calculate elapsed time
ElapsedTime = (start_time - end_time) / timer_freq * 1000
'Get next end time
Call QueryPerformanceCounter(end_time)

For i = 1 To NUM_BARCOS
    Call Barcos(i).Calcular(ElapsedTime)
Next i

End Sub

Private Function DistanciaPasos(ByVal Paso1 As Byte, ByVal Paso2 As Integer, ByVal CantPasos As Integer) As Integer
If Paso1 >= Paso2 Then
    DistanciaPasos = Paso1 - Paso2
Else
    DistanciaPasos = Paso1 - Paso2 + CantPasos
End If
End Function

Public Sub HablaMarinero(ByVal UserIndex As Integer, ByVal NpcIndex As Integer, Optional ByVal Accion As Boolean = False)
Dim i As Integer
Dim NP As Integer
Dim X As Integer
Dim Y As Integer
Dim mPaso As Integer
Dim mIdPaso As Byte
Dim Puerto As Integer
Dim Sentido As Byte

Sentido = Npclist(NpcIndex).Stats.Alineacion

For i = 1 To NUM_PUERTOS
    X = 1
    If Abs(RutaBarco(Sentido).Ruta(Puertos(i).Paso(Sentido)).X - Npclist(NpcIndex).Pos.X) < 10 And Abs(RutaBarco(Sentido).Ruta(Puertos(i).Paso(Sentido)).Y - Npclist(NpcIndex).Pos.Y) < 10 Then
    
        Exit For
    End If
Next i
Puerto = i
If Sentido = 0 Then
    NP = i + 1
    If NP > NUM_PUERTOS Then NP = 1
Else
    NP = i - 1
    If NP < 1 Then NP = NUM_PUERTOS
End If
Dim Tiempo As Integer
mPaso = 1000
For i = InicioBarcos(Sentido) To FinBarcos(Sentido)
    If Barcos(i).Paso = Puertos(Puerto).Paso(Sentido) + 1 And Barcos(i).TickPuerto > 0 Then
        Tiempo = TIEMPO_EN_PUERTO - ((GetTickCount() And &H7FFFFFFF) - Barcos(i).TickPuerto) / 1000
        Exit For
    ElseIf DistanciaPasos(Puertos(Puerto).Paso(Sentido), Barcos(i).Paso, Barcos(i).UPasos) < mPaso Then
        mPaso = DistanciaPasos(Puertos(Puerto).Paso(Sentido), Barcos(i).Paso, Barcos(i).UPasos)
        mIdPaso = i
    End If
Next i
If Not Accion Then
    If Tiempo > 0 Then
        Call WriteChatOverHead(UserIndex, "El barco zarpará hacia el puerto de " & Puertos(NP).Nombre & " en " & Tiempo & " segundos.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    Else
        Call WriteChatOverHead(UserIndex, "El proximo barco con destino a " & Puertos(NP).Nombre & " llegará en " & Barcos(mIdPaso).EstimarTiempo & " segundos.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
    End If
Else
    If UserList(UserIndex).flags.Embarcado > 0 Then
        Call Barcos(i).QuitarPasajero(UserIndex)
    ElseIf Tiempo > 0 Then
    
        If Distancia(UserList(UserIndex).Pos, Npclist(UserList(UserIndex).flags.TargetNPC).Pos) > 2 Then
            Call WriteConsoleMsg(UserIndex, "Debes ponerte al lado del marinero para subir al barco.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).flags.Muerto = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás muerto!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Navegando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco si estás navegando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Descansar = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras estás descansando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Meditando = True Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras estés meditando.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.invisible = 1 Or UserList(UserIndex).flags.Oculto = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco estando invisible.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Inmovilizado = 1 Or UserList(UserIndex).flags.Paralizado = 1 Then
            Call WriteConsoleMsg(UserIndex, "¡Estás paralizado!", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        If UserList(UserIndex).flags.Comerciando = 1 Then
            Call WriteConsoleMsg(UserIndex, "No puedes subir al barco mientras comercias.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If UserList(UserIndex).Stats.GLD < VALOR_BILLETE Then
            Call WriteChatOverHead(UserIndex, "Lo lamento, pero el billete vale " & VALOR_BILLETE & " monedas de oro.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
            Exit Sub
        End If

        If Not Barcos(i).AgregarPasajero(UserIndex) Then
            Call WriteChatOverHead(UserIndex, "Lo lamento, el barco ya está completo, deberás esperar al próximo.", Npclist(NpcIndex).Char.CharIndex, vbWhite)
        End If
    End If
End If
End Sub

Public Function BarcoEn(ByVal X As Integer, ByVal Y As Integer) As clsBarco
Dim i As Byte
For i = 1 To NUM_BARCOS
    If Barcos(i).X = X And Barcos(i).Y = Y Then
        Set BarcoEn = Barcos(i)
        Exit Function
    End If
Next i
Set BarcoEn = Nothing
End Function

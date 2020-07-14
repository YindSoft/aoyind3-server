Attribute VB_Name = "ES"
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

Public Sub CargarSpawnList()
    Dim N As Integer, LoopC As Integer
    N = val(GetVar(App.Path & "\Dat\Invokar.dat", "INIT", "NumNPCs"))
    ReDim SpawnList(N) As tCriaturasEntrenador
    For LoopC = 1 To N
        SpawnList(LoopC).NpcIndex = val(GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NI" & LoopC))
        SpawnList(LoopC).NpcName = GetVar(App.Path & "\Dat\Invokar.dat", "LIST", "NN" & LoopC)
    Next LoopC
    
End Sub

Function EsAdmin(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Admines"))

For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Admines", "Admin" & WizNum))
    
    If left$(NomB, 1) = "*" Or left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsAdmin = True
        Exit Function
    End If
Next WizNum
EsAdmin = False

End Function

Function EsDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Dioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Dioses", "Dios" & WizNum))
    
    If left$(NomB, 1) = "*" Or left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsDios = True
        Exit Function
    End If
Next WizNum
EsDios = False
End Function

Function EsSemiDios(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "SemiDioses"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "SemiDioses", "SemiDios" & WizNum))
    
    If left$(NomB, 1) = "*" Or left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsSemiDios = True
        Exit Function
    End If
Next WizNum
EsSemiDios = False

End Function

Function EsConsejero(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "Consejeros"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "Consejeros", "Consejero" & WizNum))
    
    If left$(NomB, 1) = "*" Or left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsConsejero = True
        Exit Function
    End If
Next WizNum
EsConsejero = False
End Function

Function EsRolesMaster(ByVal Name As String) As Boolean
Dim NumWizs As Integer
Dim WizNum As Integer
Dim NomB As String

NumWizs = val(GetVar(IniPath & "Server.ini", "INIT", "RolesMasters"))
For WizNum = 1 To NumWizs
    NomB = UCase$(GetVar(IniPath & "Server.ini", "RolesMasters", "RM" & WizNum))
    
    If left$(NomB, 1) = "*" Or left$(NomB, 1) = "+" Then NomB = Right$(NomB, Len(NomB) - 1)
    If UCase$(Name) = NomB Then
        EsRolesMaster = True
        Exit Function
    End If
Next WizNum
EsRolesMaster = False
End Function


Public Function TxtDimension(ByVal Name As String) As Long
Dim N As Integer, cad As String, Tam As Long
N = FreeFile(1)
Open Name For Input As #N
Tam = 0
Do While Not EOF(N)
    Tam = Tam + 1
    Line Input #N, cad
Loop
Close N
TxtDimension = Tam
End Function

Public Sub CargarForbidenWords()

ReDim ForbidenNames(1 To TxtDimension(DatPath & "NombresInvalidos.txt"))
Dim N As Integer, i As Integer
N = FreeFile(1)
Open DatPath & "NombresInvalidos.txt" For Input As #N

For i = 1 To UBound(ForbidenNames)
    Line Input #N, ForbidenNames(i)
Next i

Close N

End Sub

Public Sub CargarHechizos()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'  ¡¡¡¡ NO USAR GetVar PARA LEER Hechizos.dat !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer Hechizos.dat se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

On Error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando Hechizos."

Dim Hechizo As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Hechizos.dat")

'obtiene el numero de hechizos
NumeroHechizos = val(Leer.GetValue("INIT", "NumeroHechizos"))

ReDim Hechizos(1 To NumeroHechizos) As tHechizo

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumeroHechizos
frmCargando.cargar.value = 0

'Llena la lista
For Hechizo = 1 To NumeroHechizos

    Hechizos(Hechizo).Nombre = Leer.GetValue("Hechizo" & Hechizo, "Nombre")
    Hechizos(Hechizo).desc = Leer.GetValue("Hechizo" & Hechizo, "Desc")
    Hechizos(Hechizo).PalabrasMagicas = Leer.GetValue("Hechizo" & Hechizo, "PalabrasMagicas")
    
    Hechizos(Hechizo).HechizeroMsg = Leer.GetValue("Hechizo" & Hechizo, "HechizeroMsg")
    Hechizos(Hechizo).TargetMsg = Leer.GetValue("Hechizo" & Hechizo, "TargetMsg")
    Hechizos(Hechizo).PropioMsg = Leer.GetValue("Hechizo" & Hechizo, "PropioMsg")
    
    Hechizos(Hechizo).Tipo = val(Leer.GetValue("Hechizo" & Hechizo, "Tipo"))
    Hechizos(Hechizo).Wav = val(Leer.GetValue("Hechizo" & Hechizo, "WAV"))
    Hechizos(Hechizo).FXgrh = val(Leer.GetValue("Hechizo" & Hechizo, "Fxgrh"))
    
    Hechizos(Hechizo).Loops = val(Leer.GetValue("Hechizo" & Hechizo, "Loops"))
    Hechizos(Hechizo).Efecto = val(Leer.GetValue("Hechizo" & Hechizo, "Efecto"))
    
'    Hechizos(Hechizo).Resis = val(Leer.GetValue("Hechizo" & Hechizo, "Resis"))
    
    Hechizos(Hechizo).SubeHP = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHP"))
    Hechizos(Hechizo).MinHP = val(Leer.GetValue("Hechizo" & Hechizo, "MinHP"))
    Hechizos(Hechizo).MaxHP = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHP"))
    
    Hechizos(Hechizo).SubeMana = val(Leer.GetValue("Hechizo" & Hechizo, "SubeMana"))
    Hechizos(Hechizo).MiMana = val(Leer.GetValue("Hechizo" & Hechizo, "MinMana"))
    Hechizos(Hechizo).MaMana = val(Leer.GetValue("Hechizo" & Hechizo, "MaxMana"))
    
    Hechizos(Hechizo).SubeSta = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSta"))
    Hechizos(Hechizo).MinSta = val(Leer.GetValue("Hechizo" & Hechizo, "MinSta"))
    Hechizos(Hechizo).MaxSta = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSta"))
    
    Hechizos(Hechizo).SubeHam = val(Leer.GetValue("Hechizo" & Hechizo, "SubeHam"))
    Hechizos(Hechizo).MinHam = val(Leer.GetValue("Hechizo" & Hechizo, "MinHam"))
    Hechizos(Hechizo).MaxHam = val(Leer.GetValue("Hechizo" & Hechizo, "MaxHam"))
    
    Hechizos(Hechizo).SubeSed = val(Leer.GetValue("Hechizo" & Hechizo, "SubeSed"))
    Hechizos(Hechizo).MinSed = val(Leer.GetValue("Hechizo" & Hechizo, "MinSed"))
    Hechizos(Hechizo).MaxSed = val(Leer.GetValue("Hechizo" & Hechizo, "MaxSed"))
    
    Hechizos(Hechizo).SubeAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "SubeAG"))
    Hechizos(Hechizo).MinAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MinAG"))
    Hechizos(Hechizo).MaxAgilidad = val(Leer.GetValue("Hechizo" & Hechizo, "MaxAG"))
    
    Hechizos(Hechizo).SubeFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "SubeFU"))
    Hechizos(Hechizo).MinFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MinFU"))
    Hechizos(Hechizo).MaxFuerza = val(Leer.GetValue("Hechizo" & Hechizo, "MaxFU"))
    
    Hechizos(Hechizo).SubeCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "SubeCA"))
    Hechizos(Hechizo).MinCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MinCA"))
    Hechizos(Hechizo).MaxCarisma = val(Leer.GetValue("Hechizo" & Hechizo, "MaxCA"))
    
    
    Hechizos(Hechizo).Invisibilidad = val(Leer.GetValue("Hechizo" & Hechizo, "Invisibilidad"))
    Hechizos(Hechizo).Paraliza = val(Leer.GetValue("Hechizo" & Hechizo, "Paraliza"))
    Hechizos(Hechizo).Inmoviliza = val(Leer.GetValue("Hechizo" & Hechizo, "Inmoviliza"))
    Hechizos(Hechizo).RemoverParalisis = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverParalisis"))
    Hechizos(Hechizo).RemoverEstupidez = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverEstupidez"))
    Hechizos(Hechizo).RemueveInvisibilidadParcial = val(Leer.GetValue("Hechizo" & Hechizo, "RemueveInvisibilidadParcial"))
    
    
    Hechizos(Hechizo).CuraVeneno = val(Leer.GetValue("Hechizo" & Hechizo, "CuraVeneno"))
    Hechizos(Hechizo).Envenena = val(Leer.GetValue("Hechizo" & Hechizo, "Envenena"))
    Hechizos(Hechizo).Maldicion = val(Leer.GetValue("Hechizo" & Hechizo, "Maldicion"))
    Hechizos(Hechizo).RemoverMaldicion = val(Leer.GetValue("Hechizo" & Hechizo, "RemoverMaldicion"))
    Hechizos(Hechizo).Bendicion = val(Leer.GetValue("Hechizo" & Hechizo, "Bendicion"))
    Hechizos(Hechizo).Revivir = val(Leer.GetValue("Hechizo" & Hechizo, "Revivir"))
    
    Hechizos(Hechizo).Ceguera = val(Leer.GetValue("Hechizo" & Hechizo, "Ceguera"))
    Hechizos(Hechizo).Estupidez = val(Leer.GetValue("Hechizo" & Hechizo, "Estupidez"))
    
    Hechizos(Hechizo).Invoca = val(Leer.GetValue("Hechizo" & Hechizo, "Invoca"))
    Hechizos(Hechizo).NumNpc = val(Leer.GetValue("Hechizo" & Hechizo, "NumNpc"))
    Hechizos(Hechizo).Cant = val(Leer.GetValue("Hechizo" & Hechizo, "Cant"))
    Hechizos(Hechizo).Mimetiza = val(Leer.GetValue("hechizo" & Hechizo, "Mimetiza"))
    
    
'    Hechizos(Hechizo).Materializa = val(Leer.GetValue("Hechizo" & Hechizo, "Materializa"))
'    Hechizos(Hechizo).ItemIndex = val(Leer.GetValue("Hechizo" & Hechizo, "ItemIndex"))
    
    Hechizos(Hechizo).MinSkill = val(Leer.GetValue("Hechizo" & Hechizo, "MinSkill"))
    Hechizos(Hechizo).ManaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "ManaRequerido"))
    
    'Barrin 30/9/03
    Hechizos(Hechizo).StaRequerido = val(Leer.GetValue("Hechizo" & Hechizo, "StaRequerido"))
    
    Hechizos(Hechizo).Target = val(Leer.GetValue("Hechizo" & Hechizo, "Target"))
    frmCargando.cargar.value = frmCargando.cargar.value + 1
    
    Hechizos(Hechizo).NeedStaff = val(Leer.GetValue("Hechizo" & Hechizo, "NeedStaff"))
    Hechizos(Hechizo).StaffAffected = CBool(val(Leer.GetValue("Hechizo" & Hechizo, "StaffAffected")))
    
Next Hechizo

Set Leer = Nothing
Exit Sub

Errhandler:
 MsgBox "Error cargando hechizos.dat " & Err.Number & ": " & Err.Description
 
End Sub

Sub LoadMotd()
Dim i As Integer

MaxLines = val(GetVar(App.Path & "\Dat\Motd.ini", "INIT", "NumLines"))

ReDim MOTD(1 To MaxLines)
For i = 1 To MaxLines
    MOTD(i).texto = GetVar(App.Path & "\Dat\Motd.ini", "Motd", "Line" & i)
    MOTD(i).Formato = vbNullString
Next i

End Sub

Public Sub HacerBackUp()
'Call LogTarea("Sub DoBackUp")
haciendoBK = True
Dim i As Integer



' Lo saco porque elimina elementales y mascotas - Maraxus
''''''''''''''lo pongo aca x sugernecia del yind
'For i = 1 To LastNPC
'    If Npclist(i).flags.NPCActive Then
'        If Npclist(i).Contadores.TiempoExistencia > 0 Then
'            Call MuereNpc(i, 0)
'        End If
'    End If
'Next i
'''''''''''/'lo pongo aca x sugernecia del yind



Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())


Call LimpiarMundo
Call WorldSave
Call modGuilds.v_RutinaElecciones
Call ResetCentinelaInfo     'Reseteamos al centinela


Call SendData(SendTarget.ToAll, 0, PrepareMessagePauseToggle())

'Call EstadisticasWeb.Informar(EVENTO_NUEVO_CLAN, 0)

haciendoBK = False

'Log
On Error Resume Next
Dim nfile As Integer
nfile = FreeFile ' obtenemos un canal
Open App.Path & "\logs\BackUps.log" For Append Shared As #nfile
Print #nfile, Date & " " & time
Close #nfile
End Sub

Public Sub GrabarMapa(ByVal map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim Y As Integer
    Dim X As Integer
    Dim TempInt As Integer
    Dim LoopC As Long
    
    If FileExist(MAPFILE & ".bak", vbNormal) Then
        Kill MAPFILE & ".bak"
    End If
    
    'Open .map file
    FreeFileMap = FreeFile
    Open MAPFILE & ".bak" For Binary As FreeFileMap
    Seek FreeFileMap, 1
    
    
    'Write .map file
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
                If (MapData(map, X, Y).Trigger = 4 Or MapData(map, X, Y).Trigger = 2 Or MapData(map, X, Y).Trigger = 1 Or MapData(map, X, Y).Trigger = 7) And MapData(map, X, Y).ObjInfo.ObjIndex Then
                    If ObjData(MapData(map, X, Y).ObjInfo.ObjIndex).Agarrable = 0 Then
                    Put FreeFileMap, , X
                    Put FreeFileMap, , Y
                    Put FreeFileMap, , MapData(map, X, Y).ObjInfo.ObjIndex
                    Put FreeFileMap, , MapData(map, X, Y).ObjInfo.Amount
                    End If
                End If
            
        Next X
    Next Y
    
    'Close .map file
    Close FreeFileMap


End Sub
Public Sub CargarBak(ByVal map As Long, ByVal MAPFILE As String)
On Error Resume Next
    Dim FreeFileMap As Long
    Dim Y As Integer
    Dim X As Integer
    Dim TempInt As Integer
    Dim LoopC As Long
    Dim ObjIndex As Integer
    Dim Cant As Integer
    
    If FileExist(MAPFILE & ".bak", vbNormal) Then
        
        'Open .map file
        FreeFileMap = FreeFile
        Open MAPFILE & ".bak" For Binary As FreeFileMap
        Seek FreeFileMap, 1
        
        
        Do While Not EOF(FreeFileMap)
            Get FreeFileMap, , X
            Get FreeFileMap, , Y
            Get FreeFileMap, , ObjIndex
            Get FreeFileMap, , Cant
            MapData(map, X, Y).ObjInfo.ObjIndex = ObjIndex
            MapData(map, X, Y).ObjInfo.Amount = Cant
        Loop
        
        'Close .map file
        Close FreeFileMap

    End If

End Sub
Sub LoadArmasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmasHerrero.dat", "INIT", "NumArmas"))

ReDim Preserve ArmasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmasHerrero(lc) = val(GetVar(DatPath & "ArmasHerrero.dat", "Arma" & lc, "Index"))
Next lc

End Sub
Sub CargarAreas()
'on error Resume Next
    Dim archivoC As String
    
    archivoC = DatPath & "areas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las areas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Integer
    Dim e As Integer
    Dim H As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim map As Byte
    Dim nPos As WorldPos
    nPos.map = 0
    nPos.X = 0
    nPos.Y = 0
    
    NumAreas = GetVar(archivoC, "Config", "Cantidad")
    If NumAreas > 0 Then
    ReDim Areas(1 To NumAreas)
    For i = 1 To NumAreas
        map = CByte(GetVar(archivoC, "Area" & CStr(i), "Mapa"))
        Areas(i).mapa = map
        Areas(i).X1 = CInt(GetVar(archivoC, "Area" & CStr(i), "X1"))
        Areas(i).Y1 = CInt(GetVar(archivoC, "Area" & CStr(i), "Y1"))
        Areas(i).X2 = CInt(GetVar(archivoC, "Area" & CStr(i), "X2"))
        Areas(i).Y2 = CInt(GetVar(archivoC, "Area" & CStr(i), "Y2"))
        Areas(i).NPCs = CByte(GetVar(archivoC, "Area" & CStr(i), "Npcs"))
        If Areas(i).NPCs > 0 Then
        ReDim Areas(i).NPC(1 To Areas(i).NPCs)
        For e = 1 To Areas(i).NPCs
            Areas(i).NPC(e).NpcIndex = val(GetVar(archivoC, "Area" & CStr(i), "Npc" & e))
            Areas(i).NPC(e).Cantidad = val(GetVar(archivoC, "Area" & CStr(i), "Cant" & e))
            
            For H = 1 To Areas(i).NPC(e).Cantidad
                    Call CrearNPC(Areas(i).NPC(e).NpcIndex, i, nPos)
            Next H
            
        Next e
        End If
    Next i
    End If
End Sub
Sub CargarZonas()
'on error Resume Next
    Dim archivoC As String
    
    archivoC = DatPath & "zonas.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar las zonas. Falta el archivo zonas.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Integer
    
    NumZonas = GetVar(archivoC, "Config", "Cantidad")
    
    ReDim Zonas(0 To NumZonas)
    For i = 1 To NumZonas
        Zonas(i).Nombre = GetVar(archivoC, "Zona" & CStr(i), "Nombre")
        Zonas(i).mapa = CByte(GetVar(archivoC, "Zona" & CStr(i), "Mapa"))
        Zonas(i).X1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X1"))
        Zonas(i).Y1 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y1"))
        Zonas(i).X2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "X2"))
        Zonas(i).Y2 = CInt(GetVar(archivoC, "Zona" & CStr(i), "Y2"))
        Zonas(i).Segura = val(GetVar(archivoC, "Zona" & CStr(i), "Segura"))
        Zonas(i).Terreno = val(GetVar(archivoC, "Zona" & CStr(i), "Terreno"))
        Zonas(i).InviSinEfecto = val(GetVar(archivoC, "Zona" & CStr(i), "InviSinEfecto"))
        Zonas(i).MagiaSinEfecto = val(GetVar(archivoC, "Zona" & CStr(i), "MagiaSinEfecto"))
        Zonas(i).Restringir = val(GetVar(archivoC, "Zona" & CStr(i), "Restringir"))
        Zonas(i).ResuSinEfecto = val(GetVar(archivoC, "Zona" & CStr(i), "ResuSinEfecto"))
        Zonas(i).Musica1 = val(GetVar(archivoC, "Zona" & CStr(i), "Musica1"))
        Zonas(i).Musica2 = val(GetVar(archivoC, "Zona" & CStr(i), "Musica2"))
        Zonas(i).Musica3 = val(GetVar(archivoC, "Zona" & CStr(i), "Musica3"))
        Zonas(i).Musica4 = val(GetVar(archivoC, "Zona" & CStr(i), "Musica4"))
        Zonas(i).Musica5 = val(GetVar(archivoC, "Zona" & CStr(i), "Musica5"))
        Zonas(i).Acoplar = val(GetVar(archivoC, "Zona" & CStr(i), "Acoplar"))
    Next i
End Sub
Sub LoadArmadurasHerreria()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ArmadurasHerrero.dat", "INIT", "NumArmaduras"))

ReDim Preserve ArmadurasHerrero(1 To N) As Integer

For lc = 1 To N
    ArmadurasHerrero(lc) = val(GetVar(DatPath & "ArmadurasHerrero.dat", "Armadura" & lc, "Index"))
Next lc

End Sub

Sub LoadBalance()
    Dim i As Long
    
    'Modificadores de Clase
    For i = 1 To NUMCLASES
        ModClase(i).Evasion = val(GetVar(DatPath & "Balance.dat", "MODEVASION", ListaClases(i)))
        ModClase(i).AtaqueArmas = val(GetVar(DatPath & "Balance.dat", "MODATAQUEARMAS", ListaClases(i)))
        ModClase(i).AtaqueProyectiles = val(GetVar(DatPath & "Balance.dat", "MODATAQUEPROYECTILES", ListaClases(i)))
        ModClase(i).DañoArmas = val(GetVar(DatPath & "Balance.dat", "MODDAÑOARMAS", ListaClases(i)))
        ModClase(i).DañoProyectiles = val(GetVar(DatPath & "Balance.dat", "MODDAÑOPROYECTILES", ListaClases(i)))
        ModClase(i).DañoWrestling = val(GetVar(DatPath & "Balance.dat", "MODDAÑOWRESTLING", ListaClases(i)))
        ModClase(i).Escudo = val(GetVar(DatPath & "Balance.dat", "MODESCUDO", ListaClases(i)))
    Next i
    
    'Modificadores de Raza
    For i = 1 To NUMRAZAS
        ModRaza(i).Fuerza = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Fuerza"))
        ModRaza(i).Agilidad = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Agilidad"))
        ModRaza(i).Inteligencia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Inteligencia"))
        ModRaza(i).Carisma = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Carisma"))
        ModRaza(i).Constitucion = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Constitucion"))
        ModRaza(i).Armas = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Armas")) / 100
        ModRaza(i).Magia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Magia")) / 100
        ModRaza(i).ReduceMagia = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "ReduceMagia")) / 100
        ModRaza(i).Proyectiles = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "Proyectiles")) / 100
        ModRaza(i).EvitarGolpe = val(GetVar(DatPath & "Balance.dat", "MODRAZA", ListaRazas(i) + "EvitarGolpe")) / 100
    Next i
    
    'Modificadores de Vida
    For i = 1 To NUMCLASES
        ModVida(i) = val(GetVar(DatPath & "Balance.dat", "MODVIDA", ListaClases(i)))
    Next i
    
    'Distribución de Vida
    For i = 1 To 5
        DistribucionEnteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "E" + CStr(i)))
    Next i
    For i = 1 To 4
        DistribucionSemienteraVida(i) = val(GetVar(DatPath & "Balance.dat", "DISTRIBUCION", "S" + CStr(i)))
    Next i
    
    'Extra
    PorcentajeRecuperoMana = val(GetVar(DatPath & "Balance.dat", "EXTRA", "PorcentajeRecuperoMana"))

    'Party
    ExponenteNivelParty = val(GetVar(DatPath & "Balance.dat", "PARTY", "ExponenteNivelParty"))
End Sub

Sub LoadObjCarpintero()

Dim N As Integer, lc As Integer

N = val(GetVar(DatPath & "ObjCarpintero.dat", "INIT", "NumObjs"))

ReDim Preserve ObjCarpintero(1 To N) As Integer

For lc = 1 To N
    ObjCarpintero(lc) = val(GetVar(DatPath & "ObjCarpintero.dat", "Obj" & lc, "Index"))
Next lc

End Sub



Sub LoadOBJData()

'###################################################
'#               ATENCION PELIGRO                  #
'###################################################
'
'¡¡¡¡ NO USAR GetVar PARA LEER DESDE EL OBJ.DAT !!!!
'
'El que ose desafiar esta LEY, se las tendrá que ver
'con migo. Para leer desde el OBJ.DAT se deberá usar
'la nueva clase clsLeerInis.
'
'Alejo
'
'###################################################

'Call LogTarea("Sub LoadOBJData")

'on error GoTo Errhandler

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando base de datos de los objetos."

'*****************************************************************
'Carga la lista de objetos
'*****************************************************************
Dim Object As Integer
Dim Leer As New clsIniReader

Call Leer.Initialize(DatPath & "Obj.dat")

'obtiene el numero de obj
NumObjDatas = val(Leer.GetValue("INIT", "NumObjs"))

frmCargando.cargar.min = 0
frmCargando.cargar.max = NumObjDatas
frmCargando.cargar.value = 0


ReDim Preserve ObjData(1 To NumObjDatas) As ObjData
  
'Llena la lista
For Object = 1 To NumObjDatas
        
    ObjData(Object).Name = Leer.GetValue("OBJ" & Object, "Name")
    
    'Pablo (ToxicWaste) Log de Objetos.
    ObjData(Object).Log = val(Leer.GetValue("OBJ" & Object, "Log"))
    ObjData(Object).NoLog = val(Leer.GetValue("OBJ" & Object, "NoLog"))
    '07/09/07
    
    ObjData(Object).GrhIndex = val(Leer.GetValue("OBJ" & Object, "GrhIndex"))
    If ObjData(Object).GrhIndex = 0 Then
        ObjData(Object).GrhIndex = ObjData(Object).GrhIndex
    End If
    
    ObjData(Object).OBJType = val(Leer.GetValue("OBJ" & Object, "ObjType"))
    
    ObjData(Object).Newbie = val(Leer.GetValue("OBJ" & Object, "Newbie"))
    
    Select Case ObjData(Object).OBJType
        Case eOBJType.otArmadura
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
        
        Case eOBJType.otESCUDO
            ObjData(Object).ShieldAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otCASCO
            ObjData(Object).CascoAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otWeapon
            ObjData(Object).WeaponAnim = val(Leer.GetValue("OBJ" & Object, "Anim"))
            ObjData(Object).Apuñala = val(Leer.GetValue("OBJ" & Object, "Apuñala"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).proyectil = val(Leer.GetValue("OBJ" & Object, "Proyectil"))
            ObjData(Object).Municion = val(Leer.GetValue("OBJ" & Object, "Municiones"))
            ObjData(Object).StaffPower = val(Leer.GetValue("OBJ" & Object, "StaffPower"))
            ObjData(Object).StaffDamageBonus = val(Leer.GetValue("OBJ" & Object, "StaffDamageBonus"))
            ObjData(Object).Refuerzo = val(Leer.GetValue("OBJ" & Object, "Refuerzo"))
            
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otInstrumentos
            ObjData(Object).Snd1 = val(Leer.GetValue("OBJ" & Object, "SND1"))
            ObjData(Object).Snd2 = val(Leer.GetValue("OBJ" & Object, "SND2"))
            ObjData(Object).Snd3 = val(Leer.GetValue("OBJ" & Object, "SND3"))
            'Pablo (ToxicWaste)
            ObjData(Object).Real = val(Leer.GetValue("OBJ" & Object, "Real"))
            ObjData(Object).Caos = val(Leer.GetValue("OBJ" & Object, "Caos"))
        
        Case eOBJType.otMinerales
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
        
        Case eOBJType.otPuertas, eOBJType.otBotellaVacia, eOBJType.otBotellaLlena
            ObjData(Object).IndexAbierta = val(Leer.GetValue("OBJ" & Object, "IndexAbierta"))
            ObjData(Object).IndexCerrada = val(Leer.GetValue("OBJ" & Object, "IndexCerrada"))
            ObjData(Object).IndexCerradaLlave = val(Leer.GetValue("OBJ" & Object, "IndexCerradaLlave"))
        
        Case otPociones
            ObjData(Object).TipoPocion = val(Leer.GetValue("OBJ" & Object, "TipoPocion"))
            ObjData(Object).MaxModificador = val(Leer.GetValue("OBJ" & Object, "MaxModificador"))
            ObjData(Object).MinModificador = val(Leer.GetValue("OBJ" & Object, "MinModificador"))
            ObjData(Object).DuracionEfecto = val(Leer.GetValue("OBJ" & Object, "DuracionEfecto"))
        
        Case eOBJType.otBarcos
            ObjData(Object).MinSkill = val(Leer.GetValue("OBJ" & Object, "MinSkill"))
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
        
        Case eOBJType.otFlechas
            ObjData(Object).MaxHIT = val(Leer.GetValue("OBJ" & Object, "MaxHIT"))
            ObjData(Object).MinHIT = val(Leer.GetValue("OBJ" & Object, "MinHIT"))
            ObjData(Object).Envenena = val(Leer.GetValue("OBJ" & Object, "Envenena"))
            ObjData(Object).Paraliza = val(Leer.GetValue("OBJ" & Object, "Paraliza"))
        Case eOBJType.otAnillo 'Pablo (ToxicWaste)
            ObjData(Object).LingH = val(Leer.GetValue("OBJ" & Object, "LingH"))
            ObjData(Object).LingP = val(Leer.GetValue("OBJ" & Object, "LingP"))
            ObjData(Object).LingO = val(Leer.GetValue("OBJ" & Object, "LingO"))
            ObjData(Object).SkHerreria = val(Leer.GetValue("OBJ" & Object, "SkHerreria"))
            
            
    End Select
    
    ObjData(Object).Ropaje = val(Leer.GetValue("OBJ" & Object, "NumRopaje"))
    ObjData(Object).HechizoIndex = val(Leer.GetValue("OBJ" & Object, "HechizoIndex"))
    
    ObjData(Object).LingoteIndex = val(Leer.GetValue("OBJ" & Object, "LingoteIndex"))
    
    ObjData(Object).MineralIndex = val(Leer.GetValue("OBJ" & Object, "MineralIndex"))
    
    ObjData(Object).MaxHP = val(Leer.GetValue("OBJ" & Object, "MaxHP"))
    ObjData(Object).MinHP = val(Leer.GetValue("OBJ" & Object, "MinHP"))
    
    ObjData(Object).Mujer = val(Leer.GetValue("OBJ" & Object, "Mujer"))
    ObjData(Object).Hombre = val(Leer.GetValue("OBJ" & Object, "Hombre"))
    
    ObjData(Object).MinHam = val(Leer.GetValue("OBJ" & Object, "MinHam"))
    ObjData(Object).MinSed = val(Leer.GetValue("OBJ" & Object, "MinAgu"))
    
    ObjData(Object).MinDef = val(Leer.GetValue("OBJ" & Object, "MINDEF"))
    ObjData(Object).MaxDef = val(Leer.GetValue("OBJ" & Object, "MAXDEF"))
    ObjData(Object).def = (ObjData(Object).MinDef + ObjData(Object).MaxDef) / 2
    
    ObjData(Object).RazaEnana = val(Leer.GetValue("OBJ" & Object, "RazaEnana"))
    ObjData(Object).RazaDrow = val(Leer.GetValue("OBJ" & Object, "RazaDrow"))
    ObjData(Object).RazaElfa = val(Leer.GetValue("OBJ" & Object, "RazaElfa"))
    ObjData(Object).RazaGnoma = val(Leer.GetValue("OBJ" & Object, "RazaGnoma"))
    ObjData(Object).RazaHumana = val(Leer.GetValue("OBJ" & Object, "RazaHumana"))
    
    ObjData(Object).Valor = val(Leer.GetValue("OBJ" & Object, "Valor"))
    
    ObjData(Object).Crucial = val(Leer.GetValue("OBJ" & Object, "Crucial"))
    
    ObjData(Object).Cerrada = val(Leer.GetValue("OBJ" & Object, "abierta"))
    If ObjData(Object).Cerrada = 1 Then
        ObjData(Object).Llave = val(Leer.GetValue("OBJ" & Object, "Llave"))
        ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    End If
    
    'Puertas y llaves
    ObjData(Object).clave = val(Leer.GetValue("OBJ" & Object, "Clave"))
    
    ObjData(Object).texto = Leer.GetValue("OBJ" & Object, "Texto")
    ObjData(Object).GrhSecundario = val(Leer.GetValue("OBJ" & Object, "VGrande"))
    
    ObjData(Object).Agarrable = val(Leer.GetValue("OBJ" & Object, "Agarrable"))
    
    
    'CHECK: !!! Esto es provisorio hasta que los de Dateo cambien los valores de string a numerico
    Dim i As Integer
    Dim N As Integer
    Dim S As String
    For i = 1 To NUMCLASES
        S = UCase$(Leer.GetValue("OBJ" & Object, "CP" & i))
        N = 1
        Do While LenB(S) > 0 And UCase$(ListaClases(N)) <> S
            N = N + 1
        Loop
        ObjData(Object).ClaseProhibida(i) = IIf(LenB(S) > 0, N, 0)
    Next i
    
    ObjData(Object).DefensaMagicaMax = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMax"))
    ObjData(Object).DefensaMagicaMin = val(Leer.GetValue("OBJ" & Object, "DefensaMagicaMin"))
    
    ObjData(Object).SkCarpinteria = val(Leer.GetValue("OBJ" & Object, "SkCarpinteria"))
    
    If ObjData(Object).SkCarpinteria > 0 Then
        ObjData(Object).Madera = val(Leer.GetValue("OBJ" & Object, "Madera"))
        ObjData(Object).MaderaElfica = val(Leer.GetValue("OBJ" & Object, "MaderaElfica"))
    End If
    
    'Bebidas
    ObjData(Object).MinSta = val(Leer.GetValue("OBJ" & Object, "MinST"))
    
    ObjData(Object).NoSeCae = val(Leer.GetValue("OBJ" & Object, "NoSeCae"))
    
    ObjData(Object).Upgrade = val(Leer.GetValue("OBJ" & Object, "Upgrade"))
    
    ObjData(Object).Efecto = val(Leer.GetValue("OBJ" & Object, "Efecto"))
    
    If Object = 414 Or Object = 415 Or Object = 416 Or Object = 1067 Then
        ObjData(Object).Valor = ObjData(Object).Valor * MultiplicadorORO
    End If
    
    frmCargando.cargar.value = frmCargando.cargar.value + 1
Next Object

Set Leer = Nothing

Exit Sub

Errhandler:
    MsgBox "error cargando objetos " & Err.Number & ": " & Err.Description


End Sub

Sub LoadUserStats(ByVal UserIndex As Integer, ByRef UserFile As clsMySQLRecordSet)

Dim LoopC As Long

For LoopC = 1 To NUMATRIBUTOS
  UserList(UserIndex).Stats.UserAtributos(LoopC) = CInt(UserFile("AT" & LoopC))
  UserList(UserIndex).Stats.UserAtributosBackUP(LoopC) = UserList(UserIndex).Stats.UserAtributos(LoopC)
Next LoopC

For LoopC = 1 To NUMSKILLS
  UserList(UserIndex).Stats.UserSkills(LoopC) = CInt(UserFile("SK" & LoopC))
Next LoopC

For LoopC = 1 To MAXUSERHECHIZOS
  UserList(UserIndex).Stats.UserHechizos(LoopC) = CInt(UserFile("H" & LoopC))
Next LoopC

UserList(UserIndex).Stats.GLD = CLng(UserFile("GLD"))
UserList(UserIndex).Stats.Banco = CLng(UserFile("Banco"))

UserList(UserIndex).Stats.MaxHP = CInt(UserFile("MaxHP"))
UserList(UserIndex).Stats.MinHP = CInt(UserFile("MinHP"))

UserList(UserIndex).Stats.MinSta = CInt(UserFile("MinSta"))
UserList(UserIndex).Stats.MaxSta = CInt(UserFile("MaxSta"))

UserList(UserIndex).Stats.MaxMAN = CInt(UserFile("MaxMAN"))
UserList(UserIndex).Stats.MinMAN = CInt(UserFile("MinMAN"))

UserList(UserIndex).Stats.MaxHIT = CInt(UserFile("MaxHIT"))
UserList(UserIndex).Stats.MinHIT = CInt(UserFile("MinHIT"))

UserList(UserIndex).Stats.MaxAGU = CByte(UserFile("MaxAGU"))
UserList(UserIndex).Stats.MinAGU = CByte(UserFile("MinAGU"))

UserList(UserIndex).Stats.MaxHam = CByte(UserFile("MaxHam"))
UserList(UserIndex).Stats.MinHam = CByte(UserFile("MinHam"))

UserList(UserIndex).Stats.SkillPts = CInt(UserFile("SkillPtsLibres"))

UserList(UserIndex).Stats.Exp = CDbl(UserFile("Exp"))
UserList(UserIndex).Stats.ELU = CLng(UserFile("ELU"))
UserList(UserIndex).Stats.ELV = CByte(UserFile("ELV"))


UserList(UserIndex).Stats.UsuariosMatados = CLng(UserFile("UserMuertes"))
UserList(UserIndex).Stats.NPCsMuertos = CInt(UserFile("NpcsMuertes"))

If CByte(UserFile("PerteneceReal")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.RoyalCouncil

If CByte(UserFile("PerteneceCaos")) Then _
    UserList(UserIndex).flags.Privilegios = UserList(UserIndex).flags.Privilegios Or PlayerType.ChaosCouncil

End Sub

Sub LoadUserReputacion(ByVal UserIndex As Integer, ByRef UserFile As clsMySQLRecordSet)

UserList(UserIndex).Reputacion.AsesinoRep = val(UserFile("Rep_Asesino"))
UserList(UserIndex).Reputacion.BandidoRep = val(UserFile("Rep_Bandido"))
UserList(UserIndex).Reputacion.BurguesRep = val(UserFile("Rep_Burguesia"))
UserList(UserIndex).Reputacion.LadronesRep = val(UserFile("Rep_Ladrones"))
UserList(UserIndex).Reputacion.NobleRep = val(UserFile("Rep_Nobles"))
UserList(UserIndex).Reputacion.PlebeRep = val(UserFile("Rep_Plebe"))
UserList(UserIndex).Reputacion.Promedio = val(UserFile("Rep_Promedio"))

End Sub

Sub LoadUserInit(ByVal UserIndex As Integer, ByRef UserFile As clsMySQLRecordSet)
'*************************************************
'Author: Unknown
'Last modified: 19/11/2006
'Loads the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'23/01/2007 Pablo (ToxicWaste) - Quito CriminalesMatados de Stats porque era redundante.
'*************************************************
Dim LoopC As Long
Dim ln As String
UserList(UserIndex).MySQLId = CLng(UserFile("Id"))
UserList(UserIndex).MySQLIdCuenta = CLng(UserFile("IdAccount"))
UserList(UserIndex).Faccion.ArmadaReal = CByte(UserFile("EjercitoReal"))
UserList(UserIndex).Faccion.FuerzasCaos = CByte(UserFile("EjercitoCaos"))
UserList(UserIndex).Faccion.CiudadanosMatados = CLng(UserFile("CiudMatados"))
UserList(UserIndex).Faccion.CriminalesMatados = CLng(UserFile("CrimMatados"))
UserList(UserIndex).Faccion.RecibioArmaduraCaos = CByte(UserFile("rArCaos"))
UserList(UserIndex).Faccion.RecibioArmaduraReal = CByte(UserFile("rArReal"))
UserList(UserIndex).Faccion.RecibioExpInicialCaos = CByte(UserFile("rExCaos"))
UserList(UserIndex).Faccion.RecibioExpInicialReal = CByte(UserFile("rExReal"))
UserList(UserIndex).Faccion.RecompensasCaos = CLng(UserFile("recCaos"))
UserList(UserIndex).Faccion.RecompensasReal = CLng(UserFile("recReal"))
UserList(UserIndex).Faccion.Reenlistadas = CByte(UserFile("Reenlistadas"))
UserList(UserIndex).Faccion.NivelIngreso = CInt(UserFile("NivelIngreso"))
If UserFile("FechaIngreso") <> "" Then
    UserList(UserIndex).Faccion.FechaIngreso = UserFile("FechaIngreso")
Else
    UserList(UserIndex).Faccion.FechaIngreso = "20000101"
End If
UserList(UserIndex).Faccion.MatadosIngreso = CInt(UserFile("MatadosIngreso"))
UserList(UserIndex).Faccion.NextRecompensa = CInt(UserFile("NextRecompensa"))

UserList(UserIndex).flags.Muerto = CByte(UserFile("Muerto"))
UserList(UserIndex).flags.Escondido = CByte(UserFile("Escondido"))

UserList(UserIndex).flags.Hambre = CByte(UserFile("Hambre"))
UserList(UserIndex).flags.Sed = CByte(UserFile("Sed"))
UserList(UserIndex).flags.Desnudo = CByte(UserFile("Desnudo"))
UserList(UserIndex).flags.Navegando = CByte(UserFile("Navegando"))
UserList(UserIndex).flags.Envenenado = CByte(UserFile("Envenenado"))
UserList(UserIndex).flags.Paralizado = CByte(UserFile("Paralizado"))
If UserList(UserIndex).flags.Paralizado = 1 Then
    UserList(UserIndex).Counters.Paralisis = IntervaloParalizado
End If


UserList(UserIndex).Counters.Pena = CLng(UserFile("Pena"))

UserList(UserIndex).email = UserFile("Email")

UserList(UserIndex).genero = UserFile("Genero")
UserList(UserIndex).clase = UserFile("Clase")
UserList(UserIndex).raza = UserFile("Raza")
UserList(UserIndex).Hogar = UserFile("Hogar")
UserList(UserIndex).Char.heading = CInt(UserFile("heading"))


UserList(UserIndex).OrigChar.Head = CInt(UserFile("Head"))
UserList(UserIndex).OrigChar.Body = CInt(UserFile("Body"))
UserList(UserIndex).OrigChar.WeaponAnim = CInt(UserFile("Arma"))
UserList(UserIndex).OrigChar.ShieldAnim = CInt(UserFile("Escudo"))
UserList(UserIndex).OrigChar.CascoAnim = CInt(UserFile("Casco"))

#If ConUpTime Then
    UserList(UserIndex).UpTime = CLng(UserFile("UpTime"))
#End If

UserList(UserIndex).OrigChar.heading = eHeading.SOUTH

If UserList(UserIndex).flags.Muerto = 0 Then
    UserList(UserIndex).Char = UserList(UserIndex).OrigChar
Else
    If UserFile("Rep_Promedio") < 0 Then
        UserList(UserIndex).Char.Body = iCuerpoMuertoCrimi
        UserList(UserIndex).Char.Head = iCabezaMuertoCrimi
    Else
        UserList(UserIndex).Char.Body = iCuerpoMuerto
        UserList(UserIndex).Char.Head = iCabezaMuerto
    End If
    UserList(UserIndex).Char.WeaponAnim = NingunArma
    UserList(UserIndex).Char.ShieldAnim = NingunEscudo
    UserList(UserIndex).Char.CascoAnim = NingunCasco
End If


UserList(UserIndex).desc = UserFile("Descripcion")

UserList(UserIndex).Pos.map = CInt(UserFile("Map"))
UserList(UserIndex).Pos.X = CInt(UserFile("X"))
UserList(UserIndex).Pos.Y = CInt(UserFile("Y"))

CheckZona (UserIndex)

UserList(UserIndex).Invent.NroItems = CInt(UserFile("InvCantidadItems"))

'[KEVIN]--------------------------------------------------------------------
'***********************************************************************************
UserList(UserIndex).BancoInvent.NroItems = CInt(UserFile("BanCantidadItems"))
'Lista de objetos del banco
For LoopC = 1 To MAX_BANCOINVENTORY_SLOTS
    UserList(UserIndex).BancoInvent.Object(LoopC).ObjIndex = CInt(UserFile("BanObj" & LoopC))
    UserList(UserIndex).BancoInvent.Object(LoopC).Amount = CInt(UserFile("BanCant" & LoopC))
Next LoopC
'------------------------------------------------------------------------------------
'[/KEVIN]*****************************************************************************


'Lista de objetos
For LoopC = 1 To MAX_INVENTORY_SLOTS
    UserList(UserIndex).Invent.Object(LoopC).ObjIndex = CInt(UserFile("InvObj" & LoopC))
    UserList(UserIndex).Invent.Object(LoopC).Amount = CInt(UserFile("InvCant" & LoopC))
    UserList(UserIndex).Invent.Object(LoopC).Equipped = CByte(UserFile("InvEqp" & LoopC))
Next LoopC

'Obtiene el indice-objeto del arma
UserList(UserIndex).Invent.WeaponEqpSlot = CByte(UserFile("WeaponEqpSlot"))
If UserList(UserIndex).Invent.WeaponEqpSlot > 0 Then
    UserList(UserIndex).Invent.WeaponEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.WeaponEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del armadura
UserList(UserIndex).Invent.ArmourEqpSlot = CByte(UserFile("ArmourEqpSlot"))
If UserList(UserIndex).Invent.ArmourEqpSlot > 0 Then
    UserList(UserIndex).Invent.ArmourEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.ArmourEqpSlot).ObjIndex
    UserList(UserIndex).flags.Desnudo = 0
Else
    UserList(UserIndex).flags.Desnudo = 1
End If

'Obtiene el indice-objeto del escudo
UserList(UserIndex).Invent.EscudoEqpSlot = CByte(UserFile("EscudoEqpSlot"))
If UserList(UserIndex).Invent.EscudoEqpSlot > 0 Then
    UserList(UserIndex).Invent.EscudoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.EscudoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto del casco
UserList(UserIndex).Invent.CascoEqpSlot = CByte(UserFile("CascoEqpSlot"))
If UserList(UserIndex).Invent.CascoEqpSlot > 0 Then
    UserList(UserIndex).Invent.CascoEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.CascoEqpSlot).ObjIndex
End If

'Obtiene el indice-objeto barco
UserList(UserIndex).Invent.BarcoSlot = CByte(UserFile("BarcoSlot"))
If UserList(UserIndex).Invent.BarcoSlot > 0 Then
    UserList(UserIndex).Invent.BarcoObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.BarcoSlot).ObjIndex
End If

'Obtiene el indice-objeto municion
UserList(UserIndex).Invent.MunicionEqpSlot = CByte(UserFile("MunicionSlot"))
If UserList(UserIndex).Invent.MunicionEqpSlot > 0 Then
    UserList(UserIndex).Invent.MunicionEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.MunicionEqpSlot).ObjIndex
End If

'[Alejo]
'Obtiene el indice-objeto anilo
UserList(UserIndex).Invent.AnilloEqpSlot = CByte(UserFile("AnilloSlot"))
If UserList(UserIndex).Invent.AnilloEqpSlot > 0 Then
    UserList(UserIndex).Invent.AnilloEqpObjIndex = UserList(UserIndex).Invent.Object(UserList(UserIndex).Invent.AnilloEqpSlot).ObjIndex
End If

UserList(UserIndex).NroMascotas = CInt(UserFile("NroMascotas"))
Dim NpcIndex As Integer
For LoopC = 1 To MAXMASCOTAS
    UserList(UserIndex).MascotasType(LoopC) = val(UserFile("Masc" & LoopC))
Next LoopC

UserList(UserIndex).GuildIndex = CInt(UserFile("GuildIndex"))

End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String, Optional EmptySpaces As Long = 1024) As String

Dim sSpaces As String ' This will hold the input that the program will retrieve
Dim szReturn As String ' This will be the defaul value if the string is not found
  
szReturn = vbNullString
  
sSpaces = Space$(EmptySpaces) ' This tells the computer how long the longest string can be
  
  
GetPrivateProfileString Main, Var, szReturn, sSpaces, EmptySpaces, file
  
GetVar = RTrim$(sSpaces)
GetVar = left$(GetVar, Len(GetVar) - 1)
  
End Function

Sub CargarBackUp()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

On Error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))
    
    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    

    For map = 1 To NumMaps
        tFileName = App.Path & "\WorldBackup\" & "Mapa" & map
        
        Call CargarBak(map, tFileName)
        
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)
 
End Sub

Sub LoadMapData()

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando mapas..."

Dim map As Integer
Dim TempInt As Integer
Dim tFileName As String
Dim npcfile As String

'on error GoTo man
    
    NumMaps = val(GetVar(DatPath & "Map.dat", "INIT", "NumMaps"))

    frmCargando.cargar.min = 0
    frmCargando.cargar.max = NumMaps
    frmCargando.cargar.value = 0
    
    MapPath = GetVar(DatPath & "Map.dat", "INIT", "MapPath")
    
    
    ReDim MapData(1 To NumMaps, XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
      
    For map = 1 To NumMaps
        
        tFileName = App.Path & MapPath & "SMapa" & map
        Call CargarMapa(map, tFileName)
        
        DoEvents
    Next map

Exit Sub

man:
    MsgBox ("Error durante la carga de mapas, el mapa " & map & " contiene errores")
    Call LogError(Date & " " & Err.Description & " " & Err.HelpContext & " " & Err.HelpFile & " " & Err.Source)

End Sub

Public Sub CargarMapa(ByVal map As Long, ByVal MAPFl As String)
On Error GoTo errh
    Dim FreeFileMap As Long
    Dim FreeFileInf As Long
    Dim Y As Long
    Dim X As Long
    Dim ByFlags As Integer
    Dim npcfile As String
    Dim TempInt As Integer
    Dim TempByte As Byte
    Dim tmpInt As Integer
      
    FreeFileMap = FreeFile
    
    Open MAPFl & ".map" For Binary As #FreeFileMap
    'Seek FreeFileMap, 1
    
    FreeFileInf = FreeFile
    
    frmCargando.cargar.max = YMaxMapSize
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
            
            '.dat file
            Get FreeFileMap, , ByFlags

            If ByFlags And 1 Then
                MapData(map, X, Y).Blocked = 1
            End If
            
            Get FreeFileMap, , MapData(map, X, Y).Graphic(1)
            
            'Layer 2 used?
            If ByFlags And 2 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(2)
            
            'Layer 3 used?
            If ByFlags And 4 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(3)
            
            'Layer 4 used?
            If ByFlags And 8 Then Get FreeFileMap, , MapData(map, X, Y).Graphic(4)
            
            'Trigger used?
            If ByFlags And 16 Then
                'Enums are 4 byte long in VB, so we make sure we only read 2
                Get FreeFileMap, , TempByte
                MapData(map, X, Y).Trigger = TempByte
            End If
            
            If ByFlags And 32 Then
                'Get and make NPC
                Get FreeFileMap, , MapData(map, X, Y).NpcIndex
                
                If MapData(map, X, Y).NpcIndex > 0 Then
                    npcfile = DatPath & "NPCs.dat"

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(map, X, Y).NpcIndex, "PosOrig")) = 1 Then
                        tmpInt = OpenNPC(MapData(map, X, Y).NpcIndex)
                        MapData(map, X, Y).NpcIndex = tmpInt
                        Npclist(tmpInt).Orig.map = map
                        Npclist(tmpInt).Orig.X = X
                        Npclist(tmpInt).Orig.Y = Y
                        Call CambiarOrigHeading(tmpInt, MapData(map, X, Y).Trigger)
                        
                    Else
                        tmpInt = OpenNPC(MapData(map, X, Y).NpcIndex)
                        MapData(map, X, Y).NpcIndex = tmpInt
                    End If
                    
                    Npclist(tmpInt).Pos.map = map
                    Npclist(tmpInt).Pos.X = X
                    Npclist(tmpInt).Pos.Y = Y
                    
                    Call CheckZonaNPC(tmpInt)
                    
                    Call MakeNPCChar(True, 0, tmpInt, map, X, Y)
                End If
            End If
            
            If ByFlags And 64 Then
                'Get and make Object
                Get FreeFileMap, , MapData(map, X, Y).ObjInfo.ObjIndex
                Get FreeFileMap, , MapData(map, X, Y).ObjInfo.Amount
            End If
            If ByFlags And 128 Then
                'Get and make Object
                Get FreeFileMap, , MapData(map, X, Y).TileExit.map
                Get FreeFileMap, , MapData(map, X, Y).TileExit.X
                Get FreeFileMap, , MapData(map, X, Y).TileExit.Y
            End If
            
            If MapData(map, X, Y).Graphic(3) < 0 Then
                Get FreeFileMap, , TempByte
            End If
        Next X
        frmCargando.cargar.value = Y
    Next Y
    
    
    Close FreeFileMap


Exit Sub

errh:
    Call LogError("Error cargando mapa: " & map & " - Pos: " & X & "," & Y & "." & Err.Description)
End Sub

Sub LoadSini()

Dim Temporal As Long

If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando info de inicio del server."

BootDelBackUp = val(GetVar(IniPath & "Server.ini", "INIT", "IniciarDesdeBackUp"))

'Misc
#If SeguridadAlkon Then

Call Security.SetServerIp(GetVar(IniPath & "Server.ini", "INIT", "ServerIp"))

#End If


Puerto = val(GetVar(IniPath & "Server.ini", "INIT", "StartPort"))
HideMe = val(GetVar(IniPath & "Server.ini", "INIT", "Hide"))
AllowMultiLogins = val(GetVar(IniPath & "Server.ini", "INIT", "AllowMultiLogins"))
IdleLimit = val(GetVar(IniPath & "Server.ini", "INIT", "IdleLimit"))
'Lee la version correcta del cliente
ULTIMAVERSION = GetVar(IniPath & "Server.ini", "INIT", "Version")

PuedeCrearPersonajes = val(GetVar(IniPath & "Server.ini", "INIT", "PuedeCrearPersonajes"))
ServerSoloGMs = val(GetVar(IniPath & "Server.ini", "init", "ServerSoloGMs"))

MultiplicadorEXP = CDbl(GetVar(IniPath & "Server.ini", "init", "MultiplicadorEXP"))
MultiplicadorORO = CDbl(GetVar(IniPath & "Server.ini", "init", "MultiplicadorORO"))

ArmaduraImperial1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial1"))
ArmaduraImperial2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial2"))
ArmaduraImperial3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraImperial3"))
TunicaMagoImperial = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperial"))
TunicaMagoImperialEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoImperialEnanos"))
ArmaduraCaos1 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos1"))
ArmaduraCaos2 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos2"))
ArmaduraCaos3 = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraCaos3"))
TunicaMagoCaos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaos"))
TunicaMagoCaosEnanos = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaMagoCaosEnanos"))

VestimentaImperialHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialHumano"))
VestimentaImperialEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaImperialEnano"))
TunicaConspicuaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaHumano"))
TunicaConspicuaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaConspicuaEnano"))
ArmaduraNobilisimaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaHumano"))
ArmaduraNobilisimaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraNobilisimaEnano"))
ArmaduraGranSacerdote = val(GetVar(IniPath & "Server.ini", "INIT", "ArmaduraGranSacerdote"))

VestimentaLegionHumano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionHumano"))
VestimentaLegionEnano = val(GetVar(IniPath & "Server.ini", "INIT", "VestimentaLegionEnano"))
TunicaLobregaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaHumano"))
TunicaLobregaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaLobregaEnano"))
TunicaEgregiaHumano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaHumano"))
TunicaEgregiaEnano = val(GetVar(IniPath & "Server.ini", "INIT", "TunicaEgregiaEnano"))
SacerdoteDemoniaco = val(GetVar(IniPath & "Server.ini", "INIT", "SacerdoteDemoniaco"))

MAPA_PRETORIANO = val(GetVar(IniPath & "Server.ini", "INIT", "MapaPretoriano"))

EnTesting = val(GetVar(IniPath & "Server.ini", "INIT", "Testing"))

'Intervalos
SanaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloSinDescansar"))
FrmInterv.txtSanaIntervaloSinDescansar.Text = SanaIntervaloSinDescansar

StaminaIntervaloLluvia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloLluvia"))
FrmInterv.txtStaminaIntervaloLluvia.Text = StaminaIntervaloLluvia

StaminaIntervaloSinDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloSinDescansar"))
FrmInterv.txtStaminaIntervaloSinDescansar.Text = StaminaIntervaloSinDescansar

SanaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "SanaIntervaloDescansar"))
FrmInterv.txtSanaIntervaloDescansar.Text = SanaIntervaloDescansar

StaminaIntervaloDescansar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "StaminaIntervaloDescansar"))
FrmInterv.txtStaminaIntervaloDescansar.Text = StaminaIntervaloDescansar

IntervaloSed = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloSed"))
FrmInterv.txtIntervaloSed.Text = IntervaloSed

IntervaloHambre = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloHambre"))
FrmInterv.txtIntervaloHambre.Text = IntervaloHambre

IntervaloVeneno = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloVeneno"))
FrmInterv.txtIntervaloVeneno.Text = IntervaloVeneno

IntervaloParalizado = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParalizado"))
FrmInterv.txtIntervaloParalizado.Text = IntervaloParalizado

IntervaloInvisible = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvisible"))
FrmInterv.txtIntervaloInvisible.Text = IntervaloInvisible

IntervaloFrio = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFrio"))
FrmInterv.txtIntervaloFrio.Text = IntervaloFrio

IntervaloWavFx = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWAVFX"))
FrmInterv.txtIntervaloWAVFX.Text = IntervaloWavFx

IntervaloInvocacion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloInvocacion"))
FrmInterv.txtInvocacion.Text = IntervaloInvocacion

IntervaloParaConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloParaConexion"))
FrmInterv.txtIntervaloParaConexion.Text = IntervaloParaConexion

'&&&&&&&&&&&&&&&&&&&&& TIMERS &&&&&&&&&&&&&&&&&&&&&&&


IntervaloUserPuedeCastear = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloLanzaHechizo"))
FrmInterv.txtIntervaloLanzaHechizo.Text = IntervaloUserPuedeCastear

frmMain.TIMER_AI.interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcAI"))
FrmInterv.txtAI.Text = frmMain.TIMER_AI.interval

frmMain.npcataca.interval = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloNpcPuedeAtacar"))
FrmInterv.txtNPCPuedeAtacar.Text = frmMain.npcataca.interval

IntervaloUserPuedeTrabajar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloTrabajo"))
FrmInterv.txtTrabajo.Text = IntervaloUserPuedeTrabajar

IntervaloUserPuedeAtacar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeAtacar"))
FrmInterv.txtPuedeAtacar.Text = IntervaloUserPuedeAtacar

'TODO : Agregar estos intervalos al form!!!
IntervaloMagiaGolpe = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloMagiaGolpe"))
IntervaloGolpeMagia = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeMagia"))
IntervaloGolpeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloGolpeUsar"))


MinutosWs = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloWS"))
If MinutosWs < 60 Then MinutosWs = 180

IntervaloCerrarConexion = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloCerrarConexion"))
IntervaloUserPuedeUsar = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloUserPuedeUsar"))
IntervaloFlechasCazadores = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloFlechasCazadores"))

IntervaloOculto = val(GetVar(IniPath & "Server.ini", "INTERVALOS", "IntervaloOculto"))

IntervaloPuedeSerAtacado = 5000 ' Cargar desde balance.dat
IntervaloAtacable = 60000 ' Cargar desde balance.dat
IntervaloOwnedNpc = 18000 ' Cargar desde balance.dat

'&&&&&&&&&&&&&&&&&&&&& FIN TIMERS &&&&&&&&&&&&&&&&&&&&&&&


'&&&&&&&&&&&&&&&&&&&&& MYSQL &&&&&&&&&&&&&&&&&&&&&&&
MySQL_Host = GetVar(IniPath & "Server.ini", "MYSQL", "Host")
MySQL_User = GetVar(IniPath & "Server.ini", "MYSQL", "User")
MySQL_Pass = GetVar(IniPath & "Server.ini", "MYSQL", "Pass")
MySQL_DB = GetVar(IniPath & "Server.ini", "MYSQL", "DB")
'&&&&&&&&&&&&&&&&&&&&& FIN MYSQL &&&&&&&&&&&&&&&&&&&&&&&
  
recordusuarios = val(GetVar(IniPath & "Server.ini", "INIT", "Record"))
  
'Max users
Temporal = val(GetVar(IniPath & "Server.ini", "INIT", "MaxUsers"))
If MaxUsers = 0 Then
    MaxUsers = Temporal
    ReDim UserList(1 To MaxUsers) As user
End If

'&&&&&&&&&&&&&&&&&&&&& BALANCE &&&&&&&&&&&&&&&&&&&&&&&
'Se agregó en LoadBalance y en el Balance.dat
'PorcentajeRecuperoMana = val(GetVar(IniPath & "Server.ini", "BALANCE", "PorcentajeRecuperoMana"))

''&&&&&&&&&&&&&&&&&&&&& FIN BALANCE &&&&&&&&&&&&&&&&&&&&&&&
Call Statistics.Initialize

Nix.map = GetVar(DatPath & "Ciudades.dat", "NIX", "Mapa")
Nix.X = GetVar(DatPath & "Ciudades.dat", "NIX", "X")
Nix.Y = GetVar(DatPath & "Ciudades.dat", "NIX", "Y")

Ullathorpe.map = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Mapa")
Ullathorpe.X = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "X")
Ullathorpe.Y = GetVar(DatPath & "Ciudades.dat", "Ullathorpe", "Y")

Banderbill.map = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Mapa")
Banderbill.X = GetVar(DatPath & "Ciudades.dat", "Banderbill", "X")
Banderbill.Y = GetVar(DatPath & "Ciudades.dat", "Banderbill", "Y")

Arkhein.map = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Mapa")
Arkhein.X = GetVar(DatPath & "Ciudades.dat", "Arkhein", "X")
Arkhein.Y = GetVar(DatPath & "Ciudades.dat", "Arkhein", "Y")

Arghal.map = GetVar(DatPath & "Ciudades.dat", "Arghal", "Mapa")
Arghal.X = GetVar(DatPath & "Ciudades.dat", "Arghal", "X")
Arghal.Y = GetVar(DatPath & "Ciudades.dat", "Arghal", "Y")

Lindos.map = GetVar(DatPath & "Ciudades.dat", "Lindos", "Mapa")
Lindos.X = GetVar(DatPath & "Ciudades.dat", "Lindos", "X")
Lindos.Y = GetVar(DatPath & "Ciudades.dat", "Lindos", "Y")

Dim i As Integer
For i = 1 To eCiudad.cLastCity - 1
    Hogares(i).map = GetVar(DatPath & "Hogares.dat", "Hogar" & i, "Mapa")
    Hogares(i).X = GetVar(DatPath & "Hogares.dat", "Hogar" & i, "X")
    Hogares(i).Y = GetVar(DatPath & "Hogares.dat", "Hogar" & i, "Y")
Next i

Call MD5sCarga

Call ConsultaPopular.LoadData

#If SeguridadAlkon Then
Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If

End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal value As String)
'*****************************************************************
'Escribe VAR en un archivo
'*****************************************************************

writeprivateprofilestring Main, Var, value, file
    
End Sub

Sub SaveUser(ByVal UserIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 23/01/2007
'Saves the Users records
'23/01/2007 Pablo (ToxicWaste) - Agrego NivelIngreso, FechaIngreso, MatadosIngreso y NextRecompensa.
'*************************************************

On Error GoTo Errhandler

Dim OldUserHead As Long


'ESTO TIENE QUE EVITAR ESE BUGAZO QUE NO SE POR QUE GRABA USUARIOS NULOS
'clase=0 es el error, porq el enum empieza de 1!!
If UserList(UserIndex).clase = 0 Or UserList(UserIndex).Stats.ELV = 0 Then
    Call LogCriticEvent("Estoy intentantdo guardar un usuario nulo de nombre: " & UserList(UserIndex).Name)
    Exit Sub
End If


If UserList(UserIndex).flags.Mimetizado = 1 Then
    UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
    UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
    UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
    UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
    UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
    UserList(UserIndex).Counters.Mimetismo = 0
    UserList(UserIndex).flags.Mimetizado = 0
End If



'If FileExist(UserFile, vbNormal) Then
    If UserList(UserIndex).flags.Muerto = 1 Then
        'OldUserHead = UserList(UserIndex).Char.Head
        'UserList(UserIndex).Char.Head = GetVar(UserFile, "INIT", "Head")
    End If
'       Kill UserFile
'End If

Dim LoopC As Integer

Dim Query As String

Query = "Muerto=" & CStr(UserList(UserIndex).flags.Muerto) & ", "
Query = Query & "Escondido=" & CStr(UserList(UserIndex).flags.Escondido) & ", "
Query = Query & "Hambre=" & CStr(UserList(UserIndex).flags.Hambre) & ", "
Query = Query & "Sed=" & CStr(UserList(UserIndex).flags.Sed) & ", "
Query = Query & "Desnudo=" & CStr(UserList(UserIndex).flags.Desnudo) & ", "
Query = Query & "Ban=" & CStr(UserList(UserIndex).flags.Ban) & ", "
Query = Query & "Navegando=" & CStr(UserList(UserIndex).flags.Navegando) & ", "
Query = Query & "Envenenado=" & CStr(UserList(UserIndex).flags.Envenenado) & ", "
Query = Query & "Paralizado=" & CStr(UserList(UserIndex).flags.Paralizado) & ", "

Query = Query & "PerteneceReal=" & IIf(UserList(UserIndex).flags.Privilegios And PlayerType.RoyalCouncil, "1", "0") & ", "
Query = Query & "PerteneceCaos=" & IIf(UserList(UserIndex).flags.Privilegios And PlayerType.ChaosCouncil, "1", "0") & ", "


Query = Query & "Pena=" & CStr(UserList(UserIndex).Counters.Pena) & ", "

Query = Query & "EjercitoReal=" & CStr(UserList(UserIndex).Faccion.ArmadaReal) & ", "
Query = Query & "EjercitoCaos=" & CStr(UserList(UserIndex).Faccion.FuerzasCaos) & ", "
Query = Query & "CiudMatados=" & CStr(UserList(UserIndex).Faccion.CiudadanosMatados) & ", "
Query = Query & "CrimMatados=" & CStr(UserList(UserIndex).Faccion.CriminalesMatados) & ", "
Query = Query & "rArCaos=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraCaos) & ", "
Query = Query & "rArReal=" & CStr(UserList(UserIndex).Faccion.RecibioArmaduraReal) & ", "
Query = Query & "rExCaos=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialCaos) & ", "
Query = Query & "rExReal=" & CStr(UserList(UserIndex).Faccion.RecibioExpInicialReal) & ", "
Query = Query & "recCaos=" & CStr(UserList(UserIndex).Faccion.RecompensasCaos) & ", "
Query = Query & "recReal=" & CStr(UserList(UserIndex).Faccion.RecompensasReal) & ", "
Query = Query & "Reenlistadas=" & CStr(UserList(UserIndex).Faccion.Reenlistadas) & ", "
Query = Query & "NivelIngreso=" & CStr(UserList(UserIndex).Faccion.NivelIngreso) & ", "
Query = Query & "FechaIngreso=20000101, "
Query = Query & "MatadosIngreso=" & CStr(UserList(UserIndex).Faccion.MatadosIngreso) & ", "
Query = Query & "NextRecompensa=" & CStr(UserList(UserIndex).Faccion.NextRecompensa) & ", "


'¿Fueron modificados los atributos del usuario?
If Not UserList(UserIndex).flags.TomoPocion Then
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        Query = Query & "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributos(LoopC)) & ", "
    Next
Else
    For LoopC = 1 To UBound(UserList(UserIndex).Stats.UserAtributos)
        'UserList(UserIndex).Stats.UserAtributos(LoopC) = UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)
        Query = Query & "AT" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserAtributosBackUP(LoopC)) & ", "
    Next
End If

For LoopC = 1 To NUMSKILLS
    Query = Query & "SK" & LoopC & "=" & CStr(UserList(UserIndex).Stats.UserSkills(LoopC)) & ", "
Next


Query = Query & "Email='" & UserList(UserIndex).email & "', "

Query = Query & "Genero=" & UserList(UserIndex).genero & ", "
Query = Query & "Raza=" & UserList(UserIndex).raza & ", "
Query = Query & "Hogar=" & UserList(UserIndex).Hogar & ", "
Query = Query & "Clase=" & UserList(UserIndex).clase & ", "
Query = Query & "Descripcion='" & UserList(UserIndex).desc & "', "

Query = Query & "Heading=" & CStr(UserList(UserIndex).Char.heading) & ", "

Query = Query & "Head=" & CStr(UserList(UserIndex).OrigChar.Head) & ", "

If UserList(UserIndex).flags.Muerto = 0 Then
    Query = Query & "Body=" & CStr(UserList(UserIndex).Char.Body) & ", "
End If

Query = Query & "Arma=" & CStr(UserList(UserIndex).Char.WeaponAnim) & ", "
Query = Query & "Escudo=" & CStr(UserList(UserIndex).Char.ShieldAnim) & ", "
Query = Query & "Casco=" & CStr(UserList(UserIndex).Char.CascoAnim) & ", "

#If ConUpTime Then
    Dim TempDate As Date
    TempDate = Now - UserList(UserIndex).LogOnTime
    UserList(UserIndex).LogOnTime = Now
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime + (Abs(Day(TempDate) - 30) * 24 * 3600) + Hour(TempDate) * 3600 + Minute(TempDate) * 60 + Second(TempDate)
    UserList(UserIndex).UpTime = UserList(UserIndex).UpTime
    Query = Query & "UpTime=" & UserList(UserIndex).UpTime & ", "
#End If

'First time around?
'If GetVar(UserFile, "INIT", "LastIP1") = vbNullString Then
'    Query = Query & "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & time)
'Is it a different ip from last time?
'ElseIf UserList(UserIndex).ip <> Left$(GetVar(UserFile, "INIT", "LastIP1"), InStr(1, GetVar(UserFile, "INIT", "LastIP1"), " ") - 1) Then
'    Dim i As Integer
'    For i = 5 To 2 Step -1
'        Query = Query & "LastIP" & i, GetVar(UserFile, "INIT", "LastIP" & CStr(i - 1)) & ", "
'    Next i
'    Query = Query & "LastIP1", UserList(UserIndex).ip & " - " & Date & ":" & time)
'Same ip, just update the date
'Else
    Query = Query & "LastIP='" & UserList(UserIndex).ip & "', "
    Query = Query & "LastConnect='" & Format(Date, "YYYYmmdd") & Format(time, "HHmmss") & "', "
'End If



Query = Query & "Map=" & UserList(UserIndex).Pos.map & ", "
Query = Query & "X=" & UserList(UserIndex).Pos.X & ", "
Query = Query & "Y=" & UserList(UserIndex).Pos.Y & ", "


Query = Query & "GLD=" & CStr(UserList(UserIndex).Stats.GLD) & ", "
Query = Query & "BANCO=" & CStr(UserList(UserIndex).Stats.Banco) & ", "

Query = Query & "MaxHP=" & CStr(UserList(UserIndex).Stats.MaxHP) & ", "
Query = Query & "MinHP=" & CStr(UserList(UserIndex).Stats.MinHP) & ", "

Query = Query & "MaxSTA=" & CStr(UserList(UserIndex).Stats.MaxSta) & ", "
Query = Query & "MinSTA=" & CStr(UserList(UserIndex).Stats.MinSta) & ", "

Query = Query & "MaxMAN=" & CStr(UserList(UserIndex).Stats.MaxMAN) & ", "
Query = Query & "MinMAN=" & CStr(UserList(UserIndex).Stats.MinMAN) & ", "

Query = Query & "MaxHIT=" & CStr(UserList(UserIndex).Stats.MaxHIT) & ", "
Query = Query & "MinHIT=" & CStr(UserList(UserIndex).Stats.MinHIT) & ", "

Query = Query & "MaxAGU=" & CStr(UserList(UserIndex).Stats.MaxAGU) & ", "
Query = Query & "MinAGU=" & CStr(UserList(UserIndex).Stats.MinAGU) & ", "

Query = Query & "MaxHAM=" & CStr(UserList(UserIndex).Stats.MaxHam) & ", "
Query = Query & "MinHAM=" & CStr(UserList(UserIndex).Stats.MinHam) & ", "

Query = Query & "SkillPtsLibres=" & CStr(UserList(UserIndex).Stats.SkillPts) & ", "
  
Query = Query & "EXP=" & CStr(UserList(UserIndex).Stats.Exp) & ", "
Query = Query & "ELV=" & CStr(UserList(UserIndex).Stats.ELV) & ", "





Query = Query & "ELU=" & CStr(UserList(UserIndex).Stats.ELU) & ", "
Query = Query & "UserMuertes=" & CStr(UserList(UserIndex).Stats.UsuariosMatados) & ", "
'Query = Query & "CrimMuertes" & CStr(UserList(UserIndex).Stats.CriminalesMatados) & ", "
Query = Query & "NpcsMuertes=" & CStr(UserList(UserIndex).Stats.NPCsMuertos) & ", "
  
'[KEVIN]----------------------------------------------------------------------------
'*******************************************************************************************
Query = Query & "BanCantidadItems=" & val(UserList(UserIndex).BancoInvent.NroItems) & ", "
Dim loopd As Integer
For loopd = 1 To MAX_BANCOINVENTORY_SLOTS
    Query = Query & "BanObj" & loopd & "=" & UserList(UserIndex).BancoInvent.Object(loopd).ObjIndex & ", "
    Query = Query & "BanCant" & loopd & "=" & UserList(UserIndex).BancoInvent.Object(loopd).Amount & ", "
Next loopd
'*******************************************************************************************
'[/KEVIN]-----------
  
'Save Inv
Query = Query & "InvCantidadItems=" & val(UserList(UserIndex).Invent.NroItems) & ", "

For LoopC = 1 To MAX_INVENTORY_SLOTS
    Query = Query & "InvObj" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).ObjIndex & ", "
    Query = Query & "InvCant" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).Amount & ", "
    Query = Query & "InvEqp" & LoopC & "=" & UserList(UserIndex).Invent.Object(LoopC).Equipped & ", "
Next

Query = Query & "WeaponEqpSlot=" & CStr(UserList(UserIndex).Invent.WeaponEqpSlot) & ", "
Query = Query & "ArmourEqpSlot=" & CStr(UserList(UserIndex).Invent.ArmourEqpSlot) & ", "
Query = Query & "CascoEqpSlot=" & CStr(UserList(UserIndex).Invent.CascoEqpSlot) & ", "
Query = Query & "EscudoEqpSlot=" & CStr(UserList(UserIndex).Invent.EscudoEqpSlot) & ", "
Query = Query & "BarcoSlot=" & CStr(UserList(UserIndex).Invent.BarcoSlot) & ", "
Query = Query & "MunicionSlot=" & CStr(UserList(UserIndex).Invent.MunicionEqpSlot) & ", "
'/Nacho

Query = Query & "AnilloSlot=" & CStr(UserList(UserIndex).Invent.AnilloEqpSlot) & ", "


'Reputacion
Query = Query & "Rep_Asesino=" & CStr(UserList(UserIndex).Reputacion.AsesinoRep) & ", "
Query = Query & "Rep_Bandido=" & CStr(UserList(UserIndex).Reputacion.BandidoRep) & ", "
Query = Query & "Rep_Burguesia=" & CStr(UserList(UserIndex).Reputacion.BurguesRep) & ", "
Query = Query & "Rep_Ladrones=" & CStr(UserList(UserIndex).Reputacion.LadronesRep) & ", "
Query = Query & "Rep_Nobles=" & CStr(UserList(UserIndex).Reputacion.NobleRep) & ", "
Query = Query & "Rep_Plebe=" & CStr(UserList(UserIndex).Reputacion.PlebeRep) & ", "

Dim L As Long
L = (-UserList(UserIndex).Reputacion.AsesinoRep) + _
    (-UserList(UserIndex).Reputacion.BandidoRep) + _
    UserList(UserIndex).Reputacion.BurguesRep + _
    (-UserList(UserIndex).Reputacion.LadronesRep) + _
    UserList(UserIndex).Reputacion.NobleRep + _
    UserList(UserIndex).Reputacion.PlebeRep
L = L / 6
Query = Query & "Rep_Promedio=" & CStr(L) & ", "

Dim cad As String

For LoopC = 1 To MAXUSERHECHIZOS
    cad = UserList(UserIndex).Stats.UserHechizos(LoopC)
    Query = Query & "H" & LoopC & "=" & cad & ", "
Next

Dim NroMascotas As Long
NroMascotas = UserList(UserIndex).NroMascotas

For LoopC = 1 To MAXMASCOTAS
    ' Mascota valida?
    If UserList(UserIndex).MascotasIndex(LoopC) > 0 Then
        ' Nos aseguramos que la criatura no fue invocada
        If Npclist(UserList(UserIndex).MascotasIndex(LoopC)).Contadores.TiempoExistencia = 0 Then
            cad = UserList(UserIndex).MascotasType(LoopC)
        Else 'Si fue invocada no la guardamos
            cad = "0"
            NroMascotas = NroMascotas - 1
        End If
        Query = Query & "Masc" & LoopC & "=" & cad & ", "
    Else
        cad = UserList(UserIndex).MascotasType(LoopC)
        Query = Query & "Masc" & LoopC & "=" & cad & ", "
    End If

Next

Query = Query & "NroMascotas=" & CStr(NroMascotas) & ", "

'Devuelve el head de muerto
If UserList(UserIndex).flags.Muerto = 1 Then
    UserList(UserIndex).Char.Head = iCabezaMuerto
End If
'Debug.Print ("UPDATE pjs SET " & left$(Query, Len(Query) - 2) & " WHERE Id=" & UserList(UserIndex).MySQLId)
Call Execute("UPDATE pjs SET " & left$(Query, Len(Query) - 2) & " WHERE Id=" & UserList(UserIndex).MySQLId)

Exit Sub

Errhandler:
Call LogError("Error en SaveUser")

End Sub

Sub BackUPnPc(NpcIndex As Integer)

Dim NpcNumero As Integer
Dim npcfile As String
Dim LoopC As Integer


NpcNumero = Npclist(NpcIndex).Numero

'If NpcNumero > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

'General
Call WriteVar(npcfile, "NPC" & NpcNumero, "Name", Npclist(NpcIndex).Name)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Desc", Npclist(NpcIndex).desc)
Call WriteVar(npcfile, "NPC" & NpcNumero, "Head", val(Npclist(NpcIndex).Char.Head))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Body", val(Npclist(NpcIndex).Char.Body))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Heading", val(Npclist(NpcIndex).Char.heading))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Movement", val(Npclist(NpcIndex).Movement))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Attackable", val(Npclist(NpcIndex).Attackable))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Comercia", val(Npclist(NpcIndex).Comercia))
Call WriteVar(npcfile, "NPC" & NpcNumero, "TipoItems", val(Npclist(NpcIndex).TipoItems))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveEXP", val(Npclist(NpcIndex).GiveEXP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLDMin", val(Npclist(NpcIndex).GiveGLDMin))
Call WriteVar(npcfile, "NPC" & NpcNumero, "GiveGLDMax", val(Npclist(NpcIndex).GiveGLDMax))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Hostil", val(Npclist(NpcIndex).Hostile))
Call WriteVar(npcfile, "NPC" & NpcNumero, "InvReSpawn", val(Npclist(NpcIndex).InvReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "NpcType", val(Npclist(NpcIndex).NPCtype))


'Stats
Call WriteVar(npcfile, "NPC" & NpcNumero, "Alineacion", val(Npclist(NpcIndex).Stats.Alineacion))
Call WriteVar(npcfile, "NPC" & NpcNumero, "DEF", val(Npclist(NpcIndex).Stats.def))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHit", val(Npclist(NpcIndex).Stats.MaxHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MaxHp", val(Npclist(NpcIndex).Stats.MaxHP))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHit", val(Npclist(NpcIndex).Stats.MinHIT))
Call WriteVar(npcfile, "NPC" & NpcNumero, "MinHp", val(Npclist(NpcIndex).Stats.MinHP))




'Flags
Call WriteVar(npcfile, "NPC" & NpcNumero, "ReSpawn", val(Npclist(NpcIndex).flags.ReSpawn))
Call WriteVar(npcfile, "NPC" & NpcNumero, "BackUp", val(Npclist(NpcIndex).flags.BackUp))
Call WriteVar(npcfile, "NPC" & NpcNumero, "Domable", val(Npclist(NpcIndex).flags.Domable))

'Inventario
Call WriteVar(npcfile, "NPC" & NpcNumero, "NroItems", val(Npclist(NpcIndex).Invent.NroItems))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
   For LoopC = 1 To MAX_INVENTORY_SLOTS
        Call WriteVar(npcfile, "NPC" & NpcNumero, "Obj" & LoopC, Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex & "-" & Npclist(NpcIndex).Invent.Object(LoopC).Amount)
   Next
End If


End Sub



Sub CargarNpcBackUp(NpcIndex As Integer, ByVal NpcNumber As Integer)

'Status
If frmMain.Visible Then frmMain.txStatus.Caption = "Cargando backup Npc"

Dim npcfile As String

'If NpcNumber > 499 Then
'    npcfile = DatPath & "bkNPCs-HOSTILES.dat"
'Else
    npcfile = DatPath & "bkNPCs.dat"
'End If

Npclist(NpcIndex).Numero = NpcNumber
Npclist(NpcIndex).Name = GetVar(npcfile, "NPC" & NpcNumber, "Name")
Npclist(NpcIndex).desc = GetVar(npcfile, "NPC" & NpcNumber, "Desc")
Npclist(NpcIndex).Movement = val(GetVar(npcfile, "NPC" & NpcNumber, "Movement"))
Npclist(NpcIndex).NPCtype = val(GetVar(npcfile, "NPC" & NpcNumber, "NpcType"))

Npclist(NpcIndex).Char.Body = val(GetVar(npcfile, "NPC" & NpcNumber, "Body"))
Npclist(NpcIndex).Char.Head = val(GetVar(npcfile, "NPC" & NpcNumber, "Head"))
Npclist(NpcIndex).OrigHeading = val(GetVar(npcfile, "NPC" & NpcNumber, "Heading"))
Npclist(NpcIndex).Char.heading = Npclist(NpcIndex).OrigHeading

Npclist(NpcIndex).Attackable = val(GetVar(npcfile, "NPC" & NpcNumber, "Attackable"))
Npclist(NpcIndex).Comercia = val(GetVar(npcfile, "NPC" & NpcNumber, "Comercia"))
Npclist(NpcIndex).Hostile = val(GetVar(npcfile, "NPC" & NpcNumber, "Hostile"))
Npclist(NpcIndex).GiveEXP = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveEXP")) * MultiplicadorEXP


Npclist(NpcIndex).GiveGLDMin = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLDMin")) * MultiplicadorORO
Npclist(NpcIndex).GiveGLDMax = val(GetVar(npcfile, "NPC" & NpcNumber, "GiveGLDMax")) * MultiplicadorORO

Npclist(NpcIndex).InvReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "InvReSpawn"))

Npclist(NpcIndex).Stats.MaxHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHP"))
Npclist(NpcIndex).Stats.MinHP = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHP"))
Npclist(NpcIndex).Stats.MaxHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MaxHIT"))
Npclist(NpcIndex).Stats.MinHIT = val(GetVar(npcfile, "NPC" & NpcNumber, "MinHIT"))
Npclist(NpcIndex).Stats.def = val(GetVar(npcfile, "NPC" & NpcNumber, "DEF"))
Npclist(NpcIndex).Stats.Alineacion = val(GetVar(npcfile, "NPC" & NpcNumber, "Alineacion"))



Dim LoopC As Integer
Dim ln As String
Npclist(NpcIndex).Invent.NroItems = val(GetVar(npcfile, "NPC" & NpcNumber, "NROITEMS"))
If Npclist(NpcIndex).Invent.NroItems > 0 Then
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        ln = GetVar(npcfile, "NPC" & NpcNumber, "Obj" & LoopC)
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = val(ReadField(1, ln, 45))
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = val(ReadField(2, ln, 45))
       
    Next LoopC
Else
    For LoopC = 1 To MAX_INVENTORY_SLOTS
        Npclist(NpcIndex).Invent.Object(LoopC).ObjIndex = 0
        Npclist(NpcIndex).Invent.Object(LoopC).Amount = 0
    Next LoopC
End If



Npclist(NpcIndex).flags.NPCActive = True
Npclist(NpcIndex).flags.ReSpawn = val(GetVar(npcfile, "NPC" & NpcNumber, "ReSpawn"))
Npclist(NpcIndex).flags.BackUp = val(GetVar(npcfile, "NPC" & NpcNumber, "BackUp"))
Npclist(NpcIndex).flags.Domable = val(GetVar(npcfile, "NPC" & NpcNumber, "Domable"))
Npclist(NpcIndex).flags.RespawnOrigPos = val(GetVar(npcfile, "NPC" & NpcNumber, "OrigPos"))

'Tipo de items con los que comercia
Npclist(NpcIndex).TipoItems = val(GetVar(npcfile, "NPC" & NpcNumber, "TipoItems"))

End Sub


Sub LogBan(ByVal BannedIndex As Integer, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.log", UserList(BannedIndex).Name, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, UserList(BannedIndex).Name
Close #mifile

End Sub


Sub LogBanFromName(ByVal BannedName As String, ByVal UserIndex As Integer, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", UserList(UserIndex).Name)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)

'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub


Sub Ban(ByVal BannedName As String, ByVal Baneador As String, ByVal motivo As String)

Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "BannedBy", Baneador)
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", BannedName, "Reason", motivo)


'Log interno del servidor, lo usa para hacer un UNBAN general de toda la gente banned
Dim mifile As Integer
mifile = FreeFile
Open App.Path & "\logs\GenteBanned.log" For Append Shared As #mifile
Print #mifile, BannedName
Close #mifile

End Sub

Public Sub CargaApuestas()

    Apuestas.Ganancias = val(GetVar(DatPath & "apuestas.dat", "Main", "Ganancias"))
    Apuestas.Perdidas = val(GetVar(DatPath & "apuestas.dat", "Main", "Perdidas"))
    Apuestas.Jugadas = val(GetVar(DatPath & "apuestas.dat", "Main", "Jugadas"))

End Sub

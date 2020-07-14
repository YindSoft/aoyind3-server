Attribute VB_Name = "Admin"
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

Public Type tMotd
    texto As String
    Formato As String
End Type

Public MaxLines As Integer
Public MOTD() As tMotd

Public Type tAPuestas
    Ganancias As Long
    Perdidas As Long
    Jugadas As Long
End Type
Public Apuestas As tAPuestas

Public tInicioServer As Long
Public EstadisticasWeb As New clsEstadisticasIPC

'INTERVALOS
Public SanaIntervaloSinDescansar As Integer
Public StaminaIntervaloLluvia As Integer
Public StaminaIntervaloSinDescansar As Integer
Public SanaIntervaloDescansar As Integer
Public StaminaIntervaloDescansar As Integer
Public IntervaloSed As Integer
Public IntervaloHambre As Integer
Public IntervaloVeneno As Integer
Public IntervaloParalizado As Integer
Public IntervaloInvisible As Integer
Public IntervaloFrio As Integer
Public IntervaloWavFx As Integer
Public IntervaloLanzaHechizo As Integer
Public IntervaloNPCPuedeAtacar As Integer
Public IntervaloNPCAI As Integer
Public IntervaloInvocacion As Integer
Public IntervaloOculto As Integer '[Nacho]
Public IntervaloUserPuedeAtacar As Long
Public IntervaloGolpeUsar As Long
Public IntervaloMagiaGolpe As Long
Public IntervaloGolpeMagia As Long
Public IntervaloUserPuedeCastear As Long
Public IntervaloUserPuedeTrabajar As Long
Public IntervaloParaConexion As Long
Public IntervaloCerrarConexion As Long '[Gonzalo]
Public IntervaloUserPuedeUsar As Long
Public IntervaloFlechasCazadores As Long
Public IntervaloPuedeSerAtacado As Long
Public IntervaloAtacable As Long
Public IntervaloOwnedNpc As Long

'BALANCE

Public PorcentajeRecuperoMana As Integer

Public MinutosWs As Long
Public Puerto As Integer
Public FechaHora As String
Public Hora As Byte

Public BootDelBackUp As Byte
Public Lloviendo As Boolean

Function VersionOK(ByVal Ver As String) As Boolean
VersionOK = (Ver = ULTIMAVERSION)
End Function

Public Function VersionesActuales(ByVal v1 As Integer, ByVal v2 As Integer, ByVal v3 As Integer, ByVal v4 As Integer, ByVal v5 As Integer, ByVal v6 As Integer, ByVal v7 As Integer) As Boolean
Dim rv As Boolean

rv = val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "GRAFICOS")) = v1
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "WAVS")) = v2
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MIDIS")) = v3
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "INIT")) = v4
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "MAPAS")) = v5
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "AOEXE")) = v6
rv = rv And val(GetVar(App.Path & "\AUTOUPDATER\VERSIONES.INI", "ACTUALES", "EXTRAS")) = v7
VersionesActuales = rv

End Function

Sub ReSpawnOrigPosNpcs()
On Error Resume Next

Dim i As Integer
Dim MiNPC As NPC
   
For i = 1 To LastNPC
   'OJO
   If Npclist(i).flags.NPCActive Then
        
        If InMapBounds(Npclist(i).Orig.map, Npclist(i).Orig.X, Npclist(i).Orig.Y) And Npclist(i).Numero = Guardias Then
                MiNPC = Npclist(i)
                Call QuitarNPC(i)
                Call ReSpawnNpc(MiNPC)
        End If
        
        'tildada por sugerencia de yind
        'If Npclist(i).Contadores.TiempoExistencia > 0 Then
        '        Call MuereNpc(i, 0)
        'End If
   End If
   
Next i

End Sub

Sub WorldSave()
On Error Resume Next
'Call LogTarea("Sub WorldSave")

Dim loopX As Integer
Dim Porc As Long

Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Iniciando WorldSave", FontTypeNames.FONTTYPE_SERVER))

#If SeguridadAlkon Then
    Encriptacion.StringValidacion = Encriptacion.ArmarStringValidacion
#End If


Dim j As Integer, k As Integer

FrmStat.ProgressBar1.min = 0
FrmStat.ProgressBar1.max = k
FrmStat.ProgressBar1.value = 0

For loopX = 1 To NumMaps
    'DoEvents
    Call GrabarMapa(loopX, App.Path & "\WorldBackUp\Mapa" & loopX)
    FrmStat.ProgressBar1.value = FrmStat.ProgressBar1.value + 1
Next loopX

FrmStat.Visible = False

If FileExist(DatPath & "\bkNpc.dat", vbNormal) Then Kill (DatPath & "bkNpc.dat")

For loopX = 1 To LastNPC
    If Npclist(loopX).flags.BackUp = 1 Then
        Call BackUPnPc(loopX)
    End If
Next

Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> WorldSave ha concluído", FontTypeNames.FONTTYPE_SERVER))

End Sub

Public Sub PurgarPenas()
    Dim i As Long
    
    For i = 1 To LastUser
        If UserList(i).flags.UserLogged Then
            If UserList(i).Counters.Pena > 0 Then
                UserList(i).Counters.Pena = UserList(i).Counters.Pena - 1
                
                If UserList(i).Counters.Pena < 1 Then
                    UserList(i).Counters.Pena = 0
                    'Le quito la ropa carcelera
                    If UserList(i).flags.Muerto = 0 Then
                        UserList(i).Char.Body = ObjData(UserList(i).Invent.ArmourEqpObjIndex).Ropaje
                        Call ChangeUserChar(i, UserList(i).Char.Body, UserList(i).Char.Head, UserList(i).Char.heading, UserList(i).Char.WeaponAnim, UserList(i).Char.ShieldAnim, UserList(i).Char.CascoAnim)
                    End If
                    Call WarpUserChar(i, Libertad.map, Libertad.X, Libertad.Y, True)
                    Call WriteConsoleMsg(i, "Has sido liberado!", FontTypeNames.FONTTYPE_INFO)
                    
                    Call FlushBuffer(i)
                End If
            End If
        End If
    Next i
End Sub


Public Sub Encarcelar(ByVal UserIndex As Integer, ByVal Minutos As Long, Optional ByVal GmName As String = vbNullString)
        
        UserList(UserIndex).Counters.Pena = Minutos
       
        
        Call WarpUserChar(UserIndex, Prision.map, Prision.X, Prision.Y, True)
        
        'Ropita carcelera
        If UserList(UserIndex).flags.Muerto = 0 Then
            If UserList(UserIndex).raza = eRaza.Humano Or _
               UserList(UserIndex).raza = eRaza.Elfo Or _
               UserList(UserIndex).raza = eRaza.Drow Then
                UserList(UserIndex).Char.Body = ROPA_CARCEL_ALTOS
            Else
                UserList(UserIndex).Char.Body = ROPA_CARCEL_BAJOS
            End If
            Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.heading, 2, 2, 2)

        End If
        
        
        If LenB(GmName) = 0 Then
            Call WriteConsoleMsg(UserIndex, "Has sido encarcelado, deberas permanecer en la carcel por " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        Else
            Call WriteConsoleMsg(UserIndex, GmName & " te ha encarcelado, deberas permanecer en la carcel por " & Minutos & " minutos.", FontTypeNames.FONTTYPE_INFO)
        End If
        If UserList(UserIndex).flags.Traveling = 1 Then
            UserList(UserIndex).flags.Traveling = 0
            UserList(UserIndex).Counters.goHome = 0
            Call WriteGotHome(UserIndex, False)
        End If
        
End Sub


Public Sub BorrarUsuario(ByVal UserName As String)
On Error Resume Next
'If FileExist(CharPath & UCase$(UserName) & ".chr", vbNormal) Then
'    Kill CharPath & UCase$(UserName) & ".chr"
'End If
Call Execute("DELETE FROM pjs WHERE Nombre=" & Comillas(UserName))
End Sub

Public Function BANCheck(ByVal Name As String) As Boolean

BANCheck = CStrNull(GetByCampo("SELECT Ban FROM pjs WHERE Nombre=" & Comillas(Name), "Ban")) = "1"

End Function

Public Function PersonajeExiste(ByVal Name As String) As Boolean

PersonajeExiste = GetByCampo("SELECT COUNT(Id) as 'Cantidad' FROM pjs WHERE Nombre=" & Comillas(Name), "Cantidad") = "1"

End Function

Public Function IdPersonaje(ByVal Name As String) As Long

IdPersonaje = val(GetByCampo("SELECT Id FROM pjs WHERE Nombre=" & Comillas(Name), "Id"))

End Function

Public Function UnBan(ByVal Name As String) As Boolean
'Unban the character
'Call WriteVar(App.Path & "\charfile\" & name & ".chr", "FLAGS", "Ban", "0")
Call Execute("UPDATE pjs SET Ban=0 WHERE Nombre=" & Comillas(Name))

'Remove it from the banned people database
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "BannedBy", "NOBODY")
Call WriteVar(App.Path & "\logs\" & "BanDetail.dat", Name, "Reason", "NO REASON")
End Function

Public Function MD5ok(ByVal md5formateado As String) As Boolean
Dim i As Integer

If MD5ClientesActivado = 1 Then
    For i = 0 To UBound(MD5s)
        If (md5formateado = MD5s(i)) Then
            MD5ok = True
            Exit Function
        End If
    Next i
    MD5ok = False
Else
    MD5ok = True
End If

End Function

Public Sub MD5sCarga()
Dim LoopC As Integer

MD5ClientesActivado = val(GetVar(IniPath & "Server.ini", "MD5Hush", "Activado"))

If MD5ClientesActivado = 1 Then
    ReDim MD5s(val(GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptados")))
    For LoopC = 0 To UBound(MD5s)
        MD5s(LoopC) = GetVar(IniPath & "Server.ini", "MD5Hush", "MD5Aceptado" & (LoopC + 1))
        MD5s(LoopC) = txtOffset(hexMd52Asc(MD5s(LoopC)), 55)
    Next LoopC
End If

End Sub

Public Sub BanIpAgrega(ByVal ip As String)
    BanIps.Add ip
    
    Call BanIpGuardar
End Sub

Public Function BanIpBuscar(ByVal ip As String) As Long
Dim Dale As Boolean
Dim LoopC As Long

Dale = True
LoopC = 1
Do While LoopC <= BanIps.Count And Dale
    Dale = (BanIps.Item(LoopC) <> ip)
    LoopC = LoopC + 1
Loop

If Dale Then
    BanIpBuscar = 0
Else
    BanIpBuscar = LoopC - 1
End If
End Function

Public Function BanIpQuita(ByVal ip As String) As Boolean

On Error Resume Next

Dim N As Long

N = BanIpBuscar(ip)
If N > 0 Then
    BanIps.Remove N
    BanIpGuardar
    BanIpQuita = True
Else
    BanIpQuita = False
End If

End Function

Public Sub BanIpGuardar()
Dim ArchivoBanIp As String
Dim ArchN As Long
Dim LoopC As Long

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

ArchN = FreeFile()
Open ArchivoBanIp For Output As #ArchN

For LoopC = 1 To BanIps.Count
    Print #ArchN, BanIps.Item(LoopC)
Next LoopC

Close #ArchN

End Sub

Public Sub BanIpCargar()
Dim ArchN As Long
Dim Tmp As String
Dim ArchivoBanIp As String

ArchivoBanIp = App.Path & "\Dat\BanIps.dat"

Do While BanIps.Count > 0
    BanIps.Remove 1
Loop

ArchN = FreeFile()
Open ArchivoBanIp For Input As #ArchN

Do While Not EOF(ArchN)
    Line Input #ArchN, Tmp
    BanIps.Add Tmp
Loop

Close #ArchN

End Sub

Public Sub ActualizaEstadisticasWeb()

Static Andando As Boolean
Static Contador As Long
Dim Tmp As Boolean

Contador = Contador + 1

If Contador >= 10 Then
    Contador = 0
    Tmp = EstadisticasWeb.EstadisticasAndando()
    
    If Andando = False And Tmp = True Then
        Call InicializaEstadisticas
    End If
    
    Andando = Tmp
End If

End Sub

Public Function UserDarPrivilegioLevel(ByVal Name As String) As PlayerType
'***************************************************
'Author: Unknown
'Last Modification: 03/02/07
'Last Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    If EsAdmin(Name) Then
        UserDarPrivilegioLevel = PlayerType.Admin
    ElseIf EsDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.Dios
    ElseIf EsSemiDios(Name) Then
        UserDarPrivilegioLevel = PlayerType.SemiDios
    ElseIf EsConsejero(Name) Then
        UserDarPrivilegioLevel = PlayerType.Consejero
    Else
        UserDarPrivilegioLevel = PlayerType.user
    End If
End Function

Public Sub BanCharacter(ByVal bannerUserIndex As Integer, ByVal UserName As String, ByVal Reason As String)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 03/02/07
'
'***************************************************
    Dim tUser As Integer
    Dim userPriv As Byte
    Dim cantPenas As Byte
    Dim rank As Integer
    Dim IdPj As Long
    If InStrB(UserName, "+") Then
        UserName = Replace(UserName, "+", " ")
    End If
    
    tUser = NameIndex(UserName)
    
    rank = PlayerType.Admin Or PlayerType.Dios Or PlayerType.SemiDios Or PlayerType.Consejero
    
    With UserList(bannerUserIndex)
        If tUser <= 0 Then
            Call WriteConsoleMsg(bannerUserIndex, "El usuario no esta online.", FontTypeNames.FONTTYPE_TALK)
            
            If PersonajeExiste(UserName) Then
                userPriv = UserDarPrivilegioLevel(UserName)
                
                If (userPriv And rank) > (.flags.Privilegios And rank) Then
                    Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
                Else
                    If GetByCampo("SELECT Ban FROM pjs WHERE Nombre=" & Comillas(UserName), "Ban") <> "0" Then
                        Call WriteConsoleMsg(bannerUserIndex, "El personaje ya se encuentra baneado.", FontTypeNames.FONTTYPE_INFO)
                    Else
                        Call LogBanFromName(UserName, bannerUserIndex, Reason)
                        Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserName & ".", FontTypeNames.FONTTYPE_SERVER))
                        
                        IdPj = val(GetByCampo("SELECT Id FROM pjs WHERE Nombre=" & Comillas(UserName), "Id"))
                        
                        'ponemos el flag de ban a 1
                        Execute ("UPDATE pjs SET Penas=Penas+1, Ban=1 WHERE Id=" & IdPj)
                        'ponemos la pena
                        Execute ("INSERT INTO penas (IdPj, Razon, Fecha, IdGM, Tiempo) VALUES (" & IdPj & "," & Comillas(LCase$(.Name) & ": BAN POR " & Reason) & ",NOW()," & .MySQLId & ",0)")

                        If (userPriv And rank) = (.flags.Privilegios And rank) Then
                            .flags.Ban = 1
                            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                            Call CloseSocket(bannerUserIndex)
                        End If
                        
                        Call LogGM(.Name, "BAN a " & UserName)
                    End If
                End If
            Else
                Call WriteConsoleMsg(bannerUserIndex, "El pj " & UserName & " no existe.", FontTypeNames.FONTTYPE_INFO)
            End If
        Else
            If (UserList(tUser).flags.Privilegios And rank) > (.flags.Privilegios And rank) Then
                Call WriteConsoleMsg(bannerUserIndex, "No podes banear a al alguien de mayor jerarquia.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Call LogBan(tUser, bannerUserIndex, Reason)
            Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg("Servidor> " & .Name & " ha baneado a " & UserList(tUser).Name & ".", FontTypeNames.FONTTYPE_SERVER))
            
            'Ponemos el flag de ban a 1
            UserList(tUser).flags.Ban = 1
            
            If (UserList(tUser).flags.Privilegios And rank) = (.flags.Privilegios And rank) Then
                .flags.Ban = 1
                Call SendData(SendTarget.ToAdmins, 0, PrepareMessageConsoleMsg(.Name & " banned by the server por bannear un Administrador.", FontTypeNames.FONTTYPE_FIGHT))
                Call CloseSocket(bannerUserIndex)
            End If
            
            Call LogGM(.Name, "BAN a " & UserName)
            
            IdPj = UserList(tUser).MySQLId
                        
            'ponemos el flag de ban a 1
            Execute ("UPDATE pjs SET Penas=Penas+1, Ban=1 WHERE Id=" & IdPj)
            'ponemos la pena
            Execute ("INSERT INTO penas (IdPj, Razon, Fecha, IdGM, Tiempo) VALUES (" & IdPj & "," & Comillas(LCase$(.Name) & ": BAN POR " & Reason) & ",NOW()," & .MySQLId & ",0)")

            Call CloseSocket(tUser)
        End If
    End With
End Sub


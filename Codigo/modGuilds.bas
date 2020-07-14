Attribute VB_Name = "modGuilds"
'**************************************************************
' modGuilds.bas - Module to allow the usage of areas instead of maps.
' Saves a lot of bandwidth.
'
' Implemented by Mariano Barrou (El Oso)
'**************************************************************

'**************************************************************************
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
'**************************************************************************

Option Explicit

'guilds nueva version. Hecho por el oso, eliminando los problemas
'de sincronizacion con los datos en el HD... entre varios otros
'º¬

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'DECLARACIOENS PUBLICAS CONCERNIENTES AL JUEGO
'Y CONFIGURACION DEL SISTEMA DE CLANES
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const MAX_GUILDS As Integer = 1000
'cantidad maxima de guilds en el servidor

Public CANTIDADDECLANES As Integer
'cantidad actual de clanes en el servidor

Private guilds As Dictionary
'array global de guilds, se indexa por userlist().guildindex

Private Const CANTIDADMAXIMACODEX As Byte = 8
'cantidad maxima de codecs que se pueden definir

Public Const MAXASPIRANTES As Byte = 10
'cantidad maxima de aspirantes que puede tener un clan acumulados a la vez

Private Const MAXANTIFACCION As Byte = 5
'puntos maximos de antifaccion que un clan tolera antes de ser cambiada su alineacion

Public Enum ALINEACION_GUILD
    ALINEACION_LEGION = 1
    ALINEACION_CRIMINAL = 2
    ALINEACION_NEUTRO = 3
    ALINEACION_CIUDA = 4
    ALINEACION_ARMADA = 5
    ALINEACION_MASTER = 6
End Enum
'alineaciones permitidas

Public Enum SONIDOS_GUILD
    SND_CREACIONCLAN = 44
    SND_ACEPTADOCLAN = 43
    SND_DECLAREWAR = 45
End Enum
'numero de .wav del cliente

Public Enum RELACIONES_GUILD
    GUERRA = -1
    PAZ = 0
    ALIADOS = 1
End Enum

Public Const OROFUNDARCLAN As Long = 2000000
'estado entre clanes
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub LoadGuildsDB()

Dim i           As Integer
Dim Id As Integer
Dim Clan As clsClan

    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Set guilds = New Dictionary
    
    Cant = mySQL.SQLQuery("SELECT * FROM clanes", Datos)
    
    If Cant > 0 Then
        CANTIDADDECLANES = Cant
    End If
    
    
    'Datos("Rep_Promedio
            
    For i = 1 To CANTIDADDECLANES
        Set Clan = New clsClan
        Id = Datos("Id")
        Call Clan.Inicializar(Datos)
        
        Call guilds.Add(Id, Clan)
        
        Datos.MoveNext
    Next i
    
End Sub

Public Function m_ConectarMiembroAClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
Dim NuevoL  As Boolean
Dim NuevaA  As Boolean
Dim News    As String

    If GuildIndex <= 0 Then Exit Function 'x las dudas...
    If m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
        Call guilds(GuildIndex).ConectarMiembro(UserIndex)
        UserList(UserIndex).GuildIndex = GuildIndex
        m_ConectarMiembroAClan = True
    Else
        m_ConectarMiembroAClan = m_ValidarPermanencia(UserIndex, True, NuevaA, NuevoL)
        If NuevoL Then News = "El clan tiene nuevo líder."
        If NuevaA Then News = News & "El clan tiene nueva alineación."
        If NuevoL Or NuevaA Then Call guilds(GuildIndex).SetGuildNews(News)
    End If

End Function


Public Function m_ValidarPermanencia(ByVal UserIndex As Integer, ByVal SumaAntifaccion As Boolean, ByRef CambioAlineacion As Boolean, ByRef CambioLider As Boolean) As Boolean
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 25/03/2009
'25/03/2009: ZaMa - Desequipo los items faccionarios que tenga el funda al abandonar la faccion
'***************************************************

Dim GuildIndex  As Integer
Dim ML()        As String
Dim M           As String
Dim UI          As Integer
Dim Sale        As Boolean
Dim i           As Long
Dim totalMembers As Integer

    m_ValidarPermanencia = True
    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex <= 0 Then Exit Function
    
    If Not m_EstadoPermiteEntrar(UserIndex, GuildIndex) Then
    
        Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & " es expulsado en validar permanencia")
    
        m_ValidarPermanencia = False
        If SumaAntifaccion Then guilds(GuildIndex).PuntosAntifaccion = guilds(GuildIndex).PuntosAntifaccion + 1
        
        CambioAlineacion = (m_EsGuildFounder(UserList(UserIndex).Name, GuildIndex) Or guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION)
        
        Call LogClanes(UserList(UserIndex).Name & " de " & guilds(GuildIndex).GuildName & IIf(CambioAlineacion, " SI ", " NO ") & "provoca cambio de alinaecion. MAXANT:" & (guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION) & ", GUILDFOU:" & m_EsGuildFounder(UserList(UserIndex).Name, GuildIndex))
        
        If CambioAlineacion Then
            'aca tenemos un problema, el fundador acaba de cambiar el rumbo del clan o nos zarpamos de antifacciones
            'Tenemos que resetear el lider, revisar si el lider permanece y si no asignarle liderazgo al fundador

            Call guilds(GuildIndex).CambiarAlineacion(ALINEACION_NEUTRO)
            guilds(GuildIndex).PuntosAntifaccion = MAXANTIFACCION
            'para la nueva alineacion, hay que revisar a todos los Pjs!

            'uso GetMemberList y no los iteradores pq voy a rajar gente y puedo alterar
            'internamente al iterador en el proceso
            CambioLider = False
            ML = guilds(GuildIndex).GetMemberList()
            totalMembers = UBound(ML)
            
            'El user en ML(0) es el funda, no tiene setnido verificar su permanencia!
            For i = 1 To totalMembers
                M = ML(i)
                
                'vamos a violar un poco de capas..
                UI = NameIndex(M)
                If UI > 0 Then
                    Sale = Not m_EstadoPermiteEntrar(UI, GuildIndex)
                Else
                    Sale = Not m_EstadoPermiteEntrarChar(M, GuildIndex)
                End If

                If Sale Then
                    If m_EsGuildFounder(M, GuildIndex) Then 'hay que sacarlo de las armadas
                     
                        If UI > 0 Then
                            If UserList(UI).Faccion.ArmadaReal <> 0 Then
                                Call ExpulsarFaccionReal(UI)
                            ElseIf UserList(UI).Faccion.FuerzasCaos <> 0 Then
                                Call ExpulsarFaccionCaos(UI)
                            End If
                           UserList(UI).Faccion.Reenlistadas = 200
                        Else
                            If PersonajeExiste(M) Then
                                Execute ("UPDATE pjs SET EjercitoCaos=0, EjercitoReal=0, Reenlistadas=200 WHERE Nombre=" & Comillas(M))
                            End If
                        End If
                        m_ValidarPermanencia = True
                    Else    'sale si no es guildfounder
                        If m_EsGuildLeader(M, GuildIndex) Then
                            'pierde el liderazgo
                            CambioLider = True
                            Call guilds(GuildIndex).SetLeader(guilds(GuildIndex).Fundador)
                        End If

                        Call m_EcharMiembroDeClan(-1, M)
                    End If
                End If
            Next i
        Else
            'no se va el fundador, el peor caso es que se vaya el lider
            
            'If m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
            '    Call LogClanes("Se transfiere el liderazgo de: " & Guilds(GuildIndex).GuildName & " a " & Guilds(GuildIndex).Fundador)
            '    Call Guilds(GuildIndex).SetLeader(Guilds(GuildIndex).Fundador)  'transferimos el lideraztgo
            'End If
            Call m_EcharMiembroDeClan(-1, UserList(UserIndex).Name)   'y lo echamos
        End If
    End If
End Function

Public Sub m_DesconectarMiembroDelClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
    If UserList(UserIndex).GuildIndex < 1 Then Exit Sub
    Call guilds(GuildIndex).DesConectarMiembro(UserIndex)
End Sub

Private Function m_EsGuildLeader(ByRef Pj As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildLeader = (UCase$(Pj) = UCase$(Trim$(guilds(GuildIndex).GetLeader)))
End Function

Private Function m_EsGuildFounder(ByRef Pj As String, ByVal GuildIndex As Integer) As Boolean
    m_EsGuildFounder = (UCase$(Pj) = UCase$(Trim$(guilds(GuildIndex).Fundador)))
End Function

Public Function m_EcharMiembroDeClan(ByVal Expulsador As Integer, ByVal Expulsado As String) As Integer
'UI echa a Expulsado del clan de Expulsado
Dim UserIndex   As Integer
Dim GI          As Integer
    
    m_EcharMiembroDeClan = 0

    UserIndex = NameIndex(Expulsado)
    If UserIndex > 0 Then
        'pj online
        GI = UserList(UserIndex).GuildIndex
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).DesConectarMiembro(UserIndex)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                UserList(UserIndex).GuildIndex = 0
                Call RefreshCharStatus(UserIndex)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    Else
        'pj offline
        GI = GetGuildIndexFromChar(Expulsado)
        If GI > 0 Then
            If m_PuedeSalirDeClan(Expulsado, GI, Expulsador) Then
                If m_EsGuildLeader(Expulsado, GI) Then guilds(GI).SetLeader (guilds(GI).Fundador)
                Call guilds(GI).ExpulsarMiembro(Expulsado)
                Call LogClanes(Expulsado & " ha sido expulsado de " & guilds(GI).GuildName & " Expulsador = " & Expulsador)
                m_EcharMiembroDeClan = GI
            Else
                m_EcharMiembroDeClan = 0
            End If
        Else
            m_EcharMiembroDeClan = 0
        End If
    End If

End Function

Public Sub ActualizarWebSite(ByVal UserIndex As Integer, ByRef Web As String)
Dim GI As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetURL(Web)
    
End Sub


Public Sub ChangeCodexAndDesc(ByRef desc As String, ByRef Codex() As String, ByVal GuildIndex As Integer)
    Dim i As Long
    
    If GuildIndex < 1 Then Exit Sub
    
    With guilds(GuildIndex)
        Call .SetDesc(desc)
        
        For i = 0 To UBound(Codex())
            Call .SetCodex(i, Codex(i))
        Next i
        
        For i = i To CANTIDADMAXIMACODEX
            Call .SetCodex(i, vbNullString)
        Next i
    End With
End Sub

Public Sub ActualizarNoticias(ByVal UserIndex As Integer, ByRef Datos As String)
Dim GI              As Integer

    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Then Exit Sub
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then Exit Sub
    
    Call guilds(GI).SetGuildNews(Datos)
        
End Sub

Public Function CrearNuevoClan(ByVal FundadorIndex As Integer, ByRef desc As String, ByRef GuildName As String, ByRef URL As String, ByRef Codex() As String, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean
Dim CantCodex       As Integer
Dim i               As Integer
Dim DummyString     As String
Dim Clan As clsClan
Dim Id As Integer
Dim Datos As clsMySQLRecordSet
Dim Cant As Long
    CrearNuevoClan = False
    If Not PuedeFundarUnClan(FundadorIndex, Alineacion, DummyString) Then
        refError = DummyString
        Exit Function
    End If

    If GuildName = vbNullString Or Not GuildNameValido(GuildName) Then
        refError = "Nombre de clan inválido."
        Exit Function
    End If
    
    If YaExiste(GuildName) Then
        refError = "Ya existe un clan con ese nombre."
        Exit Function
    End If

    CantCodex = UBound(Codex()) + 1

    UserList(FundadorIndex).Stats.GLD = UserList(FundadorIndex).Stats.GLD - OROFUNDARCLAN
    Call WriteUpdateGold(FundadorIndex)

        'constructor custom de la clase clan
        Set Clan = New clsClan
        CANTIDADDECLANES = CANTIDADDECLANES + 1
        'Damos de alta al clan como nuevo inicializando sus archivos
        Id = Clan.InicializarNuevoClan(GuildName, UserList(FundadorIndex).Name, Alineacion, desc, URL, Codex)
        Cant = mySQL.SQLQuery("SELECT * FROM clanes WHERE Id=" & Id, Datos)
        Call Clan.Inicializar(Datos)
        Clan.AceptarNuevoMiembro (UserList(FundadorIndex).Name)
        Clan.ConectarMiembro (FundadorIndex)
        
        Call guilds.Add(Id, Clan)
        
        UserList(FundadorIndex).GuildIndex = Id
        Call RefreshCharStatus(FundadorIndex)
        
    CrearNuevoClan = True
End Function

Public Sub SendGuildNews(ByVal UserIndex As Integer)
Dim GuildIndex  As Integer
Dim i               As Integer
Dim go As Integer

    GuildIndex = UserList(UserIndex).GuildIndex
    If GuildIndex = 0 Then Exit Sub

    Dim enemies() As String
    
    If guilds(GuildIndex).CantidadEnemys Then
        ReDim enemies(0 To guilds(GuildIndex).CantidadEnemys - 1) As String
    Else
        ReDim enemies(0)
    End If
    
    Dim allies() As String
    
    If guilds(GuildIndex).CantidadAllies Then
        ReDim allies(0 To guilds(GuildIndex).CantidadAllies - 1) As String
    Else
        ReDim allies(0)
    End If
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
    go = 0
    
    While i > 0
        enemies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.GUERRA)
        go = go + 1
    Wend
    
    i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    go = 0
    
    While i > 0
        allies(go) = guilds(i).GuildName
        i = guilds(GuildIndex).Iterador_ProximaRelacion(RELACIONES_GUILD.ALIADOS)
    Wend

    Call WriteGuildNews(UserIndex, guilds(GuildIndex).GetGuildNews, enemies, allies)

    If guilds(GuildIndex).EleccionesAbiertas Then
        Call WriteConsoleMsg(UserIndex, "Hoy es la votacion para elegir un nuevo líder para el clan!!.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "La eleccion durara 24 horas, se puede votar a cualquier miembro del clan.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Para votar escribe /VOTO NICKNAME.", FontTypeNames.FONTTYPE_GUILD)
        Call WriteConsoleMsg(UserIndex, "Solo se computara un voto por miembro. Tu voto no puede ser cambiado.", FontTypeNames.FONTTYPE_GUILD)
    End If

End Sub

Public Function m_PuedeSalirDeClan(ByRef Nombre As String, ByVal GuildIndex As Integer, ByVal QuienLoEchaUI As Integer) As Boolean
'sale solo si no es fundador del clan.

    m_PuedeSalirDeClan = False
    If GuildIndex = 0 Then Exit Function
    
    'esto es un parche, si viene en -1 es porque la invoca la rutina de expulsion automatica de clanes x antifacciones
    If QuienLoEchaUI = -1 Then
        m_PuedeSalirDeClan = True
        Exit Function
    End If

    'cuando UI no puede echar a nombre?
    'si no es gm Y no es lider del clan del pj Y no es el mismo que se va voluntariamente
    If UserList(QuienLoEchaUI).flags.Privilegios And PlayerType.user Then
        If Not m_EsGuildLeader(UCase$(UserList(QuienLoEchaUI).Name), GuildIndex) Then
            If UCase$(UserList(QuienLoEchaUI).Name) <> UCase$(Nombre) Then      'si no sale voluntariamente...
                Exit Function
            End If
        End If
    End If

    m_PuedeSalirDeClan = UCase$(guilds(GuildIndex).Fundador) <> UCase$(Nombre)

End Function

Public Function PuedeFundarUnClan(ByVal UserIndex As Integer, ByVal Alineacion As ALINEACION_GUILD, ByRef refError As String) As Boolean

    PuedeFundarUnClan = False
    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, no puedes fundar otro"
        Exit Function
    End If
    
    If UserList(UserIndex).Stats.ELV < 35 Or UserList(UserIndex).Stats.UserSkills(eSkill.Liderazgo) < 90 Or UserList(UserIndex).Stats.GLD < OROFUNDARCLAN Then
        refError = "Para fundar un clan debes ser nivel 35, tener 90 en liderazgo y pagar " & OROFUNDARCLAN & " monedas de oro."
        Exit Function
    End If
    
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            If UserList(UserIndex).Faccion.ArmadaReal <> 1 Then
                refError = "Para fundar un clan real debes ser miembro de la armada."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            If Criminal(UserIndex) Then
                refError = "Para fundar un clan de ciudadanos no debes ser criminal."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            If Not Criminal(UserIndex) Then
                refError = "Para fundar un clan de criminales no debes ser ciudadano."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_LEGION
            If UserList(UserIndex).Faccion.FuerzasCaos <> 1 Then
                refError = "Para fundar un clan del mal debes pertenecer a la legión oscura"
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_MASTER
            If UserList(UserIndex).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios) Then
                refError = "Para fundar un clan sin alineación debes ser un dios."
                Exit Function
            End If
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            If UserList(UserIndex).Faccion.ArmadaReal <> 0 Or UserList(UserIndex).Faccion.FuerzasCaos <> 0 Then
                refError = "Para fundar un clan neutro no debes pertenecer a ninguna facción."
                Exit Function
            End If
    End Select
    
    PuedeFundarUnClan = True
    
End Function

Private Function m_EstadoPermiteEntrarChar(ByRef Personaje As String, ByVal GuildIndex As Integer) As Boolean
Dim Promedio    As Long
Dim ELV         As Integer
Dim f           As Byte

    m_EstadoPermiteEntrarChar = False
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace(Personaje, ".", vbNullString)
    End If
    
    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT ELV, Rep_Promedio, EjercitoReal, EjercitoCaos FROM pjs WHERE Nombre=" & Comillas(Personaje), Datos)
    
    If Cant > 0 Then
        Promedio = CLng(Datos("Rep_Promedio"))
        Select Case guilds(GuildIndex).Alineacion
            Case ALINEACION_GUILD.ALINEACION_ARMADA
                If Promedio >= 0 Then
                    ELV = CInt(Datos("ELV"))
                    If ELV >= 25 Then
                        f = CByte(Datos("EjercitoReal"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case ALINEACION_GUILD.ALINEACION_CIUDA
                m_EstadoPermiteEntrarChar = Promedio >= 0
            Case ALINEACION_GUILD.ALINEACION_CRIMINAL
                m_EstadoPermiteEntrarChar = Promedio < 0
            Case ALINEACION_GUILD.ALINEACION_NEUTRO
                m_EstadoPermiteEntrarChar = CByte(Datos("EjercitoReal")) = 0
                m_EstadoPermiteEntrarChar = m_EstadoPermiteEntrarChar And (CByte(Datos("EjercitoCaos")) = 0)
            Case ALINEACION_GUILD.ALINEACION_LEGION
                If Promedio < 0 Then
                    ELV = CInt(Datos("ELV"))
                    If ELV >= 25 Then
                        f = CByte(Datos("EjercitoCaos"))
                    End If
                    m_EstadoPermiteEntrarChar = IIf(ELV >= 25, f <> 0, True)
                End If
            Case Else
                m_EstadoPermiteEntrarChar = True
        End Select
    End If
End Function

Private Function m_EstadoPermiteEntrar(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As Boolean
    Select Case guilds(GuildIndex).Alineacion
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            m_EstadoPermiteEntrar = Not Criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.ArmadaReal <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_LEGION
            m_EstadoPermiteEntrar = Criminal(UserIndex) And _
                    IIf(UserList(UserIndex).Stats.ELV >= 25, UserList(UserIndex).Faccion.FuerzasCaos <> 0, True)
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            m_EstadoPermiteEntrar = UserList(UserIndex).Faccion.ArmadaReal = 0 And UserList(UserIndex).Faccion.FuerzasCaos = 0
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            m_EstadoPermiteEntrar = Not Criminal(UserIndex)
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            m_EstadoPermiteEntrar = Criminal(UserIndex)
        Case Else   'game masters
            m_EstadoPermiteEntrar = True
    End Select
End Function


Public Function String2Alineacion(ByRef S As String) As ALINEACION_GUILD
    Select Case S
        Case "Neutro"
            String2Alineacion = ALINEACION_NEUTRO
        Case "Legión oscura"
            String2Alineacion = ALINEACION_LEGION
        Case "Armada Real"
            String2Alineacion = ALINEACION_ARMADA
        Case "Game Masters"
            String2Alineacion = ALINEACION_MASTER
        Case "Legal"
            String2Alineacion = ALINEACION_CIUDA
        Case "Criminal"
            String2Alineacion = ALINEACION_CRIMINAL
    End Select
End Function

Public Function Alineacion2String(ByVal Alineacion As ALINEACION_GUILD) As String
    Select Case Alineacion
        Case ALINEACION_GUILD.ALINEACION_NEUTRO
            Alineacion2String = "Neutro"
        Case ALINEACION_GUILD.ALINEACION_LEGION
            Alineacion2String = "Legión oscura"
        Case ALINEACION_GUILD.ALINEACION_ARMADA
            Alineacion2String = "Armada Real"
        Case ALINEACION_GUILD.ALINEACION_MASTER
            Alineacion2String = "Game Masters"
        Case ALINEACION_GUILD.ALINEACION_CIUDA
            Alineacion2String = "Legal"
        Case ALINEACION_GUILD.ALINEACION_CRIMINAL
            Alineacion2String = "Criminal"
    End Select
End Function

Public Function Relacion2String(ByVal Relacion As RELACIONES_GUILD) As String
    Select Case Relacion
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "A"
        Case RELACIONES_GUILD.GUERRA
            Relacion2String = "G"
        Case RELACIONES_GUILD.PAZ
            Relacion2String = "P"
        Case RELACIONES_GUILD.ALIADOS
            Relacion2String = "?"
    End Select
End Function

Public Function String2Relacion(ByVal S As String) As RELACIONES_GUILD
    Select Case UCase$(Trim$(S))
        Case vbNullString, "P"
            String2Relacion = RELACIONES_GUILD.PAZ
        Case "G"
            String2Relacion = RELACIONES_GUILD.GUERRA
        Case "A"
            String2Relacion = RELACIONES_GUILD.ALIADOS
        Case Else
            String2Relacion = RELACIONES_GUILD.PAZ
    End Select
End Function

Private Function GuildNameValido(ByVal cad As String) As Boolean
Dim car     As Byte
Dim i       As Integer

'old function by morgo

cad = LCase$(cad)

For i = 1 To Len(cad)
    car = Asc(mId$(cad, i, 1))

    If (car < 97 Or car > 122) And (car <> 255) And (car <> 32) Then
        GuildNameValido = False
        Exit Function
    End If
    
Next i

GuildNameValido = True

End Function

Private Function YaExiste(ByVal GuildName As String) As Boolean
Dim i   As Integer

YaExiste = False
GuildName = UCase$(GuildName)

For i = 0 To CANTIDADDECLANES - 1
    YaExiste = (UCase$(guilds.Items(i).GuildName) = GuildName)
    If YaExiste Then Exit Function
Next i



End Function

Public Function v_AbrirElecciones(ByVal UserIndex As Integer, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer

    v_AbrirElecciones = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GuildIndex) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If guilds(GuildIndex).EleccionesAbiertas Then
        refError = "Las elecciones ya están abiertas"
        Exit Function
    End If
    
    v_AbrirElecciones = True
    Call guilds(GuildIndex).AbrirElecciones
    
End Function

Public Function v_UsuarioVota(ByVal UserIndex As Integer, ByRef Votado As String, ByRef refError As String) As Boolean
Dim GuildIndex      As Integer
Dim list()          As String
Dim i As Long

    v_UsuarioVota = False
    GuildIndex = UserList(UserIndex).GuildIndex
    
    If GuildIndex = 0 Then
        refError = "Tu no perteneces a ningún clan"
        Exit Function
    End If

    If Not guilds(GuildIndex).EleccionesAbiertas Then
        refError = "No hay elecciones abiertas en tu clan."
        Exit Function
    End If
    
    
    list = guilds(GuildIndex).GetMemberList()
    For i = 0 To UBound(list())
        If UCase$(Votado) = UCase$(list(i)) Then Exit For
    Next i
    
    If i > UBound(list()) Then
        refError = Votado & " no pertenece al clan"
        Exit Function
    End If
    
    
    If guilds(GuildIndex).YaVoto(UserList(UserIndex).Name) Then
        refError = "Ya has votado, no puedes cambiar tu voto"
        Exit Function
    End If
    
    Call guilds(GuildIndex).ContabilizarVoto(UserList(UserIndex).Name, Votado)
    v_UsuarioVota = True

End Function

Public Sub v_RutinaElecciones()
Dim i       As Integer

On Error GoTo errh
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Revisando elecciones", FontTypeNames.FONTTYPE_SERVER))
    For i = 0 To CANTIDADDECLANES - 1
        If Not guilds.Items(i) Is Nothing Then
            If guilds.Items(i).RevisarElecciones Then
                Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> " & guilds.Items(i).GetLeader & " es el nuevo lider de " & guilds.Items(i).GuildName & "!", FontTypeNames.FONTTYPE_SERVER))
            End If
        End If
proximo:
    Next i
    Call SendData(SendTarget.ToAll, 0, PrepareMessageConsoleMsg("Servidor> Elecciones revisadas", FontTypeNames.FONTTYPE_SERVER))
Exit Sub
errh:
    Call LogError("modGuilds.v_RutinaElecciones():" & Err.Description)
    Resume proximo
End Sub

Private Function GetGuildIndexFromChar(ByRef PlayerName As String) As Integer
'aca si que vamos a violar las capas deliveradamente ya que
'visual basic no permite declarar metodos de clase
Dim Temps   As String
    If InStrB(PlayerName, "\") <> 0 Then
        PlayerName = Replace(PlayerName, "\", vbNullString)
    End If
    If InStrB(PlayerName, "/") <> 0 Then
        PlayerName = Replace(PlayerName, "/", vbNullString)
    End If
    If InStrB(PlayerName, ".") <> 0 Then
        PlayerName = Replace(PlayerName, ".", vbNullString)
    End If
    Temps = GetByCampo("SELECT GuildIndex FROM pjs WHERE Nombre=" & Comillas(PlayerName), "GuildIndex")
    If IsNumeric(Temps) Then
        GetGuildIndexFromChar = CInt(Temps)
    Else
        GetGuildIndexFromChar = 0
    End If
End Function

Public Function GuildIndex(ByRef GuildName As String) As Integer
'me da el indice del guildname
Dim i As Integer

    GuildIndex = 0
    GuildName = UCase$(GuildName)
    For i = 0 To guilds.Count - 1
        If UCase$(guilds.Items(i).GuildName) = GuildName Then
            GuildIndex = guilds.Keys(i)
            Exit Function
        End If
    Next i
End Function

Public Function m_ListaDeMiembrosOnline(ByVal UserIndex As Integer, ByVal GuildIndex As Integer) As String
Dim i As Integer
    
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        While i > 0
            'No mostramos dioses y admins
            If i <> UserIndex And ((UserList(i).flags.Privilegios And (PlayerType.user Or PlayerType.Consejero Or PlayerType.SemiDios)) <> 0 Or (UserList(UserIndex).flags.Privilegios And (PlayerType.Dios Or PlayerType.Admin) <> 0)) Then _
                m_ListaDeMiembrosOnline = m_ListaDeMiembrosOnline & UserList(i).Name & ","
            i = guilds(GuildIndex).m_Iterador_ProximoUserIndex
        Wend
    End If
    If Len(m_ListaDeMiembrosOnline) > 0 Then
        m_ListaDeMiembrosOnline = left$(m_ListaDeMiembrosOnline, Len(m_ListaDeMiembrosOnline) - 1)
    End If
End Function

Public Function PrepareGuildsList() As String()
    Dim tStr() As String
    Dim i As Long
    
    If CANTIDADDECLANES = 0 Then
        ReDim tStr(0) As String
    Else
        ReDim tStr(CANTIDADDECLANES - 1) As String
        
        For i = 0 To CANTIDADDECLANES - 1
            tStr(i) = guilds.Items(i).GuildName
        Next i
    End If
    
    PrepareGuildsList = tStr
End Function

Public Sub SendGuildDetails(ByVal UserIndex As Integer, ByRef GuildName As String)
    Dim Codex(CANTIDADMAXIMACODEX - 1)  As String
    Dim GI      As Integer
    Dim i       As Long

    GI = GuildIndex(GuildName)
    If GI = 0 Then Exit Sub
    
    With guilds(GI)
        For i = 1 To CANTIDADMAXIMACODEX
            Codex(i - 1) = .GetCodex(i)
        Next i
        
        Call Protocol.WriteGuildDetails(UserIndex, GuildName, .Fundador, .GetFechaFundacion, .GetLeader, _
                                    .GetURL, .CantidadDeMiembros, .EleccionesAbiertas, Alineacion2String(.Alineacion), _
                                    .CantidadEnemys, .CantidadAllies, .PuntosAntifaccion & "/" & CStr(MAXANTIFACCION), _
                                    Codex, .GetDesc)
    End With
End Sub

Public Sub SendGuildLeaderInfo(ByVal UserIndex As Integer)
'***************************************************
'Autor: Mariano Barrou (El Oso)
'Last Modification: 12/10/06
'Las Modified By: Juan Martín Sotuyo Dodero (Maraxus)
'***************************************************
    Dim GI      As Integer
    Dim guildList() As String
    Dim MemberList() As String
    Dim aspirantsList() As String

    With UserList(UserIndex)
        GI = .GuildIndex
        
        guildList = PrepareGuildsList()
        
        If GI <= 0 Then
            'Send the guild list instead
            Call WriteGuildList(UserIndex, guildList)
            Exit Sub
        End If
        
        MemberList = guilds(GI).GetMemberList()
        
        If Not m_EsGuildLeader(.Name, GI) Then
            'Send the guild list instead
            Call WriteGuildMemberInfo(UserIndex, guildList, MemberList)
            Exit Sub
        End If
        
        aspirantsList = guilds(GI).GetAspirantes()
        
        Call WriteGuildLeaderInfo(UserIndex, guildList, MemberList, guilds(GI).GetGuildNews(), aspirantsList)
    End With
End Sub


Public Function m_Iterador_ProximoUserIndex(ByVal GuildIndex As Integer) As Integer
    'itera sobre los onlinemembers
    m_Iterador_ProximoUserIndex = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        m_Iterador_ProximoUserIndex = guilds(GuildIndex).m_Iterador_ProximoUserIndex()
    End If
End Function

Public Function Iterador_ProximoGM(ByVal GuildIndex As Integer) As Integer
    'itera sobre los gms escuchando este clan
    Iterador_ProximoGM = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        Iterador_ProximoGM = guilds(GuildIndex).Iterador_ProximoGM()
    End If
End Function

Public Function r_Iterador_ProximaPropuesta(ByVal GuildIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As Integer
    'itera sobre las propuestas
    r_Iterador_ProximaPropuesta = 0
    If GuildIndex > 0 And GuildIndex <= CANTIDADDECLANES Then
        r_Iterador_ProximaPropuesta = guilds(GuildIndex).Iterador_ProximaPropuesta(Tipo)
    End If
End Function

Public Function GMEscuchaClan(ByVal UserIndex As Integer, ByVal GuildName As String) As Integer
Dim GI As Integer

    'listen to no guild at all
    If LenB(GuildName) = 0 And UserList(UserIndex).EscucheClan <> 0 Then
        'Quit listening to previous guild!!
        Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
        guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
        Exit Function
    End If
    
'devuelve el guildindex
    GI = GuildIndex(GuildName)
    If GI > 0 Then
        If UserList(UserIndex).EscucheClan <> 0 Then
            If UserList(UserIndex).EscucheClan = GI Then
                'Already listening to them...
                Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
                GMEscuchaClan = GI
                Exit Function
            Else
                'Quit listening to previous guild!!
                Call WriteConsoleMsg(UserIndex, "Dejas de escuchar a : " & guilds(UserList(UserIndex).EscucheClan).GuildName, FontTypeNames.FONTTYPE_GUILD)
                guilds(UserList(UserIndex).EscucheClan).DesconectarGM (UserIndex)
            End If
        End If
        
        Call guilds(GI).ConectarGM(UserIndex)
        Call WriteConsoleMsg(UserIndex, "Conectado a : " & GuildName, FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = GI
        UserList(UserIndex).EscucheClan = GI
    Else
        Call WriteConsoleMsg(UserIndex, "Error, el clan no existe", FontTypeNames.FONTTYPE_GUILD)
        GMEscuchaClan = 0
    End If
    
End Function

Public Sub GMDejaDeEscucharClan(ByVal UserIndex As Integer, ByVal GuildIndex As Integer)
'el index lo tengo que tener de cuando me puse a escuchar
    UserList(UserIndex).EscucheClan = 0
    Call guilds(GuildIndex).DesconectarGM(UserIndex)
End Sub
Public Function r_DeclararGuerra(ByVal UserIndex As Integer, ByRef GuildGuerra As String, ByRef refError As String) As Integer
Dim GI  As Integer
Dim GIG As Integer

    r_DeclararGuerra = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildGuerra) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If
    
    GIG = GuildIndex(GuildGuerra)
    If guilds(GI).GetRelacion(GIG) = GUERRA Then
        refError = "Tu clan ya está en guerra con " & GuildGuerra & "."
        Exit Function
    End If
        
    If GI = GIG Then
        refError = "No puedes declarar la guerra a tu mismo clan"
        Exit Function
    End If

    If GIG < 1 Then
        Call LogError("ModGuilds.r_DeclararGuerra: " & GI & " declara a " & GuildGuerra)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.GUERRA)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.GUERRA)
    
    r_DeclararGuerra = GIG

End Function


Public Function r_AceptarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPaz As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPaz) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPaz)
    
    If GIG < 1 Then
        Call LogError("ModGuilds.r_AceptarPropuestaDePaz: " & GI & " acepta de " & GuildPaz)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.GUERRA Then
        refError = "No estás en guerra con ese clan"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay ninguna propuesta de paz para aceptar"
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.PAZ)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.PAZ)
    
    r_AceptarPropuestaDePaz = GIG
End Function

Public Function r_RechazarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Then
        Call LogError("ModGuilds.r_RechazarPropuestaDeAlianza: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, ALIADOS) Then
        refError = "No hay propuesta de alianza del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de alianza. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDeAlianza = GIG

End Function


Public Function r_RechazarPropuestaDePaz(ByVal UserIndex As Integer, ByRef GuildPro As String, ByRef refError As String) As Integer
'devuelve el index al clan guildPro
Dim GI      As Integer
Dim GIG     As Integer

    r_RechazarPropuestaDePaz = 0
    GI = UserList(UserIndex).GuildIndex
    
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildPro) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildPro)
    
    If GIG < 1 Then
        Call LogError("ModGuilds.r_RechazarPropuestaDePaz: " & GI & " acepta de " & GuildPro)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.PAZ) Then
        refError = "No hay propuesta de paz del clan " & GuildPro
        Exit Function
    End If
    
    Call guilds(GI).AnularPropuestas(GIG)
    'avisamos al otro clan
    Call guilds(GIG).SetGuildNews(guilds(GI).GuildName & " ha rechazado nuestra propuesta de paz. " & guilds(GIG).GetGuildNews())
    r_RechazarPropuestaDePaz = GIG

End Function


Public Function r_AceptarPropuestaDeAlianza(ByVal UserIndex As Integer, ByRef GuildAllie As String, ByRef refError As String) As Integer
'el clan de userindex acepta la propuesta de paz de guildpaz, con quien esta en guerra
Dim GI      As Integer
Dim GIG     As Integer

    r_AceptarPropuestaDeAlianza = 0
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    If Trim$(GuildAllie) = vbNullString Then
        refError = "No has seleccionado ningún clan"
        Exit Function
    End If

    GIG = GuildIndex(GuildAllie)
    
    If GIG < 1 Then
        Call LogError("ModGuilds.r_AceptarPropuestaDeAlianza: " & GI & " acepta de " & GuildAllie)
        refError = "Inconsistencia en el sistema de clanes. Avise a un administrador (GIG fuera de rango)"
        Exit Function
    End If

    If guilds(GI).GetRelacion(GIG) <> RELACIONES_GUILD.PAZ Then
        refError = "No estás en paz con el clan, solo puedes aceptar propuesas de alianzas con alguien que estes en paz."
        Exit Function
    End If
    
    If Not guilds(GI).HayPropuesta(GIG, RELACIONES_GUILD.ALIADOS) Then
        refError = "No hay ninguna propuesta de alianza para aceptar."
        Exit Function
    End If

    Call guilds(GI).AnularPropuestas(GIG)
    Call guilds(GIG).AnularPropuestas(GI)
    Call guilds(GI).SetRelacion(GIG, RELACIONES_GUILD.ALIADOS)
    Call guilds(GIG).SetRelacion(GI, RELACIONES_GUILD.ALIADOS)
    
    r_AceptarPropuestaDeAlianza = GIG

End Function


Public Function r_ClanGeneraPropuesta(ByVal UserIndex As Integer, ByRef OtroClan As String, ByVal Tipo As RELACIONES_GUILD, ByRef Detalle As String, ByRef refError As String) As Boolean
Dim OtroClanGI      As Integer
Dim GI              As Integer

    r_ClanGeneraPropuesta = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroClan)
    
    If OtroClanGI = GI Then
        refError = "No puedes declarar relaciones con tu propio clan"
        Exit Function
    End If
    
    If OtroClanGI <= 0 Then
        refError = "El sistema de clanes esta inconsistente, el otro clan no existe!"
        Exit Function
    End If
    
    If guilds(OtroClanGI).HayPropuesta(GI, Tipo) Then
        refError = "Ya hay propuesta de " & Relacion2String(Tipo) & " con " & OtroClan
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    'de acuerdo al tipo procedemos validando las transiciones
    If Tipo = RELACIONES_GUILD.PAZ Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.GUERRA Then
            refError = "No estás en guerra con " & OtroClan
            Exit Function
        End If
    ElseIf Tipo = RELACIONES_GUILD.GUERRA Then
        'por ahora no hay propuestas de guerra
    ElseIf Tipo = RELACIONES_GUILD.ALIADOS Then
        If guilds(GI).GetRelacion(OtroClanGI) <> RELACIONES_GUILD.PAZ Then
            refError = "Para solicitar alianza no debes estar ni aliado ni en guerra con " & OtroClan
            Exit Function
        End If
    End If
    
    Call guilds(OtroClanGI).SetPropuesta(Tipo, GI, Detalle)
    r_ClanGeneraPropuesta = True

End Function

Public Function r_VerPropuesta(ByVal UserIndex As Integer, ByRef OtroGuild As String, ByVal Tipo As RELACIONES_GUILD, ByRef refError As String) As String
Dim OtroClanGI      As Integer
Dim GI              As Integer
    
    r_VerPropuesta = vbNullString
    refError = vbNullString
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No eres miembro de ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    OtroClanGI = GuildIndex(OtroGuild)
    
    If Not guilds(GI).HayPropuesta(OtroClanGI, Tipo) Then
        refError = "No existe la propuesta solicitada"
        Exit Function
    End If
    
    r_VerPropuesta = guilds(GI).GetPropuesta(OtroClanGI, Tipo)
    
End Function

Public Function r_ListaDePropuestas(ByVal UserIndex As Integer, ByVal Tipo As RELACIONES_GUILD) As String()

    Dim GI  As Integer
    Dim i   As Integer
    Dim proposalCount As Integer
    Dim proposals() As String
    Dim P As Integer
    
    GI = UserList(UserIndex).GuildIndex
    
    If GI > 0 Then
        With guilds(GI)
            proposalCount = .CantidadPropuestas(Tipo)
            
            'Resize array to contain all proposals
            If proposalCount > 0 Then
                ReDim proposals(proposalCount - 1) As String
            Else
                ReDim proposals(0) As String
            End If
            
            'Store each guild name
            For i = 0 To proposalCount - 1
                P = .Iterador_ProximaPropuesta(Tipo)
                proposals(i) = guilds(P).GuildName
            Next i
        End With
    End If
    
    r_ListaDePropuestas = proposals
End Function

Public Sub a_RechazarAspiranteChar(ByRef Aspirante As String, ByVal guild As Integer, ByRef Detalles As String)
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    Call guilds(guild).InformarRechazoEnChar(Aspirante, Detalles)
End Sub

Public Function a_ObtenerRechazoDeChar(ByRef Aspirante As String) As String
    If InStrB(Aspirante, "\") <> 0 Then
        Aspirante = Replace(Aspirante, "\", "")
    End If
    If InStrB(Aspirante, "/") <> 0 Then
        Aspirante = Replace(Aspirante, "/", "")
    End If
    If InStrB(Aspirante, ".") <> 0 Then
        Aspirante = Replace(Aspirante, ".", "")
    End If
    a_ObtenerRechazoDeChar = GetByCampo("SELECT MotivoRechazo FROM pjs WHERE Nombre=" & Comillas(Aspirante), "MotivoRechazo")
    Execute ("UPDATE pjs SET MotivoRechazo='' WHERE Nombre=" & Comillas(Aspirante))
End Function

Public Function a_RechazarAspirante(ByVal UserIndex As Integer, ByRef Nombre As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer

    a_RechazarAspirante = False
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If

    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)

    If NroAspirante = 0 Then
        refError = Nombre & " no es aspirante a tu clan"
        Exit Function
    End If

    Call guilds(GI).RetirarAspirante(Nombre, NroAspirante)
    refError = "Fue rechazada tu solicitud de ingreso a " & guilds(GI).GuildName
    a_RechazarAspirante = True

End Function

Public Function a_DetallesAspirante(ByVal UserIndex As Integer, ByRef Nombre As String) As String
Dim GI              As Integer
Dim NroAspirante    As Integer

    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Nombre)
    If NroAspirante > 0 Then
        a_DetallesAspirante = guilds(GI).DetallesSolicitudAspirante(NroAspirante)
    End If
    
End Function

Public Sub SendDetallesPersonaje(ByVal UserIndex As Integer, ByVal Personaje As String)
    Dim GI          As Integer
    Dim NroAsp      As Integer
    Dim GuildName   As String
    Dim Miembro     As String
    Dim GuildActual As Integer
    Dim list()      As String
    Dim i           As Long
    
    GI = UserList(UserIndex).GuildIndex
    
    Personaje = UCase$(Personaje)
    
    If GI <= 0 Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No perteneces a ningún clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        Call Protocol.WriteConsoleMsg(UserIndex, "No eres el líder de tu clan", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If InStrB(Personaje, "\") <> 0 Then
        Personaje = Replace$(Personaje, "\", vbNullString)
    End If
    If InStrB(Personaje, "/") <> 0 Then
        Personaje = Replace$(Personaje, "/", vbNullString)
    End If
    If InStrB(Personaje, ".") <> 0 Then
        Personaje = Replace$(Personaje, ".", vbNullString)
    End If
    
    NroAsp = guilds(GI).NumeroDeAspirante(Personaje)
    
    If NroAsp = 0 Then
        list = guilds(GI).GetMemberList()
        
        For i = 0 To UBound(list())
            If UCase$(Personaje) = UCase$(list(i)) Then Exit For
        Next i
        
        If i > UBound(list()) Then
            Call Protocol.WriteConsoleMsg(UserIndex, "El personaje no es ni aspirante ni miembro del clan", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    'ahora traemos la info
    
    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT GuildIndex, Miembro, Raza, Clase, Genero, ELV, GLD, Banco, Rep_Promedio, Pedidos, EjercitoReal, EjercitoCaos, CiudMatados, CrimMatados FROM pjs WHERE Nombre=" & Comillas(Personaje), Datos)


        
        ' Get the character's current guild
        GuildActual = val(Datos("GuildIndex"))
        If GuildActual > 0 And GuildActual <= CANTIDADDECLANES Then
            GuildName = "<" & guilds(GuildActual).GuildName & ">"
        Else
            GuildName = "Ninguno"
        End If
        
        'Get previous guilds
        If IsNull(Datos("Miembro")) Then
            Miembro = ""
        Else
            Miembro = Datos("Miembro")
        End If
        If Len(Miembro) > 400 Then
            Miembro = ".." & Right$(Miembro, 398)
        End If
        
        Call Protocol.WriteCharacterInfo(UserIndex, Personaje, Datos("Raza"), Datos("Clase"), _
                                Datos("Genero"), Datos("ELV"), Datos("GLD"), _
                                Datos("Banco"), Datos("Rep_Promedio"), Datos("Pedidos"), _
                                GuildName, Miembro, Datos("EjercitoReal"), Datos("EjercitoCaos"), _
                                Datos("CiudMatados"), Datos("CrimMatados"))


End Sub

Public Function a_NuevoAspirante(ByVal UserIndex As Integer, ByRef Clan As String, ByRef Solicitud As String, ByRef refError As String) As Boolean
Dim ViejoSolicitado     As String
Dim ViejoGuildINdex     As Integer
Dim ViejoNroAspirante   As Integer
Dim NuevoGuildIndex     As Integer

    a_NuevoAspirante = False

    If UserList(UserIndex).GuildIndex > 0 Then
        refError = "Ya perteneces a un clan, debes salir del mismo antes de solicitar ingresar a otro"
        Exit Function
    End If
    
    If EsNewbie(UserIndex) Then
        refError = "Los newbies no tienen derecho a entrar a un clan."
        Exit Function
    End If

    NuevoGuildIndex = GuildIndex(Clan)
    If NuevoGuildIndex = 0 Then
        refError = "Ese clan no existe! Avise a un administrador."
        Exit Function
    End If
    
    If Not m_EstadoPermiteEntrar(UserIndex, NuevoGuildIndex) Then
        refError = "Tu no puedes entrar a un clan de alineación " & Alineacion2String(guilds(NuevoGuildIndex).Alineacion)
        Exit Function
    End If

    If guilds(NuevoGuildIndex).CantidadAspirantes >= MAXASPIRANTES Then
        refError = "El clan tiene demasiados aspirantes. Contáctate con un miembro para que procese las solicitudes."
        Exit Function
    End If

    ViejoSolicitado = GetByCampo("SELECT AspiranteA FROM pjs WHERE Id=" & UserList(UserIndex).MySQLId, "AspiranteA")

    If LenB(ViejoSolicitado) <> 0 Then
        'borramos la vieja solicitud
        ViejoGuildINdex = CInt(ViejoSolicitado)
        If ViejoGuildINdex <> 0 Then
            ViejoNroAspirante = guilds(ViejoGuildINdex).NumeroDeAspirante(UserList(UserIndex).Name)
            If ViejoNroAspirante > 0 Then
                Call guilds(ViejoGuildINdex).RetirarAspirante(UserList(UserIndex).Name, ViejoNroAspirante)
            End If
        Else
            'RefError = "Inconsistencia en los clanes, avise a un administrador"
            'Exit Function
        End If
    End If
    
    Call guilds(NuevoGuildIndex).NuevoAspirante(UserList(UserIndex).Name, Solicitud)
    a_NuevoAspirante = True
End Function

Public Function a_AceptarAspirante(ByVal UserIndex As Integer, ByRef Aspirante As String, ByRef refError As String) As Boolean
Dim GI              As Integer
Dim NroAspirante    As Integer
Dim AspiranteUI     As Integer

    'un pj ingresa al clan :D

    a_AceptarAspirante = False
    
    GI = UserList(UserIndex).GuildIndex
    If GI <= 0 Then
        refError = "No perteneces a ningún clan"
        Exit Function
    End If
    
    If Not m_EsGuildLeader(UserList(UserIndex).Name, GI) Then
        refError = "No eres el líder de tu clan"
        Exit Function
    End If
    
    NroAspirante = guilds(GI).NumeroDeAspirante(Aspirante)
    
    If NroAspirante = 0 Then
        refError = "El Pj no es aspirante al clan"
        Exit Function
    End If
    
    AspiranteUI = NameIndex(Aspirante)
    If AspiranteUI > 0 Then
        'pj Online
        If Not m_EstadoPermiteEntrar(AspiranteUI, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf Not UserList(AspiranteUI).GuildIndex = 0 Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    Else
        If Not m_EstadoPermiteEntrarChar(Aspirante, GI) Then
            refError = Aspirante & " no puede entrar a un clan " & Alineacion2String(guilds(GI).Alineacion)
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        ElseIf GetGuildIndexFromChar(Aspirante) Then
            refError = Aspirante & " ya es parte de otro clan."
            Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
            Exit Function
        End If
    End If
    'el pj es aspirante al clan y puede entrar
    
    Call guilds(GI).RetirarAspirante(Aspirante, NroAspirante)
    Call guilds(GI).AceptarNuevoMiembro(Aspirante)
    
    ' If player is online, update tag
    If AspiranteUI > 0 Then
        Call RefreshCharStatus(AspiranteUI)
    End If
    
    a_AceptarAspirante = True
End Function

Public Function GuildName(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Then _
        Exit Function
    
    GuildName = guilds(GuildIndex).GuildName
End Function

Public Function GuildLeader(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Then _
        Exit Function
    
    GuildLeader = guilds(GuildIndex).GetLeader
End Function

Public Function GuildAlignment(ByVal GuildIndex As Integer) As String
    If GuildIndex <= 0 Then _
        Exit Function
    
    GuildAlignment = Alineacion2String(guilds(GuildIndex).Alineacion)
End Function

Public Function GuildFounder(ByVal GuildIndex As Integer) As String
'***************************************************
'Autor: ZaMa
'Returns the guild founder's name
'Last Modification: 25/03/2009
'***************************************************
    If GuildIndex <= 0 Then _
        Exit Function
    
    GuildFounder = guilds(GuildIndex).Fundador
End Function

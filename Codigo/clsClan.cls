VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************
' clsClan.cls
'
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

''
' clase clan
'
' Es el "ADO" de los clanes. La interfaz entre el disco y
' el juego. Los datos no se guardan en memoria
' para evitar problemas de sincronizacion, y considerando
' que la performance de estas rutinas NO es critica.
' by el oso :p

Private p_GuildName                 As String

    Dim p_Founder As String
    Dim p_Leader As String
    Dim p_Fecha As Date
    Dim p_Desc As String
    Dim p_Antifaccion As Long
    Dim p_GuildNews As String
    Dim p_URL As String
    Dim p_Codex(1 To 8) As String
    Dim p_CantMiembros As Integer


Private p_Alineacion                As ALINEACION_GUILD
Private p_OnlineMembers             As Collection   'Array de UserIndexes!
Private p_GMsOnline                 As Collection
Private p_PropuestasDePaz           As Collection
Private p_PropuestasDeAlianza       As Collection
Private p_IteradorRelaciones        As Integer
Private p_IteradorOnlineMembers     As Integer
Private p_IteradorPropuesta         As Integer
Private p_IteradorOnlineGMs         As Integer
Private p_GuildNumber               As Integer      'Numero de guild en el mundo
Private p_Relaciones              As Dictionary

Private Const NEWSLENGTH = 1024
Private Const DESCLENGTH = 256
Private Const CODEXLENGTH = 256

Public Property Get GuildName() As String
    GuildName = p_GuildName
End Property


'
'ALINEACION Y ANTIFACCION
'

Public Property Get Alineacion() As ALINEACION_GUILD
    Alineacion = p_Alineacion
End Property

Public Property Get PuntosAntifaccion() As Integer
    PuntosAntifaccion = p_Antifaccion
End Property

Public Property Let PuntosAntifaccion(ByVal P As Integer)
    p_Antifaccion = P
    Call Execute("UPDATE clanes SET Antifaccion=" & P & " WHERE Id=" & p_GuildNumber)
End Property

Public Sub CambiarAlineacion(ByVal NuevaAlineacion As ALINEACION_GUILD)
    p_Alineacion = NuevaAlineacion
    Call Execute("UPDATE clanes SET Alineacion=" & NuevaAlineacion & " WHERE Id=" & p_GuildNumber)
End Sub

'
'INICIALIZADORES
'

Private Sub Class_Initialize()

End Sub

Private Sub Class_Terminate()
    Set p_OnlineMembers = Nothing
    Set p_GMsOnline = Nothing
    Set p_PropuestasDePaz = Nothing
    Set p_PropuestasDeAlianza = Nothing
End Sub



Public Sub Inicializar(ByRef Datos As clsMySQLRecordSet)
Dim i As Integer

    p_GuildName = Datos("GuildName")
    p_GuildNumber = Datos("Id")
    p_Alineacion = Datos("Alineacion")
    
    
    p_Founder = Datos("founder")
    p_Leader = Datos("leader")
    p_Fecha = Datos("Fecha")
    p_Desc = Datos("desc")
    p_Antifaccion = Datos("Antifaccion")
    p_GuildNews = Datos("guildNews")
    p_URL = Datos("URL")
    p_Codex(1) = Datos("Codex1")
    p_Codex(2) = Datos("Codex2")
    p_Codex(3) = Datos("Codex3")
    p_Codex(4) = Datos("Codex4")
    p_Codex(5) = Datos("Codex5")
    p_Codex(6) = Datos("Codex6")
    p_Codex(7) = Datos("Codex7")
    p_Codex(8) = Datos("Codex8")
    p_CantMiembros = Datos("CantMiembros")
    
    Set p_OnlineMembers = New Collection
    Set p_GMsOnline = New Collection
    Set p_PropuestasDePaz = New Collection
    Set p_PropuestasDeAlianza = New Collection
    Set p_Relaciones = New Dictionary

    p_IteradorOnlineMembers = 0
    p_IteradorPropuesta = 0
    p_IteradorOnlineGMs = 0
    p_IteradorRelaciones = 0
    

    Dim DatosRel As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT IdClanTo,Relacion FROM clanes_relaciones WHERE IdClan=" & p_GuildNumber, DatosRel)
    
    If Cant > 0 Then
    While Not DatosRel.EOF
        Call p_Relaciones.Add(CInt(DatosRel!IdClanTo), CInt(DatosRel!Relacion))
        DatosRel.MoveNext
    Wend
    End If
    
    
     Cant = mySQL.SQLQuery("SELECT IdClan,Tipo FROM clanes_propuestas WHERE IdClanTo=" & p_GuildNumber, DatosRel)
    If Cant > 0 Then
    While Not DatosRel.EOF
        Select Case DatosRel!Tipo
            Case RELACIONES_GUILD.ALIADOS
                p_PropuestasDeAlianza.Add DatosRel!IdClan
            Case RELACIONES_GUILD.PAZ
                p_PropuestasDePaz.Add DatosRel!IdClan
        End Select
        DatosRel.MoveNext
    Wend
    End If
End Sub

' @param Fundador Nombre del fundador del clan
'
Public Function InicializarNuevoClan(ByVal GuildName As String, ByVal Fundador As String, ByVal Alineacion As ALINEACION_GUILD, ByVal desc As String, ByVal URL As String, ByRef Codex() As String) As Integer
    
    Execute ("INSERT INTO clanes (Founder, GuildName, Fecha, Antifaccion, Alineacion, `Desc`, GuildNews, URL, Leader, CantMiembros, Codex1,Codex2,Codex3,Codex4,Codex5,Codex6,Codex7,Codex8) VALUES (" & _
             Comillas(Fundador) & "," & Comillas(GuildName) & ",NOW(),0," & Alineacion & "," & Comillas(desc) & "," & Comillas("Clan creado con alineación : " & Alineacion2String(Alineacion)) & "," & Comillas(URL) & "," & Comillas(Fundador) & ",0," & Comillas(Codex(0)) & "," & Comillas(Codex(1)) & "," & Comillas(Codex(2)) & "," & Comillas(Codex(3)) & "," & Comillas(Codex(4)) & "," & Comillas(Codex(5)) & "," & Comillas(Codex(6)) & "," & Comillas(Codex(7)) & ")")
    
    InicializarNuevoClan = GetByCampo("SELECT LAST_INSERT_ID() as L", "L")
End Function
'
'MEMBRESIAS
'

Public Property Get Fundador() As String
    Fundador = p_Founder
End Property

'Public Property Get JugadoresOnline() As String
'Dim i As Integer
'    'leve violacion de capas x aqui, je
'    For i = 1 To p_OnlineMembers.Count
'        JugadoresOnline = UserList(p_OnlineMembers.Item(i)).Name & "," & JugadoresOnline
'    Next i
'End Property

Public Property Get CantidadDeMiembros() As Integer
    CantidadDeMiembros = p_CantMiembros
End Property

Public Sub SetLeader(ByRef leader As String)
    p_Leader = leader
    Execute ("UPDATE clanes SET Leader=" & Comillas(leader) & " WHERE Id=" & p_GuildNumber)
End Sub

Public Function GetLeader() As String
    GetLeader = p_Leader
End Function

Public Function GetMemberList() As String()
    Dim OldQ    As Integer
    Dim list() As String
    Dim i As Long
      
    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT Nombre FROM pjs WHERE GuildIndex=" & p_GuildNumber & " ORDER BY Nombre ASC", Datos)
    

    If Cant > 1 Then
        ReDim list(Cant - 1) As String
    Else
        ReDim list(0) As String
    End If
    
    For i = 0 To Cant - 1
        list(i) = Datos("Nombre")
        Datos.MoveNext
    Next i
    
    GetMemberList = list
End Function

Public Sub ConectarMiembro(ByVal UserIndex As Integer)
    p_OnlineMembers.Add UserIndex
End Sub

Public Sub DesConectarMiembro(ByVal UserIndex As Integer)
Dim i   As Integer
    For i = 1 To p_OnlineMembers.Count
        If p_OnlineMembers.Item(i) = UserIndex Then
            p_OnlineMembers.Remove i
            Exit Sub
        End If
    Next i
End Sub

Public Sub AceptarNuevoMiembro(ByRef Nombre As String)
Dim OldQ    As Integer
Dim OldQs   As String

    If PersonajeExiste(Nombre) Then
        Execute ("UPDATE pjs SET GuildIndex=" & p_GuildNumber & ", AspiranteA=0 WHERE Nombre=" & Comillas(Nombre))
        
        p_CantMiembros = p_CantMiembros + 1
        Execute ("UPDATE clanes SET CantMiembros=CantMiembros+1 WHERE Id=" & p_GuildNumber)
    End If

End Sub

Public Sub ExpulsarMiembro(ByRef Nombre As String)
Dim OldQ        As Integer
Dim Temps       As String
Dim i           As Integer
Dim EsMiembro   As Boolean
Dim MiembroDe   As String

    If PersonajeExiste(Nombre) Then
        Nombre = UCase$(Nombre)

        EsMiembro = CIntNull(GetByCampo("SELECT COUNT(Id) as Cant FROM pjs WHERE Nombre=" & Comillas(Nombre) & " AND GuildIndex=" & p_GuildNumber, "Cant")) > 0
        
        If EsMiembro Then
            'lo echo a el
            Execute ("UPDATE clanes SET CantMiembros=CantMiembros-1 WHERE Id=" & p_GuildNumber)
            
            MiembroDe = CStrNull(GetByCampo("SELECT Miembro FROM pjs WHERE Nombre=" & Comillas(Nombre), "Miembro"))
            If Not InStr(1, MiembroDe, p_GuildName, vbTextCompare) > 0 Then
                If LenB(MiembroDe) <> 0 Then
                    MiembroDe = MiembroDe & ","
                End If
                MiembroDe = ", Miembro=" & Comillas(MiembroDe & p_GuildName)
            Else
                MiembroDe = ""
            End If
            Execute ("UPDATE pjs SET GuildIndex=0" & MiembroDe & " WHERE Nombre=" & Comillas(Nombre))
        End If

    End If

End Sub

'
'ASPIRANTES
'

Public Function GetAspirantes() As String()
    Dim OldQ    As Integer
    Dim list() As String
    Dim i As Long
    
    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT Nombre FROM clanes_solicitudes WHERE IdClan=" & p_GuildNumber, Datos)
 
    If Cant > 1 Then
        ReDim list(Cant - 1) As String
    Else
        ReDim list(0) As String
    End If
    
    For i = 0 To Cant - 1
        list(i) = Datos("Nombre")
        Datos.MoveNext
    Next i
    
    GetAspirantes = list
End Function

Public Function CantidadAspirantes() As Integer
Dim Cant
Cant = GetByCampo("SELECT COUNT(Id) as Cant FROM clanes_solicitudes WHERE IdClan=" & p_GuildNumber, "Cant")
CantidadAspirantes = Cant
End Function

Public Function DetallesSolicitudAspirante(ByVal NroAspirante As Integer) As String
    DetallesSolicitudAspirante = GetByCampo("SELECT Solicitud FROM clanes_solicitudes WHERE Id=" & NroAspirante, "Solicitud")
End Function

Public Function NumeroDeAspirante(ByRef Nombre As String) As Integer
    
    NumeroDeAspirante = CIntNull(GetByCampo("SELECT Id FROM clanes_solicitudes WHERE Nombre=" & Comillas(Nombre), "Id"))
    
End Function

Public Sub NuevoAspirante(ByRef Nombre As String, ByRef Peticion As String)

    Peticion = IIf(Trim$(Peticion) = vbNullString, "Peticion vacia", Peticion)
    If Len(Peticion) > 255 Then Peticion = left(Peticion, 255)
    Execute ("INSERT INTO clanes_solicitudes (IdClan, Nombre, Solicitud) VALUES (" & p_GuildNumber & "," & Comillas(Nombre) & "," & Comillas(Peticion) & ")")
    Execute ("UPDATE pjs SET AspiranteA=" & p_GuildNumber & " WHERE Nombre=" & Comillas(Nombre))

End Sub

Public Sub RetirarAspirante(ByRef Nombre As String, ByRef NroAspirante As Integer)

Dim Pedidos As String

    
    Pedidos = GetByCampo("SELECT Pedidos FROM pjs WHERE Nombre=" & Comillas(Nombre), "Pedidos")
    If Not InStr(1, Pedidos, p_GuildName, vbTextCompare) > 0 Then
        If LenB(Pedidos) <> 0 Then
            Pedidos = Pedidos & ","
        End If
        Pedidos = Pedidos & p_GuildName
    End If
    Execute ("UPDATE pjs SET AspiranteA=0, Pedidos=" & Comillas(Pedidos) & " WHERE Nombre=" & Comillas(Nombre))
    
    Execute ("DELETE FROM clanes_solicitudes WHERE Id=" & NroAspirante)

End Sub

Public Sub InformarRechazoEnChar(ByRef Nombre As String, ByRef Detalles As String)
    Execute ("UPDATE pjs SET MotivoRechazo=" & Comillas(Detalles) & " WHERE Nombre=" & Comillas(Nombre))
End Sub

'
'DEFINICION DEL CLAN (CODEX Y NOTICIAS)
'

Public Function GetFechaFundacion() As Date
    GetFechaFundacion = p_Fecha
End Function

Public Sub SetCodex(ByVal CodexNumber As Integer, ByRef Codex As String)
    Call ReplaceInvalidChars(Codex)
    Codex = left$(Codex, CODEXLENGTH)
    p_Codex(CodexNumber) = Codex
    Execute ("UPDATE clanes SET Codex" & CodexNumber & "=" & Comillas(Codex) & " WHERE Id=" & p_GuildNumber)
End Sub

Public Function GetCodex(ByVal CodexNumber As Integer) As String
    GetCodex = p_Codex(CodexNumber)
End Function


Public Sub SetURL(ByRef URL As String)
    p_URL = left$(URL, 80)
    Execute ("UPDATE clanes SET URL=" & Comillas(p_URL) & " WHERE Id=" & p_GuildNumber)
End Sub

Public Function GetURL() As String
    GetURL = p_URL
End Function

Public Sub SetGuildNews(ByRef News As String)
    Call ReplaceInvalidChars(News)
    
    News = left$(News, NEWSLENGTH)
    
    p_GuildNews = News
    
    Execute ("UPDATE clanes SET GuildNews=" & Comillas(News) & " WHERE Id=" & p_GuildNumber)
    
End Sub

Public Function GetGuildNews() As String
    GetGuildNews = p_GuildNews
End Function

Public Sub SetDesc(ByRef desc As String)
    Call ReplaceInvalidChars(desc)
    desc = left$(desc, DESCLENGTH)
    p_Desc = desc
    Execute ("UPDATE clanes SET Desc=" & Comillas(desc) & " WHERE Id=" & p_GuildNumber)

End Sub

Public Function GetDesc() As String
    GetDesc = p_Desc
End Function



'
'
'ELECCIONES
'
'

Public Function EleccionesAbiertas() As Boolean
    EleccionesAbiertas = CIntNull(GetByCampo("SELECT COUNT(*) as Cant FROM clanes WHERE Id=" & p_GuildNumber & " AND EleccionesFinalizan IS NULL", "Cant")) = 0
End Function

Public Sub AbrirElecciones()
    Execute ("UPDATE clanes SET EleccionesFinalizan=DATE_ADD(NOW(), INTERVAL 1 DAY) WHERE Id=" & p_GuildNumber)
End Sub

Private Sub CerrarElecciones()  'solo pueden cerrarse mediante recuento de votos
    Execute ("UPDATE clanes SET EleccionesFinalizan=NULL WHERE Id=" & p_GuildNumber)
    Execute ("DELETE FROM clanes_votos WHERE IdClan=" & p_GuildNumber)
End Sub

Public Sub ContabilizarVoto(ByRef Votante As String, ByRef Votado As String)
    Execute ("INSERT INTO clanes_votos (IdClan,Nombre,Voto) VALUES (" & p_GuildNumber & "," & Comillas(Votante) & "," & Comillas(Votado) & ") ON DUPLICATE KEY UPDATE Voto=" & Comillas(Votado) & ";")
End Sub

Public Function YaVoto(ByVal Votante As String) As Boolean
    YaVoto = CIntNull(GetByCampo("SELECT COUNT(*) as Cant FROM clanes_votos WHERE IdClan=" & p_GuildNumber & " AND Nombre=" & Comillas(Votante), "Cant")) > 0
End Function

Private Function ContarVotos(ByRef CantGanadores As Integer) As String
Dim i           As Integer

Dim Cant As Integer

On Error GoTo errh

    Dim Datos As clsMySQLRecordSet
    Dim CantR As Long
    CantR = mySQL.SQLQuery("SELECT COUNT(*) as 'Cant', Voto FROM clanes_votos WHERE IdClan=" & p_GuildNumber & " GROUP BY Voto ORDER BY Cant DESC", Datos)

    If CantR = 0 Then
        CantGanadores = 0
        ContarVotos = vbNullString
    Else
        Cant = Datos("Cant")
        ContarVotos = Datos("Voto")
        CantGanadores = 1
        For i = 2 To CantR
            Datos.MoveNext
            If Cant = Datos("Cant") Then
                ContarVotos = ContarVotos & ", " & Datos("Voto")
                CantGanadores = CantGanadores + 1
            Else
                Exit For
            End If
        Next i
    End If

    
Exit Function
errh:
    LogError ("clsClan.Contarvotos: " & Err.Description)
    ContarVotos = vbNullString
End Function

Public Function RevisarElecciones() As Boolean
Dim FechaSufragio   As Date
Dim Temps
Dim Ganador         As String
Dim CantGanadores   As Integer
Dim list()          As String
Dim i               As Long

    RevisarElecciones = False
    Temps = GetByCampo("SELECT EleccionesFinalizan FROM clanes WHERE Id=" & p_GuildNumber, "EleccionesFinalizan")
    
    If IsNull(Temps) Or Trim(Temps) = "" Then Exit Function
    
    If IsDate(Temps) Then
        FechaSufragio = CDate(Temps)
        If FechaSufragio < Now Then     'toca!
            Ganador = ContarVotos(CantGanadores)

            If CantGanadores > 1 Then
                'empate en la votacion
                Call SetGuildNews("*Empate en la votación. " & Ganador & " empataron en votos en las elecciones del clan.")
            ElseIf CantGanadores = 1 Then
                list = Me.GetMemberList()
                
                For i = 0 To UBound(list())
                    If UCase$(Ganador) = UCase$(list(i)) Then Exit For
                Next i
                
                If i <= UBound(list()) Then
                    Call SetGuildNews("*" & Ganador & " ganó la elección del clan*")
                    Call Me.SetLeader(Ganador)
                    RevisarElecciones = True
                Else
                    Call SetGuildNews("*" & Ganador & " ganó la elección del clan pero abandonó las filas por lo que la votación queda desierta*")
                End If
            Else
                Call SetGuildNews("*El período de votación se cerró sin votos*")
            End If
            
            Call CerrarElecciones
            
        End If
    Else
        Call LogError("clsClan.RevisarElecciones: tempS is not Date")
    End If

End Function

'/VOTACIONES


'
'RELACIONES
'

Public Property Get CantidadPropuestas(ByVal Tipo As RELACIONES_GUILD) As Integer
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            CantidadPropuestas = p_PropuestasDeAlianza.Count
        Case RELACIONES_GUILD.GUERRA
        
        Case RELACIONES_GUILD.PAZ
            CantidadPropuestas = p_PropuestasDePaz.Count
    End Select
End Property

Public Property Get CantidadEnemys() As Integer
Dim i As Integer
    For i = 0 To p_Relaciones.Count - 1
        CantidadEnemys = CantidadEnemys + IIf(p_Relaciones.Items(i) = RELACIONES_GUILD.GUERRA, 1, 0)
    Next i
End Property

Public Property Get CantidadAllies() As Integer
Dim i As Integer
    For i = 0 To p_Relaciones.Count - 1
        CantidadAllies = CantidadEnemys + IIf(p_Relaciones.Items(i) = RELACIONES_GUILD.ALIADOS, 1, 0)
    Next i
End Property

Public Function GetRelacion(ByVal OtroGuild As Integer) As RELACIONES_GUILD
If p_Relaciones.Exists(OtroGuild) Then
    GetRelacion = p_Relaciones(OtroGuild)
Else
    GetRelacion = RELACIONES_GUILD.PAZ
End If
End Function

Public Sub SetRelacion(ByVal GuildIndex As Integer, ByVal Relacion As RELACIONES_GUILD)
If p_Relaciones.Exists(GuildIndex) Then
    p_Relaciones(GuildIndex) = Relacion
Else
    Call p_Relaciones.Add(GuildIndex, Relacion)
End If

Execute ("INSERT INTO clanes_relaciones (IdClan,IdClanTo,Relacion) VALUES (" & p_GuildNumber & "," & GuildIndex & "," & Relacion & ") ON DUPLICATE KEY UPDATE Relacion=" & Relacion & ";")
End Sub

Public Sub SetPropuesta(ByVal Tipo As RELACIONES_GUILD, ByVal OtroGuild As Integer, ByRef Detalle As String)
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            p_PropuestasDeAlianza.Add OtroGuild
        Case RELACIONES_GUILD.PAZ
            p_PropuestasDePaz.Add OtroGuild
    End Select
    If Len(Detalle) > 255 Then Detalle = left$(Detalle, 255)
    Execute ("INSERT INTO clanes_propuestas (IdClan, IdClanTo, Detalle, Tipo) VALUES (" & p_GuildNumber & "," & OtroGuild & "," & Comillas(Detalle) & "," & Tipo & ")")
End Sub

Public Sub AnularPropuestas(ByVal OtroGuild As Integer)
Dim i As Integer

    Execute ("DELETE FROM clanes_propuestas WHERE IdClan=" & p_GuildNumber & " AND IdClanTo=" & OtroGuild)
    For i = 1 To p_PropuestasDePaz.Count
        If p_PropuestasDePaz.Item(i) = OtroGuild Then p_PropuestasDePaz.Remove (i)
        Exit Sub
    Next i
    For i = 1 To p_PropuestasDeAlianza.Count
        If p_PropuestasDeAlianza.Item(i) = OtroGuild Then p_PropuestasDeAlianza.Remove (i)
        Exit Sub
    Next i
    
End Sub

Public Function GetPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As String
'trae la solicitd que haya, no valida si es actual o de que tipo es
    Dim Datos As clsMySQLRecordSet
    Dim Cant As Long
    Cant = mySQL.SQLQuery("SELECT Detalle, Tipo FROM clanes_propuestas WHERE IdClan=" & p_GuildNumber & " AND IdClanTo=" & OtroGuild, Datos)
    
    GetPropuesta = Datos("Detalle")
    Tipo = CInt(Datos("Tipo"))
End Function

Public Function HayPropuesta(ByVal OtroGuild As Integer, ByRef Tipo As RELACIONES_GUILD) As Boolean
Dim i As Integer

    HayPropuesta = False
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            For i = 1 To p_PropuestasDeAlianza.Count
                If p_PropuestasDeAlianza.Item(i) = OtroGuild Then
                    HayPropuesta = True
                End If
            Next i
        Case RELACIONES_GUILD.PAZ
            For i = 1 To p_PropuestasDePaz.Count
                If p_PropuestasDePaz.Item(i) = OtroGuild Then
                    HayPropuesta = True
                End If
            Next i
        Case RELACIONES_GUILD.GUERRA

    End Select

End Function

'Public Function GetEnemy(ByVal EnemyIndex As Integer) As String
'    GetEnemy = GetVar(ENEMIESFILE, "ENEMYS", "ENEMY" & EnemyIndex)
'End Function

'Public Function GetAllie(ByVal AllieIndex As Integer) As String
'    GetAllie = GetVar(ALLIESFILE, "ALLIES", "ALLIE" & AllieIndex)
'End Function


'
'ITERADORES
'

Public Function Iterador_ProximaPropuesta(ByVal Tipo As RELACIONES_GUILD) As Integer
    
    Iterador_ProximaPropuesta = 0
    Select Case Tipo
        Case RELACIONES_GUILD.ALIADOS
            If p_IteradorPropuesta >= p_PropuestasDeAlianza.Count Then
                p_IteradorPropuesta = 0
            End If
            p_IteradorPropuesta = p_IteradorPropuesta + 1
            Iterador_ProximaPropuesta = p_PropuestasDeAlianza.Item(p_IteradorPropuesta)
        Case RELACIONES_GUILD.PAZ
            If p_IteradorPropuesta >= p_PropuestasDePaz.Count Then
                p_IteradorPropuesta = 0
            End If
            p_IteradorPropuesta = p_IteradorPropuesta + 1
            Iterador_ProximaPropuesta = p_PropuestasDePaz.Item(p_IteradorPropuesta)
    End Select
    
End Function

Public Function m_Iterador_ProximoUserIndex() As Integer
    
    If p_IteradorOnlineMembers < p_OnlineMembers.Count Then
        p_IteradorOnlineMembers = p_IteradorOnlineMembers + 1
        m_Iterador_ProximoUserIndex = p_OnlineMembers.Item(p_IteradorOnlineMembers)
    Else
        p_IteradorOnlineMembers = 0
        m_Iterador_ProximoUserIndex = 0
    End If
End Function

Public Function Iterador_ProximoGM() As Integer
    
    If p_IteradorOnlineGMs < p_GMsOnline.Count Then
        p_IteradorOnlineGMs = p_IteradorOnlineGMs + 1
        Iterador_ProximoGM = p_GMsOnline.Item(p_IteradorOnlineGMs)
    Else
        p_IteradorOnlineGMs = 0
        Iterador_ProximoGM = 0
    End If
End Function

Public Function Iterador_ProximaRelacion(ByVal R As RELACIONES_GUILD) As Integer
    
    While p_IteradorRelaciones < p_Relaciones.Count
        
        
        If p_Relaciones.Items(p_IteradorRelaciones) = R Then
            Iterador_ProximaRelacion = p_IteradorRelaciones
            Exit Function
        End If
        p_IteradorRelaciones = p_IteradorRelaciones + 1
    Wend

    If p_IteradorRelaciones >= p_Relaciones.Count Then
        p_IteradorRelaciones = 0
    End If
End Function
'
'
'


'
'ADMINISTRATIVAS
'

Public Sub ConectarGM(ByVal UserIndex As Integer)
    p_GMsOnline.Add UserIndex
End Sub

Public Sub DesconectarGM(ByVal UserIndex As Integer)
Dim i As Integer
    For i = 1 To p_GMsOnline.Count
        If p_GMsOnline.Item(i) = UserIndex Then
            p_GMsOnline.Remove (i)
        End If
    Next i
End Sub



'
'VARIAS, EXTRAS Y DEMASES
'

Private Sub ReplaceInvalidChars(ByRef S As String)
    If InStrB(S, Chr$(13)) <> 0 Then
        S = Replace(S, Chr$(13), vbNullString)
    End If
    If InStrB(S, Chr$(10)) <> 0 Then
        S = Replace(S, Chr$(10), vbNullString)
    End If
    If InStrB(S, "¬") <> 0 Then
        S = Replace(S, "¬", vbNullString)   'morgo usaba esto como "separador"
    End If
End Sub




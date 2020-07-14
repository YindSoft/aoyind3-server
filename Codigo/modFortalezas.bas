Attribute VB_Name = "modFortalezas"
Option Explicit

Private Fortaleza1 As clsFortaleza
Private Fortaleza2 As clsFortaleza

Public Sub initFortalezas()

Set Fortaleza1 = New clsFortaleza
Fortaleza1.Init (1)
Set Fortaleza2 = New clsFortaleza
Fortaleza2.Init (2)
End Sub

Public Sub MoverNPCFortaleza(ByVal NpcIndex As Integer)
With Npclist(NpcIndex)
    If .Pos.X <= 550 Then
        Fortaleza1.MoverNPC (NpcIndex)
    Else
        Fortaleza2.MoverNPC (NpcIndex)
    End If
End With
End Sub

Public Function FortalezaDelClan(ByVal GuildIndex As Integer, ByVal NpcIndex As Integer)
If Npclist(NpcIndex).Pos.X <= 550 Then
    FortalezaDelClan = Fortaleza1.IdClan = GuildIndex
Else
    FortalezaDelClan = Fortaleza2.IdClan = GuildIndex
End If
End Function

Public Sub ReSpawnFortaleza(ByVal Num As Byte, ByVal Izquierda As Boolean)
If Izquierda Then
    Fortaleza1.ReSpawn (Num)
Else
    Fortaleza2.ReSpawn (Num)
End If
End Sub

Public Sub AvisarAtaqueFortaleza(ByVal NpcIndex As Integer, ByVal UserIndex As Integer)
If Npclist(NpcIndex).NPCtype = eNPCType.Fortaleza Then
    If Npclist(NpcIndex).Pos.X <= 550 Then
        Call Fortaleza1.AvisarAtaque(NpcIndex, UserIndex)
    Else
        Call Fortaleza2.AvisarAtaque(NpcIndex, UserIndex)
    End If
End If
End Sub

Public Sub CheckRespawns()
If Not Fortaleza1 Is Nothing Then
    Call Fortaleza1.CheckRespawns
    Call Fortaleza2.CheckRespawns
End If
End Sub

Public Sub HandleProteger(ByVal UserIndex As Integer, ByVal Opcion As Byte)
If Opcion = 0 And UserList(UserIndex).Zona = 90 Then
    Call Fortaleza1.Proteger(UserIndex)
ElseIf Opcion = 0 And UserList(UserIndex).Zona = 89 Then
    Call Fortaleza2.Proteger(UserIndex)
ElseIf Opcion = 1 Then
    Call Fortaleza1.SumUser(UserIndex)
ElseIf Opcion = 2 Then
    Call Fortaleza2.SumUser(UserIndex)
End If
End Sub

Attribute VB_Name = "modMercader"
Option Explicit
Private Camino() As Position
Private Const CANT_PASOS As Byte = 23
Private PasoActual As Byte
Private NpcMercader As Integer
Private Const MERCADER_NPC As Integer = 617
Private RutaMercader As clsTree
Private MercaderLlega As Boolean
Private startX As Integer, startY As Integer

Public MercaderReal As clsMercader
Public MercaderCaos As clsMercader

Sub initMercader()


Set MercaderReal = New clsMercader

Call MercaderReal.Init(617, "297,859;285,863;285,845;277,838;277,802;285,794;293,793;292,687;296,683;296,617;300,615;301,557;305,556;305,520;317,519;315,492;305,491;305,346;296,344;296,282;292,277;292,222", "Ullathorpe", "Banderbill")

Set MercaderCaos = New clsMercader

Call MercaderCaos.Init(618, "195,1225;195,1235;201,1235;201,1251;217,1252;220,1261;272,1261;272,1268;304,1269;305,1296;332,1297;333,1308;424,1309;425,1332;450,1343;456,1343;456,1396;459,1404;459,1430;536,1431;536,1423;540,1420;540,1403;550,1397;552,1392", "Nix", "Arkhein")

End Sub

Public Sub ReSpawnMercader(ByVal NPC As Integer)
If NPC = MercaderReal.NpcNum Then
    MercaderReal.ReSpawn
ElseIf NPC = MercaderCaos.NpcNum Then
    MercaderCaos.ReSpawn
End If
End Sub
Public Sub MoverMercader(ByVal NpcIndex As Integer)
Call MercaderByIndex(NpcIndex).MoverMercader
End Sub
Public Sub MercaderAtacado(NpcIndex As Integer, ByVal UserIndex As Integer)
Call MercaderByIndex(NpcIndex).AgregarAgresor(UserIndex)
End Sub
Public Sub MercaderClicked(byvalNpcIndex As Integer, ByVal UserIndex As Integer)
Call MercaderByIndex(byvalNpcIndex).Clicked(UserIndex)
End Sub
Public Sub QuitarAgresorMercader(ByVal UserIndex As Integer)
MercaderReal.QuitarAgresor (UserIndex)
MercaderCaos.QuitarAgresor (UserIndex)
End Sub
Public Function MercaderByIndex(ByVal NpcIndex As Integer) As clsMercader
If NpcIndex = MercaderReal.NpcIndex Then
    Set MercaderByIndex = MercaderReal
ElseIf NpcIndex = MercaderCaos.NpcIndex Then
    Set MercaderByIndex = MercaderCaos
Else
    Set MercaderByIndex = Nothing
End If
End Function

Public Function EsMercader(ByVal NpcIndex As Integer, ByVal Bueno As Boolean)
If MercaderByIndex(NpcIndex) Is Nothing Then
    EsMercader = False
Else
    EsMercader = Npclist(NpcIndex).Stats.Alineacion = IIf(Bueno, 0, 1)
End If
End Function

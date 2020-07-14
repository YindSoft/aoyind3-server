Attribute VB_Name = "Trabajo"
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

Private Const GASTO_ENERGIA_TRABAJADOR As Byte = 2
Private Const GASTO_ENERGIA_NO_TRABAJADOR As Byte = 6


Public Sub DoPermanecerOculto(ByVal UserIndex As Integer, ByVal Mueve As Boolean)
'********************************************************
'Autor: Nacho (Integer)
'Last Modif: 11/19/2009
'Chequea si ya debe mostrarse
'Pablo (ToxicWaste): Cambie los ordenes de prioridades porque sino no andaba.
'13/01/2010: ZaMa - Now hidden on boat pirats recover the proper boat body.
'13/01/2010: ZaMa - Arreglo condicional para que el bandido camine oculto.
'********************************************************
On Error GoTo Errhandler
    With UserList(UserIndex)
        .Counters.TiempoOculto = .Counters.TiempoOculto - 1
        If .Counters.TiempoOculto <= 0 Then
            If .clase = eClass.Hunter And .Stats.UserSkills(eSkill.Ocultarse) > 90 Then
                If .Invent.ArmourEqpObjIndex = 648 Or .Invent.ArmourEqpObjIndex = 360 Then
                    .Counters.TiempoOculto = IntervaloOculto
                    Exit Sub
                End If
            End If
            .Counters.TiempoOculto = 0
            .flags.Oculto = 0
            
            If .flags.Navegando = 1 Then
                If .clase = eClass.Pirat Then
                    ' Pierde la apariencia de fragata fantasmal
                    Call ToggleBoatBody(UserIndex)
                    Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                        NingunEscudo, NingunCasco)
                End If
            Else
                If .flags.invisible = 0 Then
                    Call WriteConsoleMsg(UserIndex, "Has vuelto a ser visible.", FontTypeNames.FONTTYPE_INFO)
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
            End If
        End If
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoPermanecerOculto")



End Sub

Public Sub DoOcultarse(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'Pablo (ToxicWaste): No olvidar agregar IntervaloOculto=500 al Server.ini.
'Modifique la fórmula y ahora anda bien.
'13/01/2010: ZaMa - El pirata se transforma en galeon fantasmal cuando se oculta en agua.
'***************************************************

On Error GoTo Errhandler

    Dim Suerte As Double
    Dim res As Integer
    Dim Skill As Integer
    
    With UserList(UserIndex)
        Skill = .Stats.UserSkills(eSkill.Ocultarse)
        
        Suerte = (((0.000002 * Skill - 0.0002) * Skill + 0.0064) * Skill + 0.1124) * 100
        
        res = RandomNumber(1, 100)
        
        If res <= Suerte Then
        
            .flags.Oculto = 1
            Suerte = (-0.000001 * (100 - Skill) ^ 3)
            Suerte = Suerte + (0.00009229 * (100 - Skill) ^ 2)
            Suerte = Suerte + (-0.0088 * (100 - Skill))
            Suerte = Suerte + (0.9571)
            Suerte = Suerte * IntervaloOculto
            
            If .clase = eClass.Bandit Then
                .Counters.TiempoOculto = Int(Suerte / 2)
            Else
                .Counters.TiempoOculto = Suerte
            End If
            
            ' No es pirata o es uno sin barca
            If .flags.Navegando = 0 Then
                'Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, True)
                Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, True)
                Call WriteConsoleMsg(UserIndex, "¡Te has escondido entre las sombras!", FontTypeNames.FONTTYPE_INFO)
            ' Es un pirata navegando
            Else
                ' Le cambiamos el body a galeon fantasmal
                .Char.Body = iFragataFantasmal
                ' Actualizamos clientes
                Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, NingunArma, _
                                    NingunEscudo, NingunCasco)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse)
        Else
            '[CDT 17-02-2004]
            If Not .flags.UltimoMensaje = 4 Then
                Call WriteConsoleMsg(UserIndex, "¡No has logrado esconderte!", FontTypeNames.FONTTYPE_INFO)
                .flags.UltimoMensaje = 4
            End If
            '[/CDT]
            
            Call SubirSkill(UserIndex, eSkill.Ocultarse)
        End If
        
        .Counters.Ocultando = .Counters.Ocultando + 1
    End With
    
    Exit Sub

Errhandler:
    Call LogError("Error en Sub DoOcultarse")
End Sub

Public Sub ToggleBoatBody(ByVal UserIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modification: 25/07/2010
'Gives boat body depending on user alignment.
'25/07/2010: ZaMa - Now makes difference depending on faccion and atacable status.
'***************************************************

    Dim Ropaje As Integer
    Dim EsFaccionario As Boolean
    Dim NewBody As Integer
    
    With UserList(UserIndex)
 
        .Char.Head = 0
        If .Invent.BarcoObjIndex = 0 Then Exit Sub
        
        Ropaje = ObjData(.Invent.BarcoObjIndex).Ropaje
        
        ' Criminales y caos
        If Criminal(UserIndex) Then
            
            EsFaccionario = esCaos(UserIndex)
            
            Select Case Ropaje
                Case iBarca
                    If EsFaccionario Then
                        NewBody = iBarcaCaos
                    Else
                        NewBody = iBarcaPk
                    End If
                
                Case iGalera
                    If EsFaccionario Then
                        NewBody = iGaleraCaos
                    Else
                        NewBody = iGaleraPk
                    End If
                    
                Case iGaleon
                    If EsFaccionario Then
                        NewBody = iGaleonCaos
                    Else
                        NewBody = iGaleonPk
                    End If
            End Select
        
        ' Ciudas y Armadas
        Else
            
            EsFaccionario = esArmada(UserIndex)
            
            ' Atacable
            If False Then '.flags.AtacablePor <> 0 Then
                
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaRealAtacable
                        Else
                            NewBody = iBarcaCiudaAtacable
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraRealAtacable
                        Else
                            NewBody = iGaleraCiudaAtacable
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonRealAtacable
                        Else
                            NewBody = iGaleonCiudaAtacable
                        End If
                End Select
            
            ' Normal
            Else
            
                Select Case Ropaje
                    Case iBarca
                        If EsFaccionario Then
                            NewBody = iBarcaReal
                        Else
                            NewBody = iBarcaCiuda
                        End If
                    
                    Case iGalera
                        If EsFaccionario Then
                            NewBody = iGaleraReal
                        Else
                            NewBody = iGaleraCiuda
                        End If
                        
                    Case iGaleon
                        If EsFaccionario Then
                            NewBody = iGaleonReal
                        Else
                            NewBody = iGaleonCiuda
                        End If
                End Select
            
            End If
            
        End If
        
        .Char.Body = NewBody
        .Char.ShieldAnim = NingunEscudo
        .Char.WeaponAnim = NingunArma
        .Char.CascoAnim = NingunCasco
    End With

End Sub



Public Sub DoNavega(ByVal UserIndex As Integer, ByRef Barco As ObjData, ByVal Slot As Integer)

'***************************************************
'Author: Unknown
'Last Modification: 13/01/2010 (ZaMa)
'13/01/2010: ZaMa - El pirata pierde el ocultar si desequipa barca.
'16/09/2010: ZaMa - Ahora siempre se va el invi para los clientes al equipar la barca (Evita cortes de cabeza).
'10/12/2010: Pato - Limpio las variables del inventario que hacen referencia a la barca, sino el pirata que la última barca que equipo era el galeón no explotaba(Y capaz no la tenía equipada :P).
'***************************************************

    Dim ModNave As Single
    
    With UserList(UserIndex)
        ModNave = ModNavegacion(.clase, UserIndex)
        
        If .Stats.UserSkills(eSkill.Navegacion) / ModNave < Barco.MinSkill Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes conocimientos para usar este barco.", FontTypeNames.FONTTYPE_INFO)
            Call WriteConsoleMsg(UserIndex, "Para usar este barco necesitas " & Barco.MinSkill * ModNave & " puntos en navegacion.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        ' No estaba navegando
        If .flags.Navegando = 0 Then
            .Invent.BarcoObjIndex = .Invent.Object(Slot).ObjIndex
            .Invent.BarcoSlot = Slot
            
            .Char.Head = 0
            
            ' No esta muerto
            If .flags.Muerto = 0 Then
            
                Call ToggleBoatBody(UserIndex)
                
                ' Pierde el ocultar
                If .flags.Oculto = 1 Then
                    .flags.Oculto = 0
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                    Call WriteConsoleMsg(UserIndex, "¡Has vuelto a ser visible!", FontTypeNames.FONTTYPE_INFO)
                End If
               
                ' Siempre se ve la barca (Nunca esta invisible), pero solo para el cliente.
                If .flags.invisible = 1 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, False)
                End If
                
            ' Esta muerto
            Else
                .Char.Body = iFragataFantasmal
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            
            ' Comienza a navegar
            .flags.Navegando = 1
        
        ' Estaba navegando
        Else
            .Invent.BarcoObjIndex = 0
            .Invent.BarcoSlot = 0
        
            ' No esta muerto
            If .flags.Muerto = 0 Then
                .Char.Head = .OrigChar.Head
                
                If .clase = eClass.Pirat Then
                    If .flags.Oculto = 1 Then
                        ' Al desequipar barca, perdió el ocultar
                        .flags.Oculto = 0
                        .Counters.Ocultando = 0
                        Call WriteConsoleMsg(UserIndex, "¡Has recuperado tu apariencia normal!", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                If .Invent.ArmourEqpObjIndex > 0 Then
                    .Char.Body = ObjData(.Invent.ArmourEqpObjIndex).Ropaje
                Else
                    Call DarCuerpoDesnudo(UserIndex)
                End If
                
                If .Invent.EscudoEqpObjIndex > 0 Then _
                    .Char.ShieldAnim = ObjData(.Invent.EscudoEqpObjIndex).ShieldAnim
                If .Invent.WeaponEqpObjIndex > 0 Then _
                    .Char.WeaponAnim = ObjData(.Invent.WeaponEqpObjIndex).WeaponAnim
                If .Invent.CascoEqpObjIndex > 0 Then _
                    .Char.CascoAnim = ObjData(.Invent.CascoEqpObjIndex).CascoAnim
                
                
                ' Al dejar de navegar, si estaba invisible actualizo los clientes
                If .flags.invisible = 1 Then
                    Call UsUaRiOs.SetInvisible(UserIndex, .Char.CharIndex, True)
                End If
                
            ' Esta muerto
            Else
                .Char.Body = iCuerpoMuerto
                .Char.Head = iCabezaMuerto
                .Char.ShieldAnim = NingunEscudo
                .Char.WeaponAnim = NingunArma
                .Char.CascoAnim = NingunCasco
            End If
            
            ' Termina de navegar
            .flags.Navegando = 0
        End If
        
        ' Actualizo clientes
        Call ChangeUserChar(UserIndex, .Char.Body, .Char.Head, .Char.Heading, .Char.WeaponAnim, .Char.ShieldAnim, .Char.CascoAnim)
    End With
    
    Call WriteNavigateToggle(UserIndex)


End Sub

Public Sub FundirMineral(ByVal UserIndex As Integer)

On Error GoTo Errhandler

If UserList(UserIndex).flags.TargetObjInvIndex > 0 Then
   
   If ObjData(UserList(UserIndex).flags.TargetObjInvIndex).OBJType = eOBJType.otMinerales And ObjData(UserList(UserIndex).flags.TargetObjInvIndex).MinSkill <= UserList(UserIndex).Stats.UserSkills(eSkill.Mineria) / ModFundicion(UserList(UserIndex).clase) Then
        Call DoLingotes(UserIndex)
   Else
        Call WriteConsoleMsg(UserIndex, "No tenes conocimientos de mineria suficientes para trabajar este mineral.", FontTypeNames.FONTTYPE_INFO)
   End If

End If

Exit Sub

Errhandler:
    Call LogError("Error en FundirMineral. Error " & Err.Number & " : " & Err.Description)

End Sub
Public Sub DoDesequipar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)
'***************************************************
'Author: ZaMa
'Last Modif: 15/04/2010
'Unequips either shield, weapon or helmet from target user.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    Dim AlgoEquipado As Boolean
    
    With UserList(UserIndex)
        ' Si no tiene guantes de hurto no desequipa.
        If .Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
        ' Si no esta solo con manos, no desequipa tampoco.
        If .Invent.WeaponEqpObjIndex > 0 Then Exit Sub
        
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
   End With
   
   With UserList(VictimIndex)
        ' Si tiene escudo, intenta desequiparlo
        If .Invent.EscudoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.EscudoEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el escudo de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el escudo!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene escudo, o fallo desequiparlo, entonces trata de desequipar arma
        If .Invent.WeaponEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.WeaponEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
        
        ' No tiene arma, o fallo desequiparla, entonces trata de desequipar casco
        If .Invent.CascoEqpObjIndex > 0 Then
            
            Resultado = RandomNumber(1, 100)
            
            If Resultado <= Probabilidad Then
                ' Se lo desequipo
                Call Desequipar(VictimIndex, .Invent.CascoEqpSlot, True)
                
                Call WriteConsoleMsg(UserIndex, "Has logrado desequipar el casco de tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
                
                If .Stats.ELV < 20 Then
                    Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desequipado el casco!", FontTypeNames.FONTTYPE_FIGHT)
                End If
                
                Call FlushBuffer(VictimIndex)
                
                Exit Sub
            End If
            
            AlgoEquipado = True
        End If
    
        If AlgoEquipado Then
            Call WriteConsoleMsg(UserIndex, "Tu oponente no tiene equipado items!", FontTypeNames.FONTTYPE_FIGHT)
        Else
            Call WriteConsoleMsg(UserIndex, "No has logrado desequipar ningún item a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
        End If
    
    End With


End Sub
Function TieneObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer) As Boolean
'Call LogTarea("Sub TieneObjetos")

Dim i As Integer
Dim Total As Long
For i = 1 To MAX_INVENTORY_SLOTS
    If UserList(UserIndex).Invent.Object(i).ObjIndex = ItemIndex Then
        Total = Total + UserList(UserIndex).Invent.Object(i).Amount
    End If
Next i

If Cant <= Total Then
    TieneObjetos = True
    Exit Function
End If
        
End Function

Public Sub QuitarObjetos(ByVal ItemIndex As Integer, ByVal Cant As Integer, ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 05/08/09
'05/08/09: Pato - Cambie la funcion a procedimiento ya que se usa como procedimiento siempre, y fixie el bug 2788199
'***************************************************

'Call LogTarea("Sub QuitarObjetos")

Dim i As Integer
For i = 1 To MAX_INVENTORY_SLOTS
    With UserList(UserIndex).Invent.Object(i)
        If .ObjIndex = ItemIndex Then
            If .Amount <= Cant And .Equipped = 1 Then Call Desequipar(UserIndex, i, True)
            
            .Amount = .Amount - Cant
            If .Amount <= 0 Then
                Cant = Abs(.Amount)
                UserList(UserIndex).Invent.NroItems = UserList(UserIndex).Invent.NroItems - 1
                .Amount = 0
                .ObjIndex = 0
            Else
                Cant = 0
            End If
            
            Call UpdateUserInv(False, UserIndex, i)
            
            If Cant = 0 Then Exit Sub
        End If
    End With
Next i

End Sub

Sub HerreroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
'***************************************************
    With ObjData(ItemIndex)
        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, .LingP * CantidadItems, UserIndex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, .LingO * CantidadItems, UserIndex)
    End With
End Sub

Sub CarpinteroQuitarMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora quita tambien madera elfica
'***************************************************
    With ObjData(ItemIndex)
        If .Madera > 0 Then Call QuitarObjetos(Leña, .Madera * CantidadItems, UserIndex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, .MaderaElfica * CantidadItems, UserIndex)
    End With
End Sub

Function CarpinteroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal Cantidad As Integer, Optional ByVal ShowMsg As Boolean = False) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Agregada validacion a madera elfica.
'16/11/2009: ZaMa - Ahora considera la cantidad de items a construir
'***************************************************
    
    With ObjData(ItemIndex)
        If .Madera > 0 Then
            If Not TieneObjetos(Leña, .Madera * Cantidad, UserIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LeñaElfica, .MaderaElfica * Cantidad, UserIndex) Then
                If ShowMsg Then Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
                CarpinteroTieneMateriales = False
                Exit Function
            End If
        End If
    
    End With
    CarpinteroTieneMateriales = True

End Function
 
Function HerreroTieneMateriales(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Agregada validacion a madera elfica.
'***************************************************
    With ObjData(ItemIndex)
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, .LingH * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, .LingP * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, .LingO * CantidadItems, UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                HerreroTieneMateriales = False
                Exit Function
            End If
        End If
    End With
    HerreroTieneMateriales = True
End Function
Public Function MaxItemsConstruibles(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: ZaMa
'Last Modification: 29/01/2010
'
'***************************************************
    MaxItemsConstruibles = MaximoInt(1, CInt((UserList(UserIndex).Stats.ELV - 4) / 5))
End Function
Public Function PuedeConstruir(ByVal UserIndex As Integer, ByVal ItemIndex As Integer, ByVal CantidadItems As Integer) As Boolean
'***************************************************
'Author: Unknown
'Last Modification: 24/08/2009
'24/08/2008: ZaMa - Validates if the player has the required skill
'16/11/2009: ZaMa - Validates if the player has the required amount of materials, depending on the number of items to make
'***************************************************
PuedeConstruir = HerreroTieneMateriales(UserIndex, ItemIndex, CantidadItems) And _
                    Round(UserList(UserIndex).Stats.UserSkills(eSkill.Herreria) / ModHerreriA(UserList(UserIndex).clase), 0) >= ObjData(ItemIndex).SkHerreria
End Function

Public Function PuedeConstruirHerreria(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ArmasHerrero)
    If ArmasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
For i = 1 To UBound(ArmadurasHerrero)
    If ArmadurasHerrero(i) = ItemIndex Then
        PuedeConstruirHerreria = True
        Exit Function
    End If
Next i
PuedeConstruirHerreria = False
End Function


Public Sub HerreroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 30/05/2010
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items.
'22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
'30/05/2010: ZaMa - Los pks no suben plebe al trabajar.
'***************************************************
Dim CantidadItems As Long
Dim TieneMateriales As Boolean
Dim OtroUserIndex As Integer

With UserList(UserIndex)
    If .flags.Comerciando Then
        OtroUserIndex = .ComUsu.DestUsu
            
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
        
    CantidadItems = .Construir.PorCiclo
    
    If .Construir.Cantidad < CantidadItems Then _
        CantidadItems = .Construir.Cantidad
        
    If .Construir.Cantidad > 0 Then _
        .Construir.Cantidad = .Construir.Cantidad - CantidadItems
        
    If CantidadItems = 0 Then
        Call WriteStopWorking(UserIndex)
        Exit Sub
    End If
    
    If PuedeConstruirHerreria(ItemIndex) Then
        
        While CantidadItems > 0 And Not TieneMateriales
            If PuedeConstruir(UserIndex, ItemIndex, CantidadItems) Then
                TieneMateriales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        ' Chequeo si puede hacer al menos 1 item
        If Not TieneMateriales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes materiales.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        'Sacamos energía
        If .clase = eClass.Worker Then
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        Else
            'Chequeamos que tenga los puntos antes de sacarselos
            If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                Call WriteUpdateSta(UserIndex)
            Else
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                Exit Sub
            End If
        End If
        
        Call HerreroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
        ' AGREGAR FX
        
        Select Case ObjData(ItemIndex).OBJType
        
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armas!", "el arma!"), FontTypeNames.FONTTYPE_INFO)
            Case eOBJType.otESCUDO
                Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " escudos!", "el escudo!"), FontTypeNames.FONTTYPE_INFO)
            Case Is = eOBJType.otCASCO
                Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " cascos!", "el casco!"), FontTypeNames.FONTTYPE_INFO)
            Case eOBJType.otArmadura
                Call WriteConsoleMsg(UserIndex, "Has construido " & IIf(CantidadItems > 1, CantidadItems & " armaduras", "la armadura!"), FontTypeNames.FONTTYPE_INFO)
        
        End Select
        
        Dim MiObj As Obj
        
        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ItemIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
        If ObjData(MiObj.ObjIndex).Log = 1 Then
            Call LogDesarrollo(.Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
        End If
        
        Call SubirSkill(UserIndex, eSkill.Herreria)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.X, .Pos.Y))
        
        If Not Criminal(UserIndex) Then
            .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta
            If .Reputacion.PlebeRep > MAXREP Then _
                .Reputacion.PlebeRep = MAXREP
        End If
        
        .Counters.Trabajando = .Counters.Trabajando + 1
    End If
End With
End Sub

Public Function PuedeConstruirCarpintero(ByVal ItemIndex As Integer) As Boolean
Dim i As Long

For i = 1 To UBound(ObjCarpintero)
    If ObjCarpintero(i) = ItemIndex Then
        PuedeConstruirCarpintero = True
        Exit Function
    End If
Next i
PuedeConstruirCarpintero = False

End Function

Public Sub CarpinteroConstruirItem(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 28/05/2010
'24/08/2008: ZaMa - Validates if the player has the required skill
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'22/05/2010: ZaMa - Los caos ya no suben plebe al trabajar.
'28/05/2010: ZaMa - Los pks no suben plebe al trabajar.
'***************************************************
On Error GoTo Errhandler

    Dim CantidadItems As Long
    Dim TieneMateriales As Boolean
    Dim WeaponIndex As Integer
    Dim OtroUserIndex As Integer
    
    With UserList(UserIndex)
        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(UserIndex, "Debes tener equipado el serrucho para trabajar.", FontTypeNames.FONTTYPE_INFO)
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
        
        CantidadItems = .Construir.PorCiclo
        
        If .Construir.Cantidad < CantidadItems Then _
            CantidadItems = .Construir.Cantidad
            
        If .Construir.Cantidad > 0 Then _
            .Construir.Cantidad = .Construir.Cantidad - CantidadItems
            
        If CantidadItems = 0 Then
            Call WriteStopWorking(UserIndex)
            Exit Sub
        End If
    
        If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.clase), 0) >= _
           ObjData(ItemIndex).SkCarpinteria And _
           PuedeConstruirCarpintero(ItemIndex) Then
           
            ' Calculo cuantos item puede construir
            While CantidadItems > 0 And Not TieneMateriales
                If CarpinteroTieneMateriales(UserIndex, ItemIndex, CantidadItems) Then
                    TieneMateriales = True
                Else
                    CantidadItems = CantidadItems - 1
                End If
            Wend
            
            ' No tiene los materiales ni para construir 1 item?
            If Not TieneMateriales Then
                ' Para que muestre el mensaje
                Call CarpinteroTieneMateriales(UserIndex, ItemIndex, 1, True)
                Call WriteStopWorking(UserIndex)
                Exit Sub
            End If
           
            'Sacamos energía
            If .clase = eClass.Worker Then
                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
                    Call WriteUpdateSta(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            Else
                'Chequeamos que tenga los puntos antes de sacarselos
                If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
                    .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
                    Call WriteUpdateSta(UserIndex)
                Else
                    Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
                    Exit Sub
                End If
            End If
            
            Call CarpinteroQuitarMateriales(UserIndex, ItemIndex, CantidadItems)
            Call WriteConsoleMsg(UserIndex, "Has construido " & CantidadItems & _
                                IIf(CantidadItems = 1, " objeto!", " objetos!"), FontTypeNames.FONTTYPE_INFO)
            
            Dim MiObj As Obj
            MiObj.Amount = CantidadItems
            MiObj.ObjIndex = ItemIndex
            If Not MeterItemEnInventario(UserIndex, MiObj) Then
                Call TirarItemAlPiso(.Pos, MiObj)
            End If
            
            'Log de construcción de Items. Pablo (ToxicWaste) 10/09/07
            If ObjData(MiObj.ObjIndex).Log = 1 Then
                Call LogDesarrollo(.Name & " ha construído " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name)
            End If
            
            Call SubirSkill(UserIndex, eSkill.Carpinteria)
            Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.X, .Pos.Y))
            
            If Not Criminal(UserIndex) Then
                .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta
                If .Reputacion.PlebeRep > MAXREP Then _
                    .Reputacion.PlebeRep = MAXREP
            End If
            
            .Counters.Trabajando = .Counters.Trabajando + 1
        End If
    End With
    
    Exit Sub
Errhandler:
    Call LogError("Error en CarpinteroConstruirItem. Error " & Err.Number & " : " & Err.Description & ". UserIndex:" & UserIndex & ". ItemIndex:" & ItemIndex)
End Sub

Private Function MineralesParaLingote(ByVal Lingote As iMinerales) As Integer
    Select Case Lingote
        Case iMinerales.HierroCrudo
            MineralesParaLingote = 14
        Case iMinerales.PlataCruda
            MineralesParaLingote = 20
        Case iMinerales.OroCrudo
            MineralesParaLingote = 35
        Case Else
            MineralesParaLingote = 10000
    End Select
End Function


Public Sub DoLingotes(ByVal UserIndex As Integer)
'***************************************************
'Author: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Implementado nuevo sistema de construccion de items
'***************************************************
'    Call LogTarea("Sub DoLingotes")
    Dim Slot As Integer
    Dim obji As Integer
    Dim CantidadItems As Integer
    Dim TieneMinerales As Boolean
    Dim OtroUserIndex As Integer
    
    With UserList(UserIndex)
        If .flags.Comerciando Then
            OtroUserIndex = .ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(UserIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
        
        CantidadItems = MaximoInt(1, CInt((.Stats.ELV - 4) / 5))

        Slot = .flags.TargetObjInvSlot
        obji = .Invent.Object(Slot).ObjIndex
        
        While CantidadItems > 0 And Not TieneMinerales
            If .Invent.Object(Slot).Amount >= MineralesParaLingote(obji) * CantidadItems Then
                TieneMinerales = True
            Else
                CantidadItems = CantidadItems - 1
            End If
        Wend
        
        If Not TieneMinerales Or ObjData(obji).OBJType <> eOBJType.otMinerales Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes minerales para hacer un lingote.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        .Invent.Object(Slot).Amount = .Invent.Object(Slot).Amount - MineralesParaLingote(obji) * CantidadItems
        If .Invent.Object(Slot).Amount < 1 Then
            .Invent.Object(Slot).Amount = 0
            .Invent.Object(Slot).ObjIndex = 0
        End If
        
        Dim MiObj As Obj
        MiObj.Amount = CantidadItems
        MiObj.ObjIndex = ObjData(.flags.TargetObjInvIndex).LingoteIndex
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(.Pos, MiObj)
        End If
        
        Call UpdateUserInv(False, UserIndex, Slot)
        Call WriteConsoleMsg(UserIndex, "¡Has obtenido " & CantidadItems & " lingote" & _
                            IIf(CantidadItems = 1, "", "s") & "!", FontTypeNames.FONTTYPE_INFO)
    
        .Counters.Trabajando = .Counters.Trabajando + 1
    End With
End Sub

Function ModNavegacion(ByVal clase As eClass, ByVal UserIndex As Integer) As Single
'***************************************************
'Autor: Unknown (orginal version)
'Last Modification: 27/11/2009
'27/11/2009: ZaMa - A worker can navigate before only if it's an expert fisher
'12/04/2010: ZaMa - Arreglo modificador de pescador, para que navegue con 60 skills.
'***************************************************
Select Case clase
    Case eClass.Pirat
        ModNavegacion = 1
    Case eClass.Worker
        If UserList(UserIndex).Stats.UserSkills(eSkill.Pesca) = 100 Then
            ModNavegacion = 1.71
        Else
            ModNavegacion = 2
        End If
    Case Else
        ModNavegacion = 2
End Select

End Function


Function ModFundicion(ByVal clase As eClass) As Single

Select Case clase
    Case eClass.Worker
        ModFundicion = 1
    Case Else
        ModFundicion = 3
End Select

End Function

Function ModCarpinteria(ByVal clase As eClass) As Integer

Select Case clase
    Case eClass.Worker
        ModCarpinteria = 1
    Case Else
        ModCarpinteria = 3
End Select

End Function

Function ModHerreriA(ByVal clase As eClass) As Single
Select Case clase
    Case eClass.Worker
        ModHerreriA = 1
    Case Else
        ModHerreriA = 4
End Select

End Function

Function ModDomar(ByVal clase As eClass) As Integer
    Select Case clase
        Case eClass.Druid
            ModDomar = 6
        Case eClass.Hunter
            ModDomar = 6
        Case eClass.Cleric
            ModDomar = 7
        Case Else
            ModDomar = 10
    End Select
End Function

Function FreeMascotaIndex(ByVal UserIndex As Integer) As Integer
'***************************************************
'Author: Unknown
'Last Modification: 02/03/09
'02/03/09: ZaMa - Busca un indice libre de mascotas, revisando los types y no los indices de los npcs
'***************************************************
    Dim j As Integer
    For j = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(j) = 0 Then
            FreeMascotaIndex = j
            Exit Function
        End If
    Next j
End Function

Sub DoDomar(ByVal UserIndex As Integer, ByVal NpcIndex As Integer)
'***************************************************
'Author: Nacho (Integer)
'Last Modification: 02/03/2009
'12/15/2008: ZaMa - Limits the number of the same type of pet to 2.
'02/03/2009: ZaMa - Las criaturas domadas en zona segura, esperan afuera (desaparecen).
'***************************************************

On Error GoTo Errhandler

Dim puntosDomar As Integer
Dim puntosRequeridos As Integer
Dim CanStay As Boolean
Dim petType As Integer
Dim NroPets As Integer


If Npclist(NpcIndex).MaestroUser = UserIndex Then
    Call WriteConsoleMsg(UserIndex, "Ya domaste a esa criatura.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).NroMascotas < MAXMASCOTAS Then
    
    If Npclist(NpcIndex).MaestroNpc > 0 Or Npclist(NpcIndex).MaestroUser > 0 Then
        Call WriteConsoleMsg(UserIndex, "La criatura ya tiene amo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    If Not PuedeDomarMascota(UserIndex, NpcIndex) Then
        Call WriteConsoleMsg(UserIndex, "No puedes domar mas de dos criaturas del mismo tipo.", FontTypeNames.FONTTYPE_INFO)
        Exit Sub
    End If
    
    puntosDomar = CInt(UserList(UserIndex).Stats.UserAtributos(eAtributos.Carisma)) * CInt(UserList(UserIndex).Stats.UserSkills(eSkill.Domar))
    If UserList(UserIndex).Invent.AnilloEqpObjIndex = FLAUTAMAGICA Then
        puntosRequeridos = Npclist(NpcIndex).flags.Domable * 0.8
    Else
        puntosRequeridos = Npclist(NpcIndex).flags.Domable
    End If
    
    If puntosRequeridos <= puntosDomar And RandomNumber(1, 5) = 1 Then
        Dim index As Integer
        UserList(UserIndex).NroMascotas = UserList(UserIndex).NroMascotas + 1
        index = FreeMascotaIndex(UserIndex)
        UserList(UserIndex).MascotasIndex(index) = NpcIndex
        UserList(UserIndex).MascotasType(index) = Npclist(NpcIndex).Numero
        
        Npclist(NpcIndex).MaestroUser = UserIndex
        
        Call FollowAmo(NpcIndex)
        Call ReSpawnNpc(Npclist(NpcIndex))
        
        Call WriteConsoleMsg(UserIndex, "La criatura te ha aceptado como su amo.", FontTypeNames.FONTTYPE_INFO)
        
        ' Es zona segura?
        CanStay = (Zonas(UserList(UserIndex).Zona).Segura = 0)
        
        If Not CanStay Then
            petType = Npclist(NpcIndex).Numero
            NroPets = UserList(UserIndex).NroMascotas
            
            Call QuitarNPC(NpcIndex)
            
            UserList(UserIndex).MascotasType(index) = petType
            UserList(UserIndex).NroMascotas = NroPets
            
            Call WriteConsoleMsg(UserIndex, "No se permiten mascotas en zona segura. Éstas te esperarán afuera.", FontTypeNames.FONTTYPE_INFO)
        End If

    Else
        If Not UserList(UserIndex).flags.UltimoMensaje = 5 Then
            Call WriteConsoleMsg(UserIndex, "No has logrado domar la criatura.", FontTypeNames.FONTTYPE_INFO)
            UserList(UserIndex).flags.UltimoMensaje = 5
        End If
    End If
    
    'Entreno domar. Es un 30% más dificil si no sos druida.
    If UserList(UserIndex).clase = eClass.Druid Or (RandomNumber(1, 3) < 3) Then
        Call SubirSkill(UserIndex, Domar)
    End If
Else
    Call WriteConsoleMsg(UserIndex, "No puedes controlar más criaturas.", FontTypeNames.FONTTYPE_INFO)
End If

Exit Sub

Errhandler:
    Call LogError("Error en DoDomar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Checks if the user can tames a pet.
'
' @param integer userIndex The user id from who wants tame the pet.
' @param integer NPCindex The index of the npc to tome.
' @return boolean True if can, false if not.
Private Function PuedeDomarMascota(ByVal UserIndex As Integer, ByVal NpcIndex As Integer) As Boolean
'***************************************************
'Author: ZaMa
'This function checks how many NPCs of the same type have
'been tamed by the user.
'Returns True if that amount is less than two.
'***************************************************
    Dim i As Long
    Dim numMascotas As Long
    
    For i = 1 To MAXMASCOTAS
        If UserList(UserIndex).MascotasType(i) = Npclist(NpcIndex).Numero Then
            numMascotas = numMascotas + 1
        End If
    Next i
    
    If numMascotas <= 1 Then PuedeDomarMascota = True
    
End Function

Sub DoAdminInvisible(ByVal UserIndex As Integer)
    
    If UserList(UserIndex).flags.AdminInvisible = 0 Then
        
        ' Sacamos el mimetizmo
        If UserList(UserIndex).flags.Mimetizado = 1 Then
            UserList(UserIndex).Char.Body = UserList(UserIndex).CharMimetizado.Body
            UserList(UserIndex).Char.Head = UserList(UserIndex).CharMimetizado.Head
            UserList(UserIndex).Char.CascoAnim = UserList(UserIndex).CharMimetizado.CascoAnim
            UserList(UserIndex).Char.ShieldAnim = UserList(UserIndex).CharMimetizado.ShieldAnim
            UserList(UserIndex).Char.WeaponAnim = UserList(UserIndex).CharMimetizado.WeaponAnim
            UserList(UserIndex).Counters.Mimetismo = 0
            UserList(UserIndex).flags.Mimetizado = 0
        End If
        
        UserList(UserIndex).flags.AdminInvisible = 1
        UserList(UserIndex).flags.invisible = 1
        UserList(UserIndex).flags.Oculto = 1
        UserList(UserIndex).flags.OldBody = UserList(UserIndex).Char.Body
        UserList(UserIndex).flags.OldHead = UserList(UserIndex).Char.Head
        UserList(UserIndex).Char.Body = 0
        UserList(UserIndex).Char.Head = 0
        
    Else
        
        UserList(UserIndex).flags.AdminInvisible = 0
        UserList(UserIndex).flags.invisible = 0
        UserList(UserIndex).flags.Oculto = 0
        UserList(UserIndex).Counters.TiempoOculto = 0
        UserList(UserIndex).Char.Body = UserList(UserIndex).flags.OldBody
        UserList(UserIndex).Char.Head = UserList(UserIndex).flags.OldHead
        
    End If
    
    'vuelve a ser visible por la fuerza
    Call ChangeUserChar(UserIndex, UserList(UserIndex).Char.Body, UserList(UserIndex).Char.Head, UserList(UserIndex).Char.Heading, UserList(UserIndex).Char.WeaponAnim, UserList(UserIndex).Char.ShieldAnim, UserList(UserIndex).Char.CascoAnim)
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageSetInvisible(UserList(UserIndex).Char.CharIndex, False))
End Sub

Sub TratarDeHacerFogata(ByVal map As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal UserIndex As Integer)

Dim Suerte As Byte
Dim exito As Byte
Dim Obj As Obj
Dim posMadera As WorldPos

If Not LegalPos(map, X, Y) Then Exit Sub

With posMadera
    .map = map
    .X = X
    .Y = Y
End With

If MapData(map, X, Y).ObjInfo.ObjIndex <> 58 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas clickear sobre Leña para hacer ramitas", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If Distancia(posMadera, UserList(UserIndex).Pos) > 2 Then
    Call WriteConsoleMsg(UserIndex, "Estás demasiado lejos para prender la fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If UserList(UserIndex).flags.Muerto = 1 Then
    Call WriteConsoleMsg(UserIndex, "No puedes hacer fogatas estando muerto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If

If MapData(map, X, Y).ObjInfo.Amount < 3 Then
    Call WriteConsoleMsg(UserIndex, "Necesitas por lo menos tres troncos para hacer una fogata.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 0 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) < 6 Then
    Suerte = 3
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 6 And UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) <= 34 Then
    Suerte = 2
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Supervivencia) >= 35 Then
    Suerte = 1
End If

exito = RandomNumber(1, Suerte)

If exito = 1 Then
    Obj.ObjIndex = FOGATA_APAG
    Obj.Amount = MapData(map, X, Y).ObjInfo.Amount \ 3
    
    Call WriteConsoleMsg(UserIndex, "Has hecho " & Obj.Amount & " fogatas.", FontTypeNames.FONTTYPE_INFO)
    
    Call MakeObj(Obj, map, X, Y)
    
    'Seteamos la fogata como el nuevo TargetObj del user
    UserList(UserIndex).flags.TargetObj = FOGATA_APAG
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 10 Then
        Call WriteConsoleMsg(UserIndex, "No has podido hacer la fogata.", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 10
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Supervivencia)


End Sub

Public Sub DoPescar(ByVal UserIndex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).clase = eClass.Worker Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(UserIndex).clase = eClass.Worker Then
        MiObj.Amount = RandomNumber(1, MaximoInt(1, UserList(UserIndex).Stats.ELV))
    Else
        MiObj.Amount = 1
    End If
    MiObj.ObjIndex = Pescado
    If RandomNumber(1, 11) = 1 Then
        Select Case RandomNumber(1, 3)
            Case 1
                MiObj.ObjIndex = PECES_POSIBLES.PESCADO2
            Case 2
                MiObj.ObjIndex = PECES_POSIBLES.PESCADO3
            Case 3
                MiObj.ObjIndex = PECES_POSIBLES.PESCADO4
        End Select
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has pescado un lindo pez!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 6 Then
      Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
      UserList(UserIndex).flags.UltimoMensaje = 6
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Pesca)

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoPescar. Error " & Err.Number & " : " & Err.Description)
End Sub

Public Sub DoPescarRed(ByVal UserIndex As Integer)
On Error GoTo Errhandler

Dim iSkill As Integer
Dim Suerte As Integer
Dim res As Integer
Dim EsPescador As Boolean

If UserList(UserIndex).flags.Navegando = 0 Then
    Call WriteConsoleMsg(UserIndex, "La red sólo puede ser utilizada desde el mar abierto.", FontTypeNames.FONTTYPE_INFO)
    Exit Sub
End If


If UserList(UserIndex).clase = eClass.Worker Then
    Call QuitarSta(UserIndex, EsfuerzoPescarPescador)
    EsPescador = True
Else
    Call QuitarSta(UserIndex, EsfuerzoPescarGeneral)
    EsPescador = False
End If

iSkill = UserList(UserIndex).Stats.UserSkills(eSkill.Pesca)

' m = (60-11)/(1-10)
' y = mx - m*10 + 11

Suerte = Int(-0.00125 * iSkill * iSkill - 0.3 * iSkill + 49)

If Suerte > 0 Then
    res = RandomNumber(1, Suerte)
    
    If res < 6 Then
        Dim MiObj As Obj
        Dim PecesPosibles(1 To 4) As Integer
        
        PecesPosibles(1) = PESCADO1
        PecesPosibles(2) = PESCADO2
        PecesPosibles(3) = PESCADO3
        PecesPosibles(4) = PESCADO4
        
        If EsPescador = True Then
            MiObj.Amount = RandomNumber(MaximoInt(1, UserList(UserIndex).Stats.ELV / 3), MaximoInt(1, UserList(UserIndex).Stats.ELV + 4))
        Else
            MiObj.Amount = RandomNumber(3, 8)
        End If
        MiObj.ObjIndex = PecesPosibles(RandomNumber(LBound(PecesPosibles), UBound(PecesPosibles)))
        
        If Not MeterItemEnInventario(UserIndex, MiObj) Then
            Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        End If
        
        Call WriteConsoleMsg(UserIndex, "¡Has pescado algunos peces!", FontTypeNames.FONTTYPE_INFO)
        
    Else
        Call WriteConsoleMsg(UserIndex, "¡No has pescado nada!", FontTypeNames.FONTTYPE_INFO)
    End If
    
    Call SubirSkill(UserIndex, Pesca)
End If

    UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
    If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
        UserList(UserIndex).Reputacion.PlebeRep = MAXREP
           
UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1
        
Exit Sub

Errhandler:
    Call LogError("Error en DoPescarRed")
End Sub

''
' Try to steal an item / gold to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen

Public Sub DoRobar(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'*************************************************
'Author: Unknown
'Last modified: 05/04/2010
'Last Modification By: ZaMa
'24/07/08: Marco - Now it calls to WriteUpdateGold(VictimaIndex and LadrOnIndex) when the thief stoles gold. (MarKoxX)
'27/11/2009: ZaMa - Optimizacion de codigo.
'18/12/2009: ZaMa - Los ladrones ciudas pueden robar a pks.
'01/04/2010: ZaMa - Los ladrones pasan a robar oro acorde a su nivel.
'05/04/2010: ZaMa - Los armadas no pueden robarle a ciudadanos jamas.
'23/04/2010: ZaMa - No se puede robar mas sin energia.
'23/04/2010: ZaMa - El alcance de robo pasa a ser de 1 tile.
'*************************************************

On Error GoTo Errhandler

    Dim OtroUserIndex As Integer

    'If Not MapInfo(UserList(VictimaIndex).Pos.map).Pk Then Exit Sub
    
    'If UserList(VictimaIndex).flags.EnConsulta Then
    '    Call WriteConsoleMsg(LadrOnIndex, "¡¡¡No puedes robar a usuarios en consulta!!!", FontTypeNames.FONTTYPE_INFO)
    '    Exit Sub
    'End If
    
    With UserList(LadrOnIndex)
    
        If .flags.Seguro Then
            If Not Criminal(VictimaIndex) Then
                Call WriteConsoleMsg(LadrOnIndex, "Debes quitarte el seguro para robarle a un ciudadano.", FontTypeNames.FONTTYPE_FIGHT)
                Exit Sub
            End If
        Else
            If .Faccion.ArmadaReal = 1 Then
                If Not Criminal(VictimaIndex) Then
                    Call WriteConsoleMsg(LadrOnIndex, "Los miembros del ejército real no tienen permitido robarle a ciudadanos.", FontTypeNames.FONTTYPE_FIGHT)
                    Exit Sub
                End If
            End If
        End If
        
        ' Caos robando a caos?
        If UserList(VictimaIndex).Faccion.FuerzasCaos = 1 And .Faccion.FuerzasCaos = 1 Then
            Call WriteConsoleMsg(LadrOnIndex, "No puedes robar a otros miembros de la legión oscura.", FontTypeNames.FONTTYPE_FIGHT)
            Exit Sub
        End If
        
        If TriggerZonaPelea(LadrOnIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub
    
        
        ' Tiene energia?
        If .Stats.MinSta < 15 Then
            If .genero = eGenero.Hombre Then
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansado para robar.", FontTypeNames.FONTTYPE_INFO)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "Estás muy cansada para robar.", FontTypeNames.FONTTYPE_INFO)
            End If
            
            Exit Sub
        End If
        
        ' Quito energia
        Call QuitarSta(LadrOnIndex, 15)
        
        Dim GuantesHurto As Boolean
    
        If .Invent.AnilloEqpObjIndex = GUANTE_HURTO Then GuantesHurto = True
        
        If UserList(VictimaIndex).flags.Privilegios And PlayerType.user Then
            
            Dim Suerte As Integer
            Dim res As Integer
            Dim RobarSkill As Byte
            
            RobarSkill = .Stats.UserSkills(eSkill.Robar)
                
            If RobarSkill <= 10 Then
                Suerte = 35
            ElseIf RobarSkill <= 20 Then
                Suerte = 30
            ElseIf RobarSkill <= 30 Then
                Suerte = 28
            ElseIf RobarSkill <= 40 Then
                Suerte = 24
            ElseIf RobarSkill <= 50 Then
                Suerte = 22
            ElseIf RobarSkill <= 60 Then
                Suerte = 20
            ElseIf RobarSkill <= 70 Then
                Suerte = 18
            ElseIf RobarSkill <= 80 Then
                Suerte = 15
            ElseIf RobarSkill <= 90 Then
                Suerte = 10
            ElseIf RobarSkill < 100 Then
                Suerte = 7
            Else
                Suerte = 5
            End If
            
            res = RandomNumber(1, Suerte)
                
            If res < 3 Then 'Exito robo
                If UserList(VictimaIndex).flags.Comerciando Then
                    OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                        
                    If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                        Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                        Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                        
                        Call LimpiarComercioSeguro(VictimaIndex)
                        Call Protocol.FlushBuffer(OtroUserIndex)
                    End If
                End If
               
                If (RandomNumber(1, 50) < 25) And (.clase = eClass.Thief) Then
                    If TieneObjetosRobables(VictimaIndex) Then
                        Call RobarObjeto(LadrOnIndex, VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
                    End If
                Else 'Roba oro
                    If UserList(VictimaIndex).Stats.GLD > 0 Then
                        Dim N As Long
                        
                        If .clase = eClass.Thief Then
                        ' Si no tine puestos los guantes de hurto roba un 50% menos. Pablo (ToxicWaste)
                            If GuantesHurto Then
                                N = RandomNumber(.Stats.ELV * 50, .Stats.ELV * 100)
                            Else
                                N = RandomNumber(.Stats.ELV * 25, .Stats.ELV * 50)
                            End If
                        Else
                            N = RandomNumber(1, 100)
                        End If
                        If N > UserList(VictimaIndex).Stats.GLD Then N = UserList(VictimaIndex).Stats.GLD
                        UserList(VictimaIndex).Stats.GLD = UserList(VictimaIndex).Stats.GLD - N
                        
                        .Stats.GLD = .Stats.GLD + N
                        If .Stats.GLD > MAXORO Then _
                            .Stats.GLD = MAXORO
                        
                        Call WriteConsoleMsg(LadrOnIndex, "Le has robado " & N & " monedas de oro a " & UserList(VictimaIndex).Name, FontTypeNames.FONTTYPE_INFO)
                        Call WriteUpdateGold(LadrOnIndex) 'Le actualizamos la billetera al ladron
                        
                        Call WriteUpdateGold(VictimaIndex) 'Le actualizamos la billetera a la victima
                        Call FlushBuffer(VictimaIndex)
                    Else
                        Call WriteConsoleMsg(LadrOnIndex, UserList(VictimaIndex).Name & " no tiene oro.", FontTypeNames.FONTTYPE_INFO)
                    End If
                End If
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar)
            Else
                Call WriteConsoleMsg(LadrOnIndex, "¡No has logrado robar nada!", FontTypeNames.FONTTYPE_INFO)
                Call WriteConsoleMsg(VictimaIndex, "¡" & .Name & " ha intentado robarte!", FontTypeNames.FONTTYPE_INFO)
                Call FlushBuffer(VictimaIndex)
                
                Call SubirSkill(LadrOnIndex, eSkill.Robar)
            End If
        
            If Not Criminal(LadrOnIndex) Then
                If Not Criminal(VictimaIndex) Then
                    Call VolverCriminal(LadrOnIndex)
                End If
            End If
            
            ' Se pudo haber convertido si robo a un ciuda
            If Criminal(LadrOnIndex) Then
                .Reputacion.LadronesRep = .Reputacion.LadronesRep + vlLadron
                If .Reputacion.LadronesRep > MAXREP Then _
                    .Reputacion.LadronesRep = MAXREP
            End If
        End If
    End With

Exit Sub

Errhandler:
    Call LogError("Error en DoRobar. Error " & Err.Number & " : " & Err.Description)

End Sub

''
' Check if one item is stealable
'
' @param VictimaIndex Specifies reference to victim
' @param Slot Specifies reference to victim's inventory slot
' @return If the item is stealable
Public Function ObjEsRobable(ByVal VictimaIndex As Integer, ByVal Slot As Integer) As Boolean
' Agregué los barcos
' Esta funcion determina qué objetos son robables.

Dim OI As Integer

OI = UserList(VictimaIndex).Invent.Object(Slot).ObjIndex

ObjEsRobable = _
ObjData(OI).OBJType <> eOBJType.otLlaves And _
UserList(VictimaIndex).Invent.Object(Slot).Equipped = 0 And _
ObjData(OI).Real = 0 And _
ObjData(OI).Caos = 0 And _
ObjData(OI).OBJType <> eOBJType.otBarcos

End Function

''
' Try to steal an item to another character
'
' @param LadrOnIndex Specifies reference to user that stoles
' @param VictimaIndex Specifies reference to user that is being stolen
Public Sub RobarObjeto(ByVal LadrOnIndex As Integer, ByVal VictimaIndex As Integer)
'Call LogTarea("Sub RobarObjeto")
Dim flag As Boolean
Dim i As Integer
flag = False

If RandomNumber(1, 12) < 6 Then 'Comenzamos por el principio o el final?
    i = 1
    Do While Not flag And i <= MAX_INVENTORY_SLOTS
        'Hay objeto en este slot?
        If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
           If ObjEsRobable(VictimaIndex, i) Then
                 If RandomNumber(1, 10) < 4 Then flag = True
           End If
        End If
        If Not flag Then i = i + 1
    Loop
Else
    i = 20
    Do While Not flag And i > 0
      'Hay objeto en este slot?
      If UserList(VictimaIndex).Invent.Object(i).ObjIndex > 0 Then
         If ObjEsRobable(VictimaIndex, i) Then
               If RandomNumber(1, 10) < 4 Then flag = True
         End If
      End If
      If Not flag Then i = i - 1
    Loop
End If

If flag Then
    Dim MiObj As Obj
    Dim Num As Byte
    'Cantidad al azar
    Num = RandomNumber(1, 5)
                
    If Num > UserList(VictimaIndex).Invent.Object(i).Amount Then
         Num = UserList(VictimaIndex).Invent.Object(i).Amount
    End If
                
    MiObj.Amount = Num
    MiObj.ObjIndex = UserList(VictimaIndex).Invent.Object(i).ObjIndex
    
    UserList(VictimaIndex).Invent.Object(i).Amount = UserList(VictimaIndex).Invent.Object(i).Amount - Num
                
    If UserList(VictimaIndex).Invent.Object(i).Amount <= 0 Then
          Call QuitarUserInvItem(VictimaIndex, CByte(i), 1)
    End If
            
    Call UpdateUserInv(False, VictimaIndex, CByte(i))
                
    If Not MeterItemEnInventario(LadrOnIndex, MiObj) Then
        Call TirarItemAlPiso(UserList(LadrOnIndex).Pos, MiObj)
    End If
    
    If UserList(LadrOnIndex).clase = eClass.Thief Then
        Call WriteConsoleMsg(LadrOnIndex, "Has robado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(LadrOnIndex, "Has hurtado " & MiObj.Amount & " " & ObjData(MiObj.ObjIndex).Name, FontTypeNames.FONTTYPE_INFO)
    End If
Else
    Call WriteConsoleMsg(LadrOnIndex, "No has logrado robar ningún objeto.", FontTypeNames.FONTTYPE_INFO)
End If

'If exiting, cancel de quien es robado
Call CancelExit(VictimaIndex)

End Sub

Public Sub DoApuñalar(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Nacho (Integer) & Unknown (orginal version)
'Last Modification: 04/17/08 - (NicoNZ)
'Simplifique la cuenta que hacia para sacar la suerte
'y arregle la cuenta que hacia para sacar el daño
'***************************************************
Dim Suerte As Integer
Dim Skill As Integer

Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Apuñalar)

Select Case UserList(UserIndex).clase
    Case eClass.Assasin
        Suerte = Int(((0.00003 * Skill - 0.002) * Skill + 0.098) * Skill + 4.25)
    
    Case eClass.Cleric, eClass.PALADIN, eClass.Pirat
        Suerte = Int(((0.000003 * Skill + 0.0006) * Skill + 0.0107) * Skill + 4.93)
    
    Case eClass.Bard
        Suerte = Int(((0.000002 * Skill + 0.0002) * Skill + 0.032) * Skill + 4.81)
    
    Case Else
        Suerte = Int(0.0361 * Skill + 4.39)
End Select


If RandomNumber(0, 100) < Suerte Then
    If VictimUserIndex <> 0 Then
        If UserList(UserIndex).clase = eClass.Assasin Then
            daño = Round(daño * 1.4, 0)
        Else
            daño = Round(daño * 1.5, 0)
        End If
        
        With UserList(VictimUserIndex)
            .Stats.MinHP = .Stats.MinHP - daño
            Call WriteConsoleMsg(UserIndex, "Has apuñalado a " & .Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
            Call WriteConsoleMsg(VictimUserIndex, "Te ha apuñalado " & UserList(UserIndex).Name & " por " & daño, FontTypeNames.FONTTYPE_FIGHT)
        End With
        
        Call FlushBuffer(VictimUserIndex)
    Else
        Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - Int(daño * 2)
        Call WriteConsoleMsg(UserIndex, "Has apuñalado la criatura por " & Int(daño * 2), FontTypeNames.FONTTYPE_FIGHT)
        '[Alejo]
        Call CalcularDarExp(UserIndex, VictimNpcIndex, daño * 2)
    End If
    
    Call SubirSkill(UserIndex, eSkill.Apuñalar)
Else
    Call WriteConsoleMsg(UserIndex, "¡No has logrado apuñalar a tu enemigo!", FontTypeNames.FONTTYPE_FIGHT)
    Call SubirSkill(UserIndex, eSkill.Apuñalar)
End If

End Sub

Public Sub DoGolpeCritico(ByVal UserIndex As Integer, ByVal VictimNpcIndex As Integer, ByVal VictimUserIndex As Integer, ByVal daño As Integer)
'***************************************************
'Autor: Pablo (ToxicWaste)
'Last Modification: 28/01/2007
'01/06/2010: ZaMa - Valido si tiene arma equipada antes de preguntar si es vikinga.
'***************************************************
    Dim Suerte As Integer
    Dim Skill As Integer
    Dim WeaponIndex As Integer
    
    With UserList(UserIndex)
        ' Es bandido?
        If .clase <> eClass.Bandit Then Exit Sub
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        
        ' Es una espada vikinga?
        If WeaponIndex <> ESPADA_VIKINGA Then Exit Sub
    
        Skill = .Stats.UserSkills(eSkill.Wrestling)
    End With
    
    Suerte = Int((((0.00000003 * Skill + 0.000006) * Skill + 0.000107) * Skill + 0.0893) * 100)
    
    If RandomNumber(1, 100) <= Suerte Then
    
        daño = Int(daño * 0.75)
        
        If VictimUserIndex <> 0 Then
            
            With UserList(VictimUserIndex)
                .Stats.MinHP = .Stats.MinHP - daño
                Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a " & .Name & " por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
                Call WriteConsoleMsg(VictimUserIndex, UserList(UserIndex).Name & " te ha golpeado críticamente por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
            End With
            
        Else
        
            Npclist(VictimNpcIndex).Stats.MinHP = Npclist(VictimNpcIndex).Stats.MinHP - daño
            Call WriteConsoleMsg(UserIndex, "Has golpeado críticamente a la criatura por " & daño & ".", FontTypeNames.FONTTYPE_FIGHT)
            Call CalcularDarExp(UserIndex, VictimNpcIndex, daño)
            
        End If
        
    End If

End Sub

Public Sub QuitarSta(ByVal UserIndex As Integer, ByVal Cantidad As Integer)

On Error GoTo Errhandler

    UserList(UserIndex).Stats.MinSta = UserList(UserIndex).Stats.MinSta - Cantidad
    If UserList(UserIndex).Stats.MinSta < 0 Then UserList(UserIndex).Stats.MinSta = 0
    Call WriteUpdateSta(UserIndex)
    
Exit Sub

Errhandler:
    Call LogError("Error en QuitarSta. Error " & Err.Number & " : " & Err.Description)
    
End Sub

Public Sub DoTalar(ByVal UserIndex As Integer, Optional ByVal DarMaderaElfica As Boolean = False)
'***************************************************
'Autor: Unknown
'Last Modification: 16/11/2009
'16/11/2009: ZaMa - Ahora Se puede dar madera elfica.
'16/11/2009: ZaMa - Implementado nuevo sistema de extraccion.
'***************************************************
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer
Dim CantidadItems As Integer

If UserList(UserIndex).clase = eClass.Worker Then
    Call QuitarSta(UserIndex, EsfuerzoTalarLeñador)
Else
    Call QuitarSta(UserIndex, EsfuerzoTalarGeneral)
End If

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Talar)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 6 Then
    Dim MiObj As Obj
    
    If UserList(UserIndex).clase = eClass.Worker Then
        With UserList(UserIndex)
            CantidadItems = MaximoInt(1, CInt(.Stats.ELV / 2))
        End With
        
        MiObj.Amount = RandomNumber(MaximoInt(1, CInt(UserList(UserIndex).Stats.ELV / 5)), CantidadItems)
    Else
        MiObj.Amount = RandomNumber(1, 2)
    End If
    
    MiObj.ObjIndex = IIf(DarMaderaElfica, LeñaElfica, Leña)
    
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
        
    End If
    
    Call WriteConsoleMsg(UserIndex, "¡Has conseguido algo de leña!", FontTypeNames.FONTTYPE_INFO)
    
    Call SubirSkill(UserIndex, eSkill.Talar)
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 8 Then
        Call WriteConsoleMsg(UserIndex, "¡No has obtenido leña!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 8
    End If
    '[/CDT]
    Call SubirSkill(UserIndex, eSkill.Talar)
End If

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en DoTalar")

End Sub
Public Sub DoMineria(ByVal UserIndex As Integer)
On Error GoTo Errhandler

Dim Suerte As Integer
Dim res As Integer

If UserList(UserIndex).clase = eClass.Worker Then
    Call QuitarSta(UserIndex, EsfuerzoExcavarMinero)
Else
    Call QuitarSta(UserIndex, EsfuerzoExcavarGeneral)
End If

Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageAtaca(UserIndex, 1))

Dim Skill As Integer
Skill = UserList(UserIndex).Stats.UserSkills(eSkill.Mineria)
Suerte = Int(-0.00125 * Skill * Skill - 0.3 * Skill + 49)

res = RandomNumber(1, Suerte)

If res <= 5 Then
    Dim MiObj As Obj
    
    If UserList(UserIndex).flags.TargetObj = 0 Then Exit Sub
    
    MiObj.ObjIndex = ObjData(UserList(UserIndex).flags.TargetObj).MineralIndex
    
    If UserList(UserIndex).clase = eClass.Worker Then
        MiObj.Amount = RandomNumber(1, MaximoInt(1, CInt(UserList(UserIndex).Stats.ELV / 3))) '(NicoNZ) 04/25/2008
    Else
        MiObj.Amount = 1
    End If
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then _
        Call TirarItemAlPiso(UserList(UserIndex).Pos, MiObj)
    
    Call WriteConsoleMsg(UserIndex, "¡Has extraido algunos minerales!", FontTypeNames.FONTTYPE_INFO)
    
Else
    '[CDT 17-02-2004]
    If Not UserList(UserIndex).flags.UltimoMensaje = 9 Then
        Call WriteConsoleMsg(UserIndex, "¡No has conseguido nada!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 9
    End If
    '[/CDT]
End If

Call SubirSkill(UserIndex, Mineria)

UserList(UserIndex).Reputacion.PlebeRep = UserList(UserIndex).Reputacion.PlebeRep + vlProleta
If UserList(UserIndex).Reputacion.PlebeRep > MAXREP Then _
    UserList(UserIndex).Reputacion.PlebeRep = MAXREP

UserList(UserIndex).Counters.Trabajando = UserList(UserIndex).Counters.Trabajando + 1

Exit Sub

Errhandler:
    Call LogError("Error en Sub DoMineria")

End Sub

Public Sub DoMeditar(ByVal UserIndex As Integer)

UserList(UserIndex).Counters.IdleCount = 0

Dim Suerte As Integer
Dim res As Integer
Dim Cant As Integer

'Barrin 3/10/03
'Esperamos a que se termine de concentrar
Dim TActual As Long
TActual = (GetTickCount() And &H7FFFFFFF)
If TActual - UserList(UserIndex).Counters.tInicioMeditar < TIEMPO_INICIOMEDITAR Then
    Exit Sub
End If

If UserList(UserIndex).Counters.bPuedeMeditar = False Then
    UserList(UserIndex).Counters.bPuedeMeditar = True
End If
    
If UserList(UserIndex).Stats.MinMAN >= UserList(UserIndex).Stats.MaxMAN Then
    Call WriteConsoleMsg(UserIndex, "Has terminado de meditar.", FontTypeNames.FONTTYPE_INFO)
    Call WriteMeditateToggle(UserIndex)
    UserList(UserIndex).flags.Meditando = False
    UserList(UserIndex).Char.FX = 0
    UserList(UserIndex).Char.Loops = 0
    Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessageCreateFX(UserList(UserIndex).Char.CharIndex, 0, 0))
    Exit Sub
End If

If UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 10 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= -1 Then
                    Suerte = 25
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 20 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 11 Then
                    Suerte = 24
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 30 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 21 Then
                    Suerte = 23
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 40 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 31 Then
                    Suerte = 22
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 50 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 41 Then
                    Suerte = 20
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 60 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 51 Then
                    Suerte = 18
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 70 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 61 Then
                    Suerte = 15
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 80 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 71 Then
                    Suerte = 12
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) <= 90 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 81 Then
                    Suerte = 10
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) < 100 _
   And UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) >= 91 Then
                    Suerte = 7
ElseIf UserList(UserIndex).Stats.UserSkills(eSkill.Meditar) = 100 Then
                    Suerte = 5
End If
res = RandomNumber(1, Suerte)

If res = 1 Then
    
    Cant = Porcentaje(UserList(UserIndex).Stats.MaxMAN, PorcentajeRecuperoMana)
    If Cant <= 0 Then Cant = 1
    UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MinMAN + Cant
    If UserList(UserIndex).Stats.MinMAN > UserList(UserIndex).Stats.MaxMAN Then _
        UserList(UserIndex).Stats.MinMAN = UserList(UserIndex).Stats.MaxMAN
    
    If Not UserList(UserIndex).flags.UltimoMensaje = 22 Then
        Call WriteConsoleMsg(UserIndex, "¡Has recuperado " & Cant & " puntos de mana!", FontTypeNames.FONTTYPE_INFO)
        UserList(UserIndex).flags.UltimoMensaje = 22
    End If
    
    Call WriteUpdateMana(UserIndex)
    Call SubirSkill(UserIndex, Meditar)
End If

End Sub

Public Sub DoHurtar(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 03/03/2010
'Implements the pick pocket skill of the Bandit :)
'03/03/2010 - Pato: Sólo se puede hurtar si no está en trigger 6 :)
'***************************************************
Dim OtroUserIndex As Integer

If TriggerZonaPelea(UserIndex, VictimaIndex) <> TRIGGER6_AUSENTE Then Exit Sub

If UserList(UserIndex).clase <> eClass.Bandit Then Exit Sub
'Esto es precario y feo, pero por ahora no se me ocurrió nada mejor.
'Uso el slot de los anillos para "equipar" los guantes.
'Y los reconozco porque les puse DefensaMagicaMin y Max = 0
If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub

Dim res As Integer
res = RandomNumber(1, 100)
If (res < 20) Then
    If TieneObjetosRobables(VictimaIndex) Then
    
        If UserList(VictimaIndex).flags.Comerciando Then
            OtroUserIndex = UserList(VictimaIndex).ComUsu.DestUsu
                
            If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
                Call WriteConsoleMsg(VictimaIndex, "¡¡Comercio cancelado, te están robando!!", FontTypeNames.FONTTYPE_TALK)
                Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
                
                Call LimpiarComercioSeguro(VictimaIndex)
                Call Protocol.FlushBuffer(OtroUserIndex)
            End If
        End If
                
        Call RobarObjeto(UserIndex, VictimaIndex)
        Call WriteConsoleMsg(VictimaIndex, "¡" & UserList(UserIndex).Name & " es un Bandido!", FontTypeNames.FONTTYPE_INFO)
    Else
        Call WriteConsoleMsg(UserIndex, UserList(VictimaIndex).Name & " no tiene objetos.", FontTypeNames.FONTTYPE_INFO)
    End If
End If
End Sub

Public Sub DoHandInmo(ByVal UserIndex As Integer, ByVal VictimaIndex As Integer)
'***************************************************
'Author: Pablo (ToxicWaste)
'Last Modif: 17/02/2007
'Implements the special Skill of the Thief
'***************************************************
    If UserList(VictimaIndex).flags.Paralizado = 1 Then Exit Sub
    If UserList(UserIndex).clase <> eClass.Thief Then Exit Sub
        
    
    If UserList(UserIndex).Invent.AnilloEqpObjIndex <> GUANTE_HURTO Then Exit Sub
        
    Dim res As Integer
    res = RandomNumber(0, 100)
    If res < (UserList(UserIndex).Stats.UserSkills(eSkill.Wrestling) / 4) Then
        UserList(VictimaIndex).flags.Paralizado = 1
        UserList(VictimaIndex).Counters.Paralisis = IntervaloParalizado / 2
        
        'UserList(VictimaIndex).flags.ParalizedByIndex = UserIndex
        'UserList(VictimaIndex).flags.ParalizedBy = UserList(UserIndex).Name
        
        Call WriteParalizeOK(VictimaIndex)
        Call WriteConsoleMsg(UserIndex, "Tu golpe ha dejado inmóvil a tu oponente", FontTypeNames.FONTTYPE_INFO)
        Call WriteConsoleMsg(VictimaIndex, "¡El golpe te ha dejado inmóvil!", FontTypeNames.FONTTYPE_INFO)
    End If


End Sub

Public Sub Desarmar(ByVal UserIndex As Integer, ByVal VictimIndex As Integer)

'***************************************************
'Author: Unknown
'Last Modification: 02/04/2010 (ZaMa)
'02/04/2010: ZaMa - Nueva formula para desarmar.
'***************************************************

    Dim Probabilidad As Integer
    Dim Resultado As Integer
    Dim WrestlingSkill As Byte
    
    With UserList(UserIndex)
        WrestlingSkill = .Stats.UserSkills(eSkill.Wrestling)
        
        Probabilidad = WrestlingSkill * 0.2 + .Stats.ELV * 0.66
        
        Resultado = RandomNumber(1, 100)
        
        If Resultado <= Probabilidad Then
            Call Desequipar(VictimIndex, UserList(VictimIndex).Invent.WeaponEqpSlot, True)
            Call WriteConsoleMsg(UserIndex, "Has logrado desarmar a tu oponente!", FontTypeNames.FONTTYPE_FIGHT)
            If UserList(VictimIndex).Stats.ELV < 20 Then
                Call WriteConsoleMsg(VictimIndex, "¡Tu oponente te ha desarmado!", FontTypeNames.FONTTYPE_FIGHT)
            End If
            Call FlushBuffer(VictimIndex)
        End If
    End With
    

End Sub
Public Sub DoUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'12/08/2009: Pato - Implementado nuevo sistema de mejora de items
'***************************************************
Dim ItemUpgrade As Integer
Dim WeaponIndex As Integer
Dim OtroUserIndex As Integer

ItemUpgrade = ObjData(ItemIndex).Upgrade

With UserList(UserIndex)
    If .flags.Comerciando Then
        OtroUserIndex = .ComUsu.DestUsu
            
        If OtroUserIndex > 0 And OtroUserIndex <= MaxUsers Then
            Call WriteConsoleMsg(UserIndex, "¡¡Comercio cancelado, no puedes comerciar mientras trabajas!!", FontTypeNames.FONTTYPE_TALK)
            Call WriteConsoleMsg(OtroUserIndex, "¡¡Comercio cancelado por el otro usuario!!", FontTypeNames.FONTTYPE_TALK)
            
            Call LimpiarComercioSeguro(UserIndex)
            Call Protocol.FlushBuffer(OtroUserIndex)
        End If
    End If
        
    'Sacamos energía
    If .clase = eClass.Worker Then
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA_TRABAJADOR Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_TRABAJADOR
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    Else
        'Chequeamos que tenga los puntos antes de sacarselos
        If .Stats.MinSta >= GASTO_ENERGIA_NO_TRABAJADOR Then
            .Stats.MinSta = .Stats.MinSta - GASTO_ENERGIA_NO_TRABAJADOR
            Call WriteUpdateSta(UserIndex)
        Else
            Call WriteConsoleMsg(UserIndex, "No tienes suficiente energía.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
    End If
    
    If ItemUpgrade <= 0 Then Exit Sub
    If Not TieneMaterialesUpgrade(UserIndex, ItemIndex) Then Exit Sub
    
    If PuedeConstruirHerreria(ItemUpgrade) Then
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
    
        If WeaponIndex <> MARTILLO_HERRERO Then
            Call WriteConsoleMsg(UserIndex, "Debes equiparte el martillo de herrero.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Round(.Stats.UserSkills(eSkill.Herreria) / ModHerreriA(.clase), 0) < ObjData(ItemUpgrade).SkHerreria Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otESCUDO 'Todavía no hay, pero just in case
                Call WriteConsoleMsg(UserIndex, "Has mejorado el escudo!", FontTypeNames.FONTTYPE_INFO)
            
            Case eOBJType.otCASCO
                Call WriteConsoleMsg(UserIndex, "Has mejorado el casco!", FontTypeNames.FONTTYPE_INFO)
            
            Case eOBJType.otArmadura
                Call WriteConsoleMsg(UserIndex, "Has mejorado la armadura!", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Call SubirSkill(UserIndex, eSkill.Herreria)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(MARTILLOHERRERO, .Pos.X, .Pos.Y))
    
    ElseIf PuedeConstruirCarpintero(ItemUpgrade) Then
        
        WeaponIndex = .Invent.WeaponEqpObjIndex
        If WeaponIndex <> SERRUCHO_CARPINTERO Then
            Call WriteConsoleMsg(UserIndex, "Debes equiparte un serrucho.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        If Round(.Stats.UserSkills(eSkill.Carpinteria) \ ModCarpinteria(.clase), 0) < ObjData(ItemUpgrade).SkCarpinteria Then
            Call WriteConsoleMsg(UserIndex, "No tienes suficientes skills.", FontTypeNames.FONTTYPE_INFO)
            Exit Sub
        End If
        
        Select Case ObjData(ItemIndex).OBJType
            Case eOBJType.otFlechas
                Call WriteConsoleMsg(UserIndex, "Has mejorado la flecha!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otWeapon
                Call WriteConsoleMsg(UserIndex, "Has mejorado el arma!", FontTypeNames.FONTTYPE_INFO)
                
            Case eOBJType.otBarcos
                Call WriteConsoleMsg(UserIndex, "Has mejorado el barco!", FontTypeNames.FONTTYPE_INFO)
        End Select
        
        Call SubirSkill(UserIndex, eSkill.Carpinteria)
        Call SendData(SendTarget.ToPCArea, UserIndex, PrepareMessagePlayWave(LABUROCARPINTERO, .Pos.X, .Pos.Y))
    Else
        Exit Sub
    End If
    
    Call QuitarMaterialesUpgrade(UserIndex, ItemIndex)
    
    Dim MiObj As Obj
    MiObj.Amount = 1
    MiObj.ObjIndex = ItemUpgrade
    
    If Not MeterItemEnInventario(UserIndex, MiObj) Then
        Call TirarItemAlPiso(.Pos, MiObj)
    End If
    
    If ObjData(ItemIndex).Log = 1 Then _
        Call LogDesarrollo(.Name & " ha mejorado el ítem " & ObjData(ItemIndex).Name & " a " & ObjData(ItemUpgrade).Name)
        
    .Reputacion.PlebeRep = .Reputacion.PlebeRep + vlProleta
    If .Reputacion.PlebeRep > MAXREP Then _
        .Reputacion.PlebeRep = MAXREP
        
    .Counters.Trabajando = .Counters.Trabajando + 1
End With
End Sub
Function TieneMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer) As Boolean
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'
'***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then
            If Not TieneObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de hierro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingP > 0 Then
            If Not TieneObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de plata.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .LingO > 0 Then
            If Not TieneObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficientes lingotes de oro.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .Madera > 0 Then
            If Not TieneObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
        
        If .MaderaElfica > 0 Then
            If Not TieneObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex) Then
                Call WriteConsoleMsg(UserIndex, "No tienes suficiente madera élfica.", FontTypeNames.FONTTYPE_INFO)
                TieneMaterialesUpgrade = False
                Exit Function
            End If
        End If
    End With
    
    TieneMaterialesUpgrade = True
End Function

Sub QuitarMaterialesUpgrade(ByVal UserIndex As Integer, ByVal ItemIndex As Integer)
'***************************************************
'Author: Torres Patricio (Pato)
'Last Modification: 12/08/2009
'
'***************************************************
    Dim ItemUpgrade As Integer
    
    ItemUpgrade = ObjData(ItemIndex).Upgrade
    
    With ObjData(ItemUpgrade)
        If .LingH > 0 Then Call QuitarObjetos(LingoteHierro, CInt(.LingH - ObjData(ItemIndex).LingH * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        If .LingP > 0 Then Call QuitarObjetos(LingotePlata, CInt(.LingP - ObjData(ItemIndex).LingP * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        If .LingO > 0 Then Call QuitarObjetos(LingoteOro, CInt(.LingO - ObjData(ItemIndex).LingO * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        If .Madera > 0 Then Call QuitarObjetos(Leña, CInt(.Madera - ObjData(ItemIndex).Madera * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
        If .MaderaElfica > 0 Then Call QuitarObjetos(LeñaElfica, CInt(.MaderaElfica - ObjData(ItemIndex).MaderaElfica * PORCENTAJE_MATERIALES_UPGRADE), UserIndex)
    End With
    
    Call QuitarObjetos(ItemIndex, 1, UserIndex)
End Sub

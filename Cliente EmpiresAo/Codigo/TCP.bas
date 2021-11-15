Attribute VB_Name = "Mod_TCP"
'Argentum Online 0.9.0.9
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
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
Public Warping As Boolean
Public LlegaronSkills As Boolean
Public LlegaronAtrib As Boolean
Public LlegoFama As Boolean



Public Function PuedoQuitarFoco() As Boolean
PuedoQuitarFoco = True
'PuedoQuitarFoco = Not frmEstadisticas.Visible And _
'                 Not frmGuildAdm.Visible And _
'                 Not frmGuildDetails.Visible And _
'                 Not frmGuildBrief.Visible And _
'                 Not frmGuildFoundation.Visible And _
'                 Not frmGuildLeader.Visible And _
'                 Not frmCharInfo.Visible And _
'                 Not frmGuildNews.Visible And _
'                 Not frmGuildSol.Visible And _
'                 Not frmCommet.Visible And _
'                 Not frmPeaceProp.Visible
'
End Function

Sub HandleData(ByVal Rdata As String)
    On Error Resume Next
    
    Dim retVal As Variant
    Dim x As Integer
    Dim y As Integer
    Dim CharIndex As Integer
    Dim tempint As Integer
    Dim tempstr As String
    Dim slot As Integer
    Dim MapNumber As String
    Dim I As Integer, k As Integer
    Dim cad$, index As Integer, m As Integer
    Dim T() As String
    Dim result As Long
    
    Dim tstr As String
    Dim tstr2 As String
    
    
    Dim sData As String
    sData = UCase$(Rdata)
    
    Select Case Left(Rdata, 1)
        Case "ç"
        Rdata = Right$(Rdata, Len(Rdata) - 1)
        Rdata = Decrypt_PRO(Rdata, CRYPTO_PASS)
        sData = UCase$(Rdata)
        'Debug.Print Rdata
    End Select
    
    Select Case sData
        Case "LOGGED"            ' >>>>> LOGIN :: LOGGED
            logged = True
            'Call SetWindowLong(frmMain.RecTxt.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
            UserCiego = False
            EngineRun = True
            'frmCuent.Visible = False
            IScombate = False
            UserDescansar = False
            Nombres = True
            
            If frmCrearPersonaje.Visible Then
                'Unload frmPasswdSinPadrinos
                Unload frmCrearPersonaje
                'Unload frmCuent
                frmConnect.Visible = False
                frmMain.Show
                'frmMain.SetFocus
            End If
Call SetConnected
            'Mostramos el Tip
            If tipf = "1" And PrimeraVez Then
                 Call CargarTip
                 frmtip.Visible = True
                 PrimeraVez = False
            End If
            'result =
            
            bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            Exit Sub
        
        Case "NOTVERSION"
            If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
            MsgBox "La versión de su cliente no es actual o se encuentra dañada." & vbCrLf & "Ingrese a la sección descargas de la web oficial www.empiresao.com.ar para descargar posibles parches que hayan sido publicados.", vbCritical
            'Shell App.Path & "\Autoupdate.exe"
            'Call MatarAO
            Exit Sub
            
        Case "QTDL"              ' >>>>> Quitar Dialogos :: QTDL
            Call Dialogos.BorrarDialogos
            Exit Sub
        Case "NAVEG"
            UserNavegando = Not UserNavegando
            Exit Sub
        Case "FINOK" ' Graceful exit ;))

            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
            frmMain.Visible = False
            logged = False
            Call SetMusicInfo("", "", "", , , False)
            UserParalizado = False
            IScombate = False
            pausa = False
            UserMeditar = False
            frmMain.Timer1.Enabled = False
            UserDescansar = False
            UserNavegando = False
            frmConnect.Visible = True
            Unload frmCuent
            frmConnect.MousePointer = 99
            Call Audio.StopWave
            frmMain.IsPlaying = PlayLoop.plNone
            bRain = False
            bFogata = False
            SkillPoints = 0
            frmMain.Label1.Visible = False
            Call Dialogos.BorrarDialogos
            For I = 1 To LastChar
                charlist(I).invisible = False
            Next I
            
#If SeguridadAlkon Then
            Call MI(CualMI).Inicializar(RandomNumber(1, 1000), 10000)
#End If
            
            bK = 0
            Exit Sub
        Case "FINCOMOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmComerciar.List1(0).Clear
            frmComerciar.List1(1).Clear
            NPCInvDim = 0
            Unload frmComerciar
            Comerciando = False
            Exit Sub
        Case "FINSUBOK"          ' >>>>> Finaliza Comerciar :: FINCOMOK
            frmSubasta.List1(1).Clear
            Unload frmSubasta
            Exit Sub
        '[KEVIN]**************************************************************
        '-----------------------------------------------------------------------------
        Case "FINBANOK"          ' >>>>> Finaliza Banco :: FINBANOK
            frmBancoObj.List1(0).Clear
            frmBancoObj.List1(1).Clear
            NPCInvDim = 0
            Unload frmBancoObj
            Comerciando = False
            Exit Sub
        '[/KEVIN]***********************************************************************
        '------------------------------------------------------------------------------
        Case "INITSUB"           ' >>>>> Inicia Comerciar :: INITCOM
            I = 1
            Do While I <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(I) <> 0 Then
                        frmSubasta.List1(1).AddItem Inventario.ItemName(I)
                Else
                        frmSubasta.List1(1).AddItem "Nada"
                End If
                I = I + 1
            Loop
            'Comerciando = True
            frmSubasta.Show , frmMain
            Exit Sub
        Case "INITCOM"           ' >>>>> Inicia Comerciar :: INITCOM
            I = 1
            Do While I <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(I)
                Else
                        frmComerciar.List1(1).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmComerciar.Show , frmMain
            Exit Sub
        '[KEVIN]-----------------------------------------------
        '**************************************************************
        Case "INITBANCO"           ' >>>>> Inicia Comerciar :: INITBANCO
            Dim II As Integer
            II = 1
            Do While II <= MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(II) <> 0 Then
                        frmBancoObj.List1(1).AddItem Inventario.ItemName(II)
                Else
                        frmBancoObj.List1(1).AddItem "Nada"
                End If
                II = II + 1
            Loop
            
            
            I = 1
            Do While I <= UBound(UserBancoInventory)
                If UserBancoInventory(I).OBJIndex <> 0 Then
                        frmBancoObj.List1(0).AddItem UserBancoInventory(I).Name
                Else
                        frmBancoObj.List1(0).AddItem "Nada"
                End If
                I = I + 1
            Loop
            Comerciando = True
            frmBancoObj.Show , frmMain
            Exit Sub
        '---------------------------------------------------------------
        '[/KEVIN]******************
        '[Alejo]
        Case "INITCOMUSU"
            If frmComerciarUsu.List1.ListCount > 0 Then frmComerciarUsu.List1.Clear
            If frmComerciarUsu.List2.ListCount > 0 Then frmComerciarUsu.List2.Clear
            
            For I = 1 To MAX_INVENTORY_SLOTS
                If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciarUsu.List1.AddItem Inventario.ItemName(I)
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = Inventario.Amount(I)
                Else
                        frmComerciarUsu.List1.AddItem "Nada"
                        frmComerciarUsu.List1.ItemData(frmComerciarUsu.List1.NewIndex) = 0
                End If
            Next I
            Comerciando = True
            frmComerciarUsu.Show , frmMain
        Case "FINCOMUSUOK"
            frmComerciarUsu.List1.Clear
            frmComerciarUsu.List2.Clear
            
            Unload frmComerciarUsu
            Comerciando = False
            '[/Alejo]
        Case "RECPASSOK"
            Call MsgBox("¡¡¡El password fue enviado con éxito!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "RECPASSER"
            Call MsgBox("¡¡¡No coinciden los datos con los del personaje en el servidor, el password no ha sido enviado.!!!", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Envio de password")
            frmRecuperar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmRecuperar
            Exit Sub
        Case "BORROK"
            Call MsgBox("El personaje ha sido borrado.", vbApplicationModal + vbDefaultButton1 + vbInformation + vbOKOnly, "Borrado de personaje")
            frmBorrar.MousePointer = 0
#If UsarWrench = 1 Then
            frmMain.Socket1.Disconnect
#Else
            If frmMain.Winsock1.State <> sckClosed Then _
                frmMain.Winsock1.Close
#End If
            Unload frmBorrar
            Exit Sub
        Case "SFH"
            frmHerrero.Show , frmMain
            Exit Sub
        Case "SFC"
            frmCarp.Show , frmMain
            Exit Sub
        Case "N1" ' <--- Npc ataco y fallo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_FALLA_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "6" ' <--- Npc mata al usuario
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_CRIATURA_MATADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "7" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "8" ' <--- Ataque rechazado con el escudo
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USUARIO_RECHAZO_ATAQUE_ESCUDO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U1" ' <--- User ataco y fallo el golpe
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_FALLADO_GOLPE, 255, 0, 0, True, False, False)
            Exit Sub
        Case "REAU" '<--- Requiere AutoUpdate
            'Call frmMain.DibujarSatelite
            Exit Sub
        Case "SEGON" '  <--- Activa el seguro
            Call frmMain.DibujarSeguro
            frmMain.Label2(1).ForeColor = &H8000&
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_ACTIVADO, 0, 255, 0, True, False, False)
            Exit Sub
        Case "SEGOFF" ' <--- Desactiva el seguro
            Call frmMain.DesDibujarSeguro
            frmMain.Label2(1).ForeColor = vbRed
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_SEGURO_DESACTIVADO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "PN"     ' <--- Pierde Nobleza
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PIERDE_NOBLEZA, 255, 0, 0, False, False, False)
            Exit Sub
        Case "M!"     ' <--- Usa meditando
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_USAR_MEDITANDO, 255, 0, 0, False, False, False)
            Exit Sub
    End Select

    Select Case Left(sData, 1)
        Case "+"              ' >>>>> Mover Char >>> +
            Rdata = Right$(Rdata, Len(Rdata) - 1)

#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveCharDecrypt(Rdata, CharIndex, x, y)
#Else
            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            x = Val(ReadField(2, Rdata, Asc(",")))
            y = Val(ReadField(3, Rdata, Asc(",")))
#End If

            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4

            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).fX >= 40 And charlist(CharIndex).fX <= 49 Then   'si esta meditando
                charlist(CharIndex).fX = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If

            Call MoveCharbyPos(CharIndex, x, y)
            
            Call RefreshAllChars
            Exit Sub
        Case "*", "_"             ' >>>>> Mover NPC >>> *
            Rdata = Right$(Rdata, Len(Rdata) - 1)
            
#If SeguridadAlkon Then
            'obtengo todo
            Call CheatingDeath.MoveNPCDecrypt(Rdata, CharIndex, x, y, Left$(sData, 1) <> "*")
#Else
            CharIndex = Val(ReadField(1, Rdata, Asc(",")))
            x = Val(ReadField(2, Rdata, Asc(",")))
            y = Val(ReadField(3, Rdata, Asc(",")))
#End If
            
            'antigua codificacion del mensaje (decodificada x un chitero)
            'CharIndex = Asc(Mid$(Rdata, 1, 1)) * 64 + (Asc(Mid$(Rdata, 2, 1)) And &HFC&) / 4
            
'            If charlist(CharIndex).Body.Walk(1).GrhIndex = 4747 Then
'                Debug.Print "hola"
'            End If
            
            ' CONSTANTES TODO: De donde sale el 40-49 ?
            
            If charlist(CharIndex).fX >= 40 And charlist(CharIndex).fX <= 49 Then   'si esta meditando
                charlist(CharIndex).fX = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            ' CONSTANTES TODO: Que es .priv ?
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, x, y)
            'Call MoveCharbyPos(CharIndex, Asc(Mid$(Rdata, 3, 1)), Asc(Mid$(Rdata, 4, 1)))
            
            Call RefreshAllChars
            Exit Sub
    
    End Select

    Select Case Left$(sData, 2)
        Case "AS"
            tstr = mid$(sData, 3, 1)
            k = Val(Right$(sData, Len(sData) - 3))
            
            Select Case tstr
                Case "M": UserMinMAN = Val(Right$(sData, Len(sData) - 3))
                Case "H": UserMinHP = Val(Right$(sData, Len(sData) - 3))
                Case "S": UserMinSTA = Val(Right$(sData, Len(sData) - 3))
                Case "G": UserGLD = Val(Right$(sData, Len(sData) - 3))
                Case "E": UserExp = Val(Right$(sData, Len(sData) - 3))
            End Select
            
            frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            frmMain.lblSTA.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.lblMANA.Caption = IIf(UserMaxMAN > 0, UserMinMAN & "/" & UserMaxMAN, "")
            frmMain.lblHP.Caption = UserMinHP & "/" & UserMaxHP
            
            Exit Sub
        Case "CM"              ' >>>>> Cargar Mapa :: CM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserMap = ReadField(1, Rdata, 44)
            UserMapName = ReadField(3, Rdata, 44)
            'Obtiene la version del mapa

#If SeguridadAlkon Then
            Call InitMI
#End If
            
            If FileExist(DirMapas & "Mapa" & UserMap & ".map", vbNormal) Then
                Open DirMapas & "Mapa" & UserMap & ".map" For Binary As #1
                Seek #1, 1
                Get #1, , tempint
                Close #1
'                If tempint = Val(ReadField(2, Rdata, 44)) Then
                    'Si es la vers correcta cambiamos el mapa
                    Call SwitchMap(UserMap)
                    If bLluvia(UserMap) = 0 Then
                        If bRain Then
                            Call Audio.StopWave(RainBufferIndex)
                            RainBufferIndex = 0
                            frmMain.IsPlaying = PlayLoop.plNone
                        End If
                    End If
'                Else
'                    'vers incorrecta
'                    MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
'                    Call LiberarObjetosDX
'                    Call UnloadAllForms
'                    End
'                End If
            Else
                'no encontramos el mapa en el hd
                MsgBox "Error en los mapas, algun archivo ha sido modificado o esta dañado."
                Call LiberarObjetosDX
                Call UnloadAllForms
                Call EscribirGameIni(Config_Inicio)
                End
            End If
            Exit Sub
        
        Case "PU"                 ' >>>>> Actualiza Posición Usuario :: PU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(UserPos.x, UserPos.y).CharIndex = 0
            UserPos.x = CInt(ReadField(1, Rdata, 44))
            UserPos.y = CInt(ReadField(2, Rdata, 44))
            MapData(UserPos.x, UserPos.y).CharIndex = UserCharIndex
            charlist(UserCharIndex).Pos = UserPos
            frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
            Exit Sub
        
        Case "N2" ' <<--- Npc nos impacto (Ahorramos ancho de banda)
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "U2" ' <<--- El user ataco un npc e impacato
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_GOLPE_CRIATURA_1 & Rdata & MENSAJE_2, 255, 0, 0, True, False, False)
            Exit Sub
        Case "U3" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & Rdata & MENSAJE_ATAQUE_FALLO, 255, 0, 0, True, False, False)
            Exit Sub
        Case "II" ' <<--- El user ataco un user y falla
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            frmProcesos.Show , frmMain
            frmProcesos.List1.AddItem Rdata
            Exit Sub
        Case "N4" ' <<--- user nos impacto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_1 & ReadField(3, Rdata, 44) & MENSAJE_RECIVE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "N5" ' <<--- impactamos un user
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            I = Val(ReadField(1, Rdata, 44))
            Select Case I
                Case bCabeza
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_CABEZA & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoIzquierdo
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bBrazoDerecho
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_BRAZO_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaIzquierda
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_IZQ & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bPiernaDerecha
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_PIERNA_DER & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
                Case bTorso
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_PRODUCE_IMPACTO_1 & ReadField(3, Rdata, 44) & MENSAJE_PRODUCE_IMPACTO_TORSO & Val(ReadField(2, Rdata, 44)), 255, 0, 0, True, False, False)
            End Select
            Exit Sub
        Case "||"                 ' >>>>> Dialogo de Usuarios y NPCs :: ||
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Dim iuser As Integer
            iuser = Val(ReadField(3, Rdata, 176))
            
            If iuser > 0 Then
                Dialogos.CrearDialogo ReadField(2, Rdata, 176), iuser, Val(ReadField(1, Rdata, 176))
                'AddtoRichTextBox frmMain.RecTxt, charlist(iuser).Nombre & "> " & ReadField(2, Rdata, 176), 228, 177, 3, 0, 0
            Else
                If PuedoQuitarFoco Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                End If
            End If

            Exit Sub
        Case "|+"                 ' >>>>> Consola de clan y NPCs :: |+
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            iuser = Val(ReadField(3, Rdata, 176))

            If iuser = 0 Then
                If PuedoQuitarFoco And Not DialogosClanes.Activo Then
                    AddtoRichTextBox frmMain.RecTxt, ReadField(1, Rdata, 126), Val(ReadField(2, Rdata, 126)), Val(ReadField(3, Rdata, 126)), Val(ReadField(4, Rdata, 126)), Val(ReadField(5, Rdata, 126)), Val(ReadField(6, Rdata, 126))
                ElseIf DialogosClanes.Activo Then
                    DialogosClanes.PushBackText ReadField(1, Rdata, 126)
                End If
            End If

            Exit Sub

        Case "!!"                ' >>>>> Msgbox :: !!
            If PuedoQuitarFoco Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                frmMensaje.msg.Caption = Rdata
                frmMensaje.Show
            End If
            Exit Sub
        Case "IU"                ' >>>>> Indice de Usuario en Server :: IU
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserIndex = Val(Rdata)
            Exit Sub
        Case "IP"                ' >>>>> Indice de Personaje de Usuario :: IP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            UserCharIndex = Val(Rdata)
            UserPos = charlist(UserCharIndex).Pos
            frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
            Exit Sub
        Case "CC"              ' >>>>> Crear un Personaje :: CC
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = ReadField(4, Rdata, 44)
            x = ReadField(5, Rdata, 44)
            y = ReadField(6, Rdata, 44)
            'Debug.Print "CC"
            'If charlist(CharIndex).Pos.X Or charlist(CharIndex).Pos.Y Then
              '  Debug.Print "CHAR DUPLICADO: " & CharIndex
             '   Call EraseChar(CharIndex)
            'End If
            
            charlist(CharIndex).fX = Val(ReadField(9, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(10, Rdata, 44))
            charlist(CharIndex).Nombre = ReadField(12, Rdata, 44)
            charlist(CharIndex).Criminal = Val(ReadField(13, Rdata, 44))
            charlist(CharIndex).priv = Val(ReadField(14, Rdata, 44))
            charlist(CharIndex).aCaballo = Val(ReadField(15, Rdata, 44))
            charlist(CharIndex).invisible = Val(ReadField(16, Rdata, 44))
            
            Call MakeChar(CharIndex, ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), x, y, Val(ReadField(7, Rdata, 44)), Val(ReadField(8, Rdata, 44)), Val(ReadField(11, Rdata, 44)), charlist(CharIndex).invisible)
            Call RefreshAllChars
            Exit Sub
            
        Case "BP"             ' >>>>> Borrar un Personaje :: BP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call EraseChar(Val(Rdata))
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Call RefreshAllChars
            Exit Sub
        Case "MP"             ' >>>>> Mover un Personaje :: MP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            CharIndex = Val(ReadField(1, Rdata, 44))
            
            If charlist(CharIndex).fX >= 40 And charlist(CharIndex).fX <= 49 Then   'si esta meditando
                charlist(CharIndex).fX = 0
                charlist(CharIndex).FxLoopTimes = 0
            End If
            
            If charlist(CharIndex).priv = 0 Then
                Call DoPasosFx(CharIndex)
            End If
            
            Call MoveCharbyPos(CharIndex, ReadField(2, Rdata, 44), ReadField(3, Rdata, 44))
            
            Call RefreshAllChars
            Exit Sub
        Case "CP"             ' >>>>> Cambiar Apariencia Personaje :: CP
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).Muerto = Val(ReadField(3, Rdata, 44)) = 500
            charlist(CharIndex).Body = BodyData(Val(ReadField(2, Rdata, 44)))
            charlist(CharIndex).Head = HeadData(Val(ReadField(3, Rdata, 44)))
            charlist(CharIndex).Heading = Val(ReadField(4, Rdata, 44))
            charlist(CharIndex).fX = Val(ReadField(7, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(8, Rdata, 44))
            tempint = Val(ReadField(5, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Arma = WeaponAnimData(tempint)
            tempint = Val(ReadField(6, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Escudo = ShieldAnimData(tempint)
            tempint = Val(ReadField(9, Rdata, 44))
            If tempint <> 0 Then charlist(CharIndex).Casco = CascoAnimData(tempint)
            
            charlist(CharIndex).aCaballo = Val(ReadField(10, Rdata, 44))

            Call RefreshAllChars
            Exit Sub
        Case "HO"            ' >>>>> Crear un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            x = Val(ReadField(2, Rdata, 44))
            y = Val(ReadField(3, Rdata, 44))
            'ID DEL OBJ EN EL CLIENTE
            MapData(x, y).ObjGrh.GrhIndex = Val(ReadField(1, Rdata, 44))
            InitGrh MapData(x, y).ObjGrh, MapData(x, y).ObjGrh.GrhIndex
            Exit Sub
        Case "BO"           ' >>>>> Borrar un Objeto
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            x = Val(ReadField(1, Rdata, 44))
            y = Val(ReadField(2, Rdata, 44))
            MapData(x, y).ObjGrh.GrhIndex = 0
            Exit Sub
        Case "BQ"           ' >>>>> Bloquear Posición
            Dim b As Byte
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            MapData(Val(ReadField(1, Rdata, 44)), Val(ReadField(2, Rdata, 44))).Blocked = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "TM"           ' >>>>> Play un MIDI :: TM
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            currentMidi = Val(ReadField(1, Rdata, 45))
            
            If Musica Then
                If currentMidi <> 0 Then
                    Rdata = Right$(Rdata, Len(Rdata) - Len(ReadField(1, Rdata, 45)))
                    If Len(Rdata) > 0 Then
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid", Val(Right$(Rdata, Len(Rdata) - 1)))
                    Else
                        Call Audio.PlayMIDI(CStr(currentMidi) & ".mid")
                    End If
                End If
            End If
            Exit Sub
        Case "TW"          ' >>>>> Play un WAV :: TW
            If Sound Then
                Rdata = Right$(Rdata, Len(Rdata) - 2)
                 Call Audio.PlayWave(Rdata & ".wav")
            End If
            Exit Sub
        Case "GG"
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            SoyGM = Rdata
            Exit Sub
        Case "GL" 'Lista de guilds
            Rdata = Right$(Rdata, Len(Rdata) - 2)
            Call frmGuildAdm.ParseGuildList(Rdata)
            Exit Sub
        Case "FO"          ' >>>>> Play un WAV :: TW
            bFogata = True
            If FogataBufferIndex = 0 Then
                FogataBufferIndex = Audio.PlayWave("fuego.wav", LoopStyle.Enabled)
            End If
            Exit Sub
        Case "CA"
            CambioDeArea Asc(mid$(Rdata, 3, 1)), Asc(mid$(Rdata, 4, 1))
            Exit Sub
    End Select

    Select Case Left$(sData, 3)
        Case "VAL"                  ' >>>>> Validar Cliente :: VAL
            Dim ValString As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            bK = CLng(ReadField(1, Rdata, Asc(",")))
            bRK = ReadField(2, Rdata, Asc(","))
            ValString = ReadField(3, Rdata, Asc(","))
            CargarCabezas
            
#If SeguridadAlkon Then
            CheatingDeath.InputK
            
            If Not CheatingDeath.ValidarArchivosCriticos(ValString) Then End
#End If

            If EstadoLogin = BorrarPj Then
                Call SendData("BORR" & frmBorrar.txtNombre.Text & "," & frmBorrar.txtPasswd.Text & "," & ValidarLoginMSG(CInt(Rdata)))
            ElseIf EstadoLogin = Normal Or EstadoLogin = CrearNuevoPj Or EstadoLogin = loginaccount Then
                Call Login(ValidarLoginMSG(CInt(bRK)))
            ElseIf EstadoLogin = Dados Then
                frmCrearPersonaje.Show vbModal
            ElseIf EstadoLogin = CrearAccount Then
                frmCrearAccount.Show vbModal
            ElseIf EstadoLogin = verificaraccount Then
                frmVerificarAccount.Show vbModal
            ElseIf EstadoLogin = RecuperarAccount Then
                frmRecovery.Show vbModal
            End If
            Exit Sub
        Case "BKW"                  ' >>>>> Pausa :: BKW
            pausa = Not pausa
            Exit Sub
        Case "NIG"
        
            If nightEffect = False Then
                nightEffect = True
                Exit Sub
            End If
            
            If nightEffect = True Then
            nightEffect = False
            Exit Sub
            End If
            
            Exit Sub
        Case "LLU"                  ' >>>>> LLuvia!
            If Not InMapBounds(UserPos.x, UserPos.y) Then Exit Sub
            bTecho = IIf(MapData(UserPos.x, UserPos.y).Trigger = 1 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 2 Or _
            MapData(UserPos.x, UserPos.y).Trigger = 4, True, False)
            If Not bRain Then
                bRain = True
            Else
                If bLluvia(UserMap) <> 0 And Sound Then
                    'Stop playing the rain sound
                    Call Audio.StopWave(RainBufferIndex)
                    RainBufferIndex = 0
                    If bTecho Then
                        Call Audio.PlayWave("lluviainend.wav", LoopStyle.Disabled)
                    Else
                        Call Audio.PlayWave("lluviaoutend.wav", LoopStyle.Disabled)
                    End If
                    frmMain.IsPlaying = PlayLoop.plNone
                End If
                bRain = False
            End If
            
            Exit Sub
        Case "QDL"                  ' >>>>> Quitar Dialogo :: QDL
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call Dialogos.QuitarDialogo(Val(Rdata))
            Exit Sub
        Case "CFX"                  ' >>>>> Mostrar FX sobre Personaje :: CFX
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).fX = Val(ReadField(2, Rdata, 44))
            charlist(CharIndex).FxLoopTimes = Val(ReadField(3, Rdata, 44))
            Exit Sub
        Case "AYM"                  ' >>>>> Pone Mensaje en Cola GM :: AYM
            Dim N As String, n2 As String
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = ReadField(2, Rdata, 176)
            n2 = ReadField(1, Rdata, 176)
            frmMSG.CrearGMmSg N, n2
            frmMSG.Show , frmMain
            Exit Sub
        Case "WEP"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = Val(ReadField(1, Rdata, 44))
            n2 = Val(ReadField(2, Rdata, 44))
            frmMain.Arma(0).Caption = N
            frmMain.Arma(1).Caption = n2
            Exit Sub
        'Case "WEG"
           ' Rdata = Right$(Rdata, Len(Rdata) - 3)
           ' RcbWeight = ReadField(1, Rdata, 64)
           ' RcbWeightTotal = ReadField(2, Rdata, 64)
           ' frmMain.Weight.Caption = RcbWeight
           ' frmMain.TotalWeight.Caption = RcbWeightTotal
           ' Exit Sub
        Case "AAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = Val(ReadField(1, Rdata, 44))
            n2 = Val(ReadField(2, Rdata, 44))
            frmMain.Torso(0).Caption = N
            frmMain.Torso(1).Caption = n2
            Exit Sub
        Case "SHD"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = Val(ReadField(1, Rdata, 44))
            n2 = Val(ReadField(2, Rdata, 44))
            frmMain.Escudo(0).Caption = N
            frmMain.Escudo(1).Caption = n2
            Exit Sub
        Case "CZC"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            N = Val(ReadField(1, Rdata, 44))
            n2 = Val(ReadField(2, Rdata, 44))
            frmMain.Cabeza(0).Caption = N
            frmMain.Cabeza(1).Caption = n2
            Exit Sub
        Case "GDL"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserGLD = Val(Rdata)
            frmMain.GldLbl.Caption = UserGLD
            Exit Sub
        Case "EXP"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserPasarNivel = Val(ReadField(1, Rdata, 44))
            UserExp = Val(ReadField(2, Rdata, 44))
            frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
            Exit Sub
        Case "EST"                  ' >>>>> Actualiza Estadisticas de Usuario :: EST
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UserMaxHP = Val(ReadField(1, Rdata, 44))
            UserMinHP = Val(ReadField(2, Rdata, 44))
            UserMaxMAN = Val(ReadField(3, Rdata, 44))
            UserMinMAN = Val(ReadField(4, Rdata, 44))
            UserMaxSTA = Val(ReadField(5, Rdata, 44))
            UserMinSTA = Val(ReadField(6, Rdata, 44))
            UserGLD = Val(ReadField(7, Rdata, 44))
            UserLvl = Val(ReadField(8, Rdata, 44))
            UserPasarNivel = Val(ReadField(9, Rdata, 44))
            UserExp = Val(ReadField(10, Rdata, 44))
            frmMain.Exp.Caption = "Exp: " & UserExp & "/" & UserPasarNivel
            
            'Nivel máximo
            If UserLvl = 55 Then
                frmMain.lblPorcLvl.Caption = "Nivel máximo"
                frmMain.lblPorcLvl.ForeColor = &HFF&
            End If

            
            
            frmMain.lblPorcLvl.Caption = "[" & Round(CDbl(UserExp) * CDbl(100) / CDbl(UserPasarNivel), 2) & "%]"
            frmMain.Hpshp.Width = (((UserMinHP / 100) / (UserMaxHP / 100)) * 93)
            
            If UserMaxMAN > 0 Then
                frmMain.MANShp.Width = (((UserMinMAN + 1 / 100) / (UserMaxMAN + 1 / 100)) * 93)
            Else
                frmMain.MANShp.Width = 0
            End If
            
            frmMain.STAShp.Width = (((UserMinSTA / 100) / (UserMaxSTA / 100)) * 93)
        
            frmMain.GldLbl.Caption = UserGLD
            frmMain.LvlLbl.Caption = UserLvl
            
            If UserMinHP = 0 Then
                UserEstado = 1
            Else
                UserEstado = 0
            End If
            
            frmMain.lblSTA.Caption = UserMinSTA & "/" & UserMaxSTA
            frmMain.lblMANA.Caption = IIf(UserMaxMAN > 0, UserMinMAN & "/" & UserMaxMAN, "")
            frmMain.lblHP.Caption = UserMinHP & "/" & UserMaxHP
    
        
            Exit Sub
        Case "T01"                  ' >>>>> TRABAJANDO :: TRA
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            UsingSkill = Val(Rdata)
            frmMain.MousePointer = 2
            Select Case UsingSkill
                Case Magia
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MAGIA, 100, 100, 120, 0, 0)
                Case Pesca
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PESCA, 100, 100, 120, 0, 0)
                Case Robar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_ROBAR, 100, 100, 120, 0, 0)
                Case Talar
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_TALAR, 100, 100, 120, 0, 0)
                Case Mineria
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_MINERIA, 100, 100, 120, 0, 0)
                Case FundirMetal
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_FUNDIRMETAL, 100, 100, 120, 0, 0)
                Case Proyectiles
                    Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_TRABAJO_PROYECTILES, 100, 100, 120, 0, 0)
            End Select
            Exit Sub
        Case "PRE"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceived(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRB"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedB(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "PRT"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            Call txtReceivedT(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), ReadField(3, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), ReadField(6, Rdata, 44))
            Exit Sub
        Case "CSI"                 ' >>>>> Actualiza Slot Inventario :: CSI
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            Call Inventario.SetItem(slot, ReadField(2, Rdata, 44), ReadField(4, Rdata, 44), ReadField(5, Rdata, 44), Val(ReadField(6, Rdata, 44)), Val(ReadField(7, Rdata, 44)), _
                                    Val(ReadField(8, Rdata, 44)), Val(ReadField(9, Rdata, 44)), Val(ReadField(10, Rdata, 44)), Val(ReadField(11, Rdata, 44)), ReadField(3, Rdata, 44))
            
            Exit Sub
        '[KEVIN]-------------------------------------------------------
        '**********************************************************************
        Case "SBO"                 ' >>>>> Actualiza Inventario Banco :: SBO
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserBancoInventory(slot).OBJIndex = ReadField(2, Rdata, 44)
            UserBancoInventory(slot).Name = ReadField(3, Rdata, 44)
            UserBancoInventory(slot).Amount = ReadField(4, Rdata, 44)
            UserBancoInventory(slot).GrhIndex = Val(ReadField(5, Rdata, 44))
            UserBancoInventory(slot).OBJType = Val(ReadField(6, Rdata, 44))
            UserBancoInventory(slot).MaxHit = Val(ReadField(7, Rdata, 44))
            UserBancoInventory(slot).MinHit = Val(ReadField(8, Rdata, 44))
            UserBancoInventory(slot).Def = Val(ReadField(9, Rdata, 44))
        
            tempstr = ""
            
            If UserBancoInventory(slot).Amount > 0 Then
                tempstr = tempstr & "(" & UserBancoInventory(slot).Amount & ") " & UserBancoInventory(slot).Name
            Else
                tempstr = tempstr & UserBancoInventory(slot).Name
            End If
            
            Exit Sub
        '************************************************************************
        '[/KEVIN]-------
        Case "SHS"                ' >>>>> Agrega hechizos a Lista Spells :: SHS
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            slot = ReadField(1, Rdata, 44)
            UserHechizos(slot) = ReadField(2, Rdata, 44)
            If slot > frmMain.hlst.ListCount Then
                frmMain.hlst.AddItem ReadField(3, Rdata, 44)
            Else
                frmMain.hlst.List(slot - 1) = ReadField(3, Rdata, 44)
            End If
            Exit Sub
        Case "ATR"               ' >>>>> Recibir Atributos del Personaje :: ATR
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            For I = 1 To NUMATRIBUTOS
                UserAtributos(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronAtrib = True
            Exit Sub
        Case "LAH"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmasHerrero)
                ArmasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "LAR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ArmadurasHerrero)
                ArmadurasHerrero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ArmadurasHerrero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmHerrero.lstArmaduras.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
         Case "OBR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            
            For m = 0 To UBound(ObjCarpintero)
                ObjCarpintero(m) = 0
            Next m
            I = 1
            m = 0
            Do
                cad$ = ReadField(I, Rdata, 44)
                ObjCarpintero(m) = Val(ReadField(I + 1, Rdata, 44))
                If cad$ <> "" Then frmCarp.lstArmas.AddItem cad$
                I = I + 2
                m = m + 1
            Loop While cad$ <> ""
            Exit Sub
            
        Case "DOK"               ' >>>>> Descansar OK :: DOK
            UserDescansar = Not UserDescansar
            Exit Sub
        Case "SPL"
            Rdata = Right(Rdata, Len(Rdata) - 3)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmSpawnList.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmSpawnList.Show , frmMain
            Exit Sub
        Case "NNC"
            If frmCrearAccount.Visible Then Unload frmCrearAccount
            Exit Sub
        Case "NVC"
            If frmVerificarAccount.Visible Then Unload frmVerificarAccount
            Exit Sub
        Case "ERR"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            'frmOldPersonaje.MousePointer = 1
            frmCrearAccount.MousePointer = 99
            frmPasswdSinPadrinos.MousePointer = 99
            If Not frmCrearPersonaje.Visible Then
                If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
            End If
            frmERR.Show
            frmERR.Label1.Caption = Rdata
            Exit Sub
        Case "MSG"
            Rdata = Right$(Rdata, Len(Rdata) - 3)
            'frmCrearAccount.MousePointer = 99
            'If frmMain.Winsock1.State <> sckClosed Then frmMain.Winsock1.Close
            MsgBox Rdata
            Exit Sub
    End Select
    
    
    Select Case Left$(sData, 4)
        Case "PART"
            Call AddtoRichTextBox(frmMain.RecTxt, MENSAJE_ENTRAR_PARTY_1 & ReadField(1, Rdata, 44) & MENSAJE_ENTRAR_PARTY_2, 0, 255, 0, False, False, False)
            Exit Sub
        Case "MSTI"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
    
            frmMascota.Show , frmMain
            
            frmMascota.Nombre.Caption = ReadField(1, Rdata, 44)
            frmMascota.Vida.Caption = ReadField(6, Rdata, 44) & "/" & ReadField(7, Rdata, 44)
            frmMascota.Ataque.Caption = ReadField(8, Rdata, 44) & "/" & ReadField(9, Rdata, 44)
            frmMascota.Defensa.Caption = ReadField(5, Rdata, 44)
            frmMascota.LVL.Caption = ReadField(3, Rdata, 44)
            frmMascota.Exp.Caption = ReadField(4, Rdata, 44) & "/" & ReadField(2, Rdata, 44)


            'Call AbrirMascota(ReadField(1, Rdata, 44), ReadField(2, Rdata, 44), CLng(ReadField(3, Rdata, 44)), ReadField(4, Rdata, 44), _
            'CInt(ReadField(5, Rdata, 44)), CInt(ReadField(6, Rdata, 44)), CInt(ReadField(7, Rdata, 44)), CInt(ReadField(8, Rdata, 44)), CInt(ReadField(9, Rdata, 44)), CInt(ReadField(10, Rdata, 44)))
            Exit Sub
        Case "MSTF"
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            If Rdata = 1 Then Call DibujarTodoE(154)
            If Rdata = 2 Then Call DibujarTodoE(152)
            If Rdata = 3 Then Call DibujarTodoE(153)
            Exit Sub
        Case "CEGU"
            UserCiego = True
            Dim r As RECT
            BackBufferSurface.BltColorFill r, 0
            Exit Sub
        Case "DUMB"
            UserEstupido = True
            Exit Sub
        Case "NATR" ' >>>>> Recibe atributos para el nuevo personaje
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserAtributos(1) = ReadField(1, Rdata, 44)
            UserAtributos(2) = ReadField(2, Rdata, 44)
            UserAtributos(3) = ReadField(3, Rdata, 44)
            UserAtributos(4) = ReadField(4, Rdata, 44)
            UserAtributos(5) = ReadField(5, Rdata, 44)
            
            frmCrearPersonaje.lbFuerza.Caption = UserAtributos(1)
            frmCrearPersonaje.lbInteligencia.Caption = UserAtributos(2)
            frmCrearPersonaje.lbAgilidad.Caption = UserAtributos(3)
            frmCrearPersonaje.lbCarisma.Caption = UserAtributos(4)
            frmCrearPersonaje.lbConstitucion.Caption = UserAtributos(5)
            
            Exit Sub
        Case "MCAR"              ' >>>>> Mostrar Cartel :: MCAR
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            Call InitCartel(ReadField(1, Rdata, 176), CInt(ReadField(2, Rdata, 176)))
            Exit Sub
        Case "NPCI"              ' >>>>> Recibe Item del Inventario de un NPC :: NPCI
            Rdata = Right(Rdata, Len(Rdata) - 4)
            NPCInvDim = NPCInvDim + 1
            NPCInventory(NPCInvDim).Name = ReadField(1, Rdata, 44)
            NPCInventory(NPCInvDim).Amount = ReadField(2, Rdata, 44)
            NPCInventory(NPCInvDim).Valor = ReadField(3, Rdata, 44)
            NPCInventory(NPCInvDim).GrhIndex = ReadField(4, Rdata, 44)
            NPCInventory(NPCInvDim).OBJIndex = ReadField(5, Rdata, 44)
            NPCInventory(NPCInvDim).OBJType = ReadField(6, Rdata, 44)
            NPCInventory(NPCInvDim).MaxHit = ReadField(7, Rdata, 44)
            NPCInventory(NPCInvDim).MinHit = ReadField(8, Rdata, 44)
            NPCInventory(NPCInvDim).Def = ReadField(9, Rdata, 44)
            NPCInventory(NPCInvDim).C1 = ReadField(10, Rdata, 44)
            NPCInventory(NPCInvDim).C2 = ReadField(11, Rdata, 44)
            NPCInventory(NPCInvDim).C3 = ReadField(12, Rdata, 44)
            NPCInventory(NPCInvDim).C4 = ReadField(13, Rdata, 44)
            NPCInventory(NPCInvDim).C5 = ReadField(14, Rdata, 44)
            NPCInventory(NPCInvDim).C6 = ReadField(15, Rdata, 44)
            NPCInventory(NPCInvDim).C7 = ReadField(16, Rdata, 44)
            frmComerciar.List1(0).AddItem NPCInventory(NPCInvDim).Name
            Exit Sub
        Case "EHYS"              ' Actualiza Hambre y Sed :: EHYS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserMaxAGU = Val(ReadField(1, Rdata, 44))
            UserMinAGU = Val(ReadField(2, Rdata, 44))
            UserMaxHAM = Val(ReadField(3, Rdata, 44))
            UserMinHAM = Val(ReadField(4, Rdata, 44))
            frmMain.AGUAsp.Width = (((UserMinAGU / 100) / (UserMaxAGU / 100)) * 93)
            frmMain.COMIDAsp.Width = (((UserMinHAM / 100) / (UserMaxHAM / 100)) * 93)
            frmMain.lblAG.Caption = UserMinAGU & "/" & UserMaxAGU
            frmMain.lblCOM.Caption = UserMinHAM & "/" & UserMaxHAM
            Exit Sub
        Case "FAMA"             ' >>>>> Recibe Fama de Personaje :: FAMA
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            UserReputacion.AsesinoRep = Val(ReadField(1, Rdata, 44))
            UserReputacion.BandidoRep = Val(ReadField(2, Rdata, 44))
            UserReputacion.BurguesRep = Val(ReadField(3, Rdata, 44))
            UserReputacion.LadronesRep = Val(ReadField(4, Rdata, 44))
            UserReputacion.NobleRep = Val(ReadField(5, Rdata, 44))
            UserReputacion.PlebeRep = Val(ReadField(6, Rdata, 44))
            UserReputacion.Promedio = Val(ReadField(7, Rdata, 44))
            LlegoFama = True
            Exit Sub
        Case "MEST" ' >>>>>> Mini Estadisticas :: MEST
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            With UserEstadisticas
                .CiudadanosMatados = Val(ReadField(1, Rdata, 44))
                .CriminalesMatados = Val(ReadField(2, Rdata, 44))
                .UsuariosMatados = Val(ReadField(3, Rdata, 44))
                .NpcsMatados = Val(ReadField(4, Rdata, 44))
                .Clase = ReadField(5, Rdata, 44)
                .PenaCarcel = Val(ReadField(6, Rdata, 44))
            End With
            Exit Sub
        Case "SUNI"             ' >>>>> Subir Nivel :: SUNI
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            SkillPoints = SkillPoints + Val(Rdata)
            frmMain.Label1.Visible = True
            Exit Sub
        Case "NENE"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            AddtoRichTextBox frmMain.RecTxt, MENSAJE_NENE & Rdata, 255, 255, 255, 0, 0
            Exit Sub
        Case "IWIN"             ' >>>>> Nro de Personajes :: NENE
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            'frmProcesos.Show , frmMain
            frmProcesos.Label1.Caption = Rdata
            Exit Sub
        Case "RSOS"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmMSG.List1.AddItem Rdata
            Exit Sub
        Case "GIDN"             ' >>>>> Mensaje :: RSOS
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            If Rdata = "" Then MyGuildName = ""
            If Rdata <> "" Then MyGuildName = Rdata
            Exit Sub
        Case "MSOS"             ' >>>>> Mensaje :: MSOS
            frmMSG.Show , frmMain
            Exit Sub
        Case "FMSG"             ' >>>>> Foros :: FMSG
            Rdata = Right$(Rdata, Len(Rdata) - 4)
            frmForo.List.AddItem ReadField(1, Rdata, 176)
            frmForo.Text(frmForo.List.ListCount - 1).Text = ReadField(2, Rdata, 176)
            Load frmForo.Text(frmForo.List.ListCount)
            Exit Sub
        Case "MFOR"             ' >>>>> Foros :: MFOR
            If Not frmForo.Visible Then
                  frmForo.Show , frmMain
            End If
            Exit Sub
    End Select

    Select Case Left$(sData, 5)
        Case UCase$(Chr$(110)) & mid$("MEDOK", 4, 1) & Right$("akV", 1) & "E" & Trim$(Left$("  RS", 3))
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            CharIndex = Val(ReadField(1, Rdata, 44))
            charlist(CharIndex).invisible = (Val(ReadField(2, Rdata, 44)) = 1)
            
            Debug.Print sData
            
#If SeguridadAlkon Then
            If (10 * Val(ReadField(2, Rdata, 44)) = 10) Then
                Call MI(CualMI).SetInvisible(CharIndex)
            Else
                Call MI(CualMI).ResetInvisible(CharIndex)
            End If
#End If

            Exit Sub
        'Case "PUDVE"
           ' Rdata = Right$(Rdata, Len(Rdata) - 5)
           ' CharIndex = Val(ReadField(1, Rdata, 44))
           ' charlist(CharIndex).PuedoVerlo = (Val(ReadField(2, Rdata, 44)) = 1)
           ' Exit Sub
        Case "ZMOTD"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmCambiaMotd.Show , frmMain
            frmCambiaMotd.txtMotd.Text = Rdata
            Exit Sub
        Case "DYYSS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            frmMain.Label8.Caption = Rdata
            If MSNshow = 1 Then Call SetMusicInfo("", "", "Jugando EmpiresAO2: """ & Rdata & """ - http://www.empiresao.com.ar", , "{1}{0}", True)
            Exit Sub
            
        Case "INIAC"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            
            If ReadField(1, Rdata, 44) <> 0 Then
                frmCuent.Label4.Caption = "1"
            Else
                frmCuent.Label4.Caption = "0"
            End If
            
            frmCuent.Label6.Caption = ReadField(1, Rdata, 44)
            frmCuent.Label3.Caption = ReadField(2, Rdata, 44)
            frmCuent.Show
            'frmCuent.SetFocus
            
            Exit Sub
        
        Case "MONTA"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            userMontando = Rdata
            If userMontando = True Then IntervaloPaso = INTERVALOCABALLO
            If userMontando = False Then IntervaloPaso = INTERVALOWALK
            Exit Sub
            
        Case "ADDPJ"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            
            rcvName = ReadField(1, Rdata, 44)
            rcvIndex = ReadField(2, Rdata, 44)
            rcvHead = ReadField(3, Rdata, 44)
            rcvBody = ReadField(4, Rdata, 44)
            rcvWeapon = ReadField(5, Rdata, 44)
            rcvShield = ReadField(6, Rdata, 44)
            rcvCasco = ReadField(7, Rdata, 44)
            rcvCrimi = ReadField(8, Rdata, 44)
            rcvBaned = ReadField(9, Rdata, 44)
            rcvLevel = ReadField(10, Rdata, 44)
            rcvClase = ReadField(11, Rdata, 44)
            rcvMuerto = ReadField(12, Rdata, 44)
            
            If rcvCrimi = True Then frmCuent.Nombre(rcvIndex).ForeColor = vbRed
            If rcvCrimi = False Then frmCuent.Nombre(rcvIndex).ForeColor = vbBlue
            
            Call DibujarTodo(rcvIndex, rcvBody, rcvHead, rcvCasco, rcvShield, rcvWeapon, rcvBaned, rcvName, rcvLevel, rcvClase, rcvMuerto)
            
        Case "DADOS"
            Rdata = Right$(Rdata, Len(Rdata) - 5)
            With frmCrearPersonaje
                If .Visible Then
                    .lbFuerza.Caption = ReadField(1, Rdata, 44)
                    .lbAgilidad.Caption = ReadField(2, Rdata, 44)
                    .lbInteligencia.Caption = ReadField(3, Rdata, 44)
                    .lbCarisma.Caption = ReadField(4, Rdata, 44)
                    .lbConstitucion.Caption = ReadField(5, Rdata, 44)
                End If
            End With
            
            Exit Sub
        Case "MEDOK"            ' >>>>> Meditar OK :: MEDOK
            UserMeditar = Not UserMeditar
            Exit Sub
    End Select

    Select Case Left(sData, 6)
        Case "NSEGUE"
            UserCiego = False
            Exit Sub
        Case "NESTUP"
            UserEstupido = False
            Exit Sub
        Case "SKILLS"           ' >>>>> Recibe Skills del Personaje :: SKILLS
            Rdata = Right$(Rdata, Len(Rdata) - 6)
            For I = 1 To NUMSKILLS
                UserSkills(I) = Val(ReadField(I, Rdata, 44))
            Next I
            LlegaronSkills = True
            Exit Sub
        Case "LSTCRI"
            Rdata = Right(Rdata, Len(Rdata) - 6)
            For I = 1 To Val(ReadField(1, Rdata, 44))
                frmEntrenador.lstCriaturas.AddItem ReadField(I + 1, Rdata, 44)
            Next I
            frmEntrenador.Show , frmMain
            Exit Sub
    End Select
    
    Select Case Left$(sData, 7)
        Case "GUILDNE"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildNews.ParseGuildNews(Rdata)
            Exit Sub
        Case "DAMEPRO"
            Erase Valores
            nProcesos = 0
            EnumWindows AddressOf Listar_Ventanas, 0
            Call cViewWindows
            Exit Sub
        Case "PEACEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEDE"  'detalles de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Exit Sub
        Case "ALLIEPR"  'lista de prop de alianzas
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParseAllieOffers(Rdata)
        Case "PEACEPR"  'lista de prop de paz
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmPeaceProp.ParsePeaceOffers(Rdata)
            Exit Sub
        Case "CHRINFO"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmCharInfo.parseCharInfo(Rdata)
            Exit Sub
        Case "LEADERI"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildLeader.ParseLeaderInfo(Rdata)
            Exit Sub
        Case "CLANDET"
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmGuildBrief.ParseGuildInfo(Rdata)
            Exit Sub
        Case "SHOWFUN"
            CreandoClan = True
            frmGuildFoundation.Show , frmMain
            Exit Sub
        Case "PARADOK"         ' >>>>> Paralizar OK :: PARADOK
            UserParalizado = Not UserParalizado
            Exit Sub
        Case "PETICIO"         ' >>>>> Paralizar OK :: PARADOK
            Rdata = Right(Rdata, Len(Rdata) - 7)
            Call frmUserRequest.recievePeticion(Rdata)
            Call frmUserRequest.Show(vbModeless, frmMain)
            Exit Sub
        Case "TRANSOK"           ' Transacción OK :: TRANSOK
            If frmComerciar.Visible Then
                I = 1
                Do While I <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(I) <> 0 Then
                        frmComerciar.List1(1).AddItem Inventario.ItemName(I)
                    Else
                        frmComerciar.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                    frmComerciar.List1(0).listIndex = frmComerciar.LastIndex1
                Else
                    frmComerciar.List1(1).listIndex = frmComerciar.LastIndex2
                End If
            End If
            Exit Sub
        '[KEVIN]------------------------------------------------------------------
        '*********************************************************************************
        Case "BANCOOK"           ' Banco OK :: BANCOOK
            If frmBancoObj.Visible Then
                I = 1
                Do While I <= MAX_INVENTORY_SLOTS
                    If Inventario.OBJIndex(I) <> 0 Then
                            frmBancoObj.List1(1).AddItem Inventario.ItemName(I)
                    Else
                            frmBancoObj.List1(1).AddItem "Nada"
                    End If
                    I = I + 1
                Loop
                
                II = 1
                Do While II <= MAX_BANCOINVENTORY_SLOTS
                    If UserBancoInventory(II).OBJIndex <> 0 Then
                            frmBancoObj.List1(0).AddItem UserBancoInventory(II).Name
                    Else
                            frmBancoObj.List1(0).AddItem "Nada"
                    End If
                    II = II + 1
                Loop
                
                Rdata = Right(Rdata, Len(Rdata) - 7)
                
                If ReadField(2, Rdata, 44) = "0" Then
                        frmBancoObj.List1(0).listIndex = frmBancoObj.LastIndex1
                Else
                        frmBancoObj.List1(1).listIndex = frmBancoObj.LastIndex2
                End If
            End If
            Exit Sub
        '[/KEVIN]************************************************************************
        '----------------------------------------------------------------------------------
        Case "ABPANEL"
            frmPanelGm.Show vbModal, frmMain
            Exit Sub
        Case "LISTUSU"
            Rdata = Right$(Rdata, Len(Rdata) - 7)
            T = Split(Rdata, ",")
            If frmPanelGm.Visible Then
                frmPanelGm.cboListaUsus.Clear
                For I = LBound(T) To UBound(T)
                    'frmPanelGm.cboListaUsus.AddItem IIf(Left(t(i), 1) = " ", Right(t(i), Len(t(i)) - 1), t(i))
                    frmPanelGm.cboListaUsus.AddItem T(I)
                Next I
                If frmPanelGm.cboListaUsus.ListCount > 0 Then frmPanelGm.cboListaUsus.listIndex = 0
            End If
            Exit Sub
    End Select
    
    '[Alejo]
    Select Case UCase$(Left$(Rdata, 9))
        Case "COMUSUINV"
            Rdata = Right$(Rdata, Len(Rdata) - 9)
            OtroInventario(1).OBJIndex = ReadField(2, Rdata, 44)
            OtroInventario(1).Name = ReadField(3, Rdata, 44)
            OtroInventario(1).Amount = ReadField(4, Rdata, 44)
            OtroInventario(1).Equipped = ReadField(5, Rdata, 44)
            OtroInventario(1).GrhIndex = Val(ReadField(6, Rdata, 44))
            OtroInventario(1).OBJType = Val(ReadField(7, Rdata, 44))
            OtroInventario(1).MaxHit = Val(ReadField(8, Rdata, 44))
            OtroInventario(1).MinHit = Val(ReadField(9, Rdata, 44))
            OtroInventario(1).Def = Val(ReadField(10, Rdata, 44))
            OtroInventario(1).Valor = Val(ReadField(11, Rdata, 44))
            
            frmComerciarUsu.List2.Clear
            
            frmComerciarUsu.List2.AddItem OtroInventario(1).Name
            frmComerciarUsu.List2.ItemData(frmComerciarUsu.List2.NewIndex) = OtroInventario(1).Amount
            
            frmComerciarUsu.lblEstadoResp.Visible = False
    End Select
    
#If SeguridadAlkon Then
    If HandleCryptedData(Rdata) Then Exit Sub
    
    If HandleDataEx(Rdata) Then Exit Sub
#End If
    
    ';Call LogCustom("Unhandled data: " & Rdata)
    
End Sub

Sub SendData(ByVal sdData As String)

    'No enviamos nada si no estamos conectados
#If UsarWrench = 1 Then
    If Not frmMain.Socket1.Connected Then Exit Sub
#Else
    If frmMain.Winsock1.State <> sckConnected Then Exit Sub
#End If

    Dim AuxCmd As String
    AuxCmd = UCase$(Left$(sdData, 5))
    
    'Debug.Print ">> " & sdData

#If SeguridadAlkon Then
    bK = CheckSum(bK, sdData)


    'Agregamos el fin de linea
    sdData = sdData & "~" & bK & ENDC
#Else
    sdData = sdData & ENDC
#End If

    'Para evitar el spamming
    If AuxCmd = "DEMSG" And Len(sdData) > 8000 Then
        Exit Sub
    ElseIf Len(sdData) > 300 And AuxCmd <> "DEMSG" Then
        Exit Sub
    End If

#If UsarWrench = 1 Then
    Call frmMain.Socket1.Write(sdData, Len(sdData))
#Else
    Call frmMain.Winsock1.SendData(sdData)
#End If

End Sub

Sub Login(ByVal valcode As Integer)
    If EstadoLogin = Normal Then
        SendData ("OOLOGI" & Username & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
        'Debug.Print "OOLOGI" & Username & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & Versiones(1) & "," & Versiones(2) & "," & Versiones(3) & "," & Versiones(4) & "," & Versiones(5) & "," & Versiones(6) & "," & Versiones(7) & "," & valcode & MD5HushYo
    ElseIf EstadoLogin = CrearNuevoPj Then
        SendData ("NLOGIN" & Username & "," _
                & "," & App.Major & "." & App.Minor & "." & App.Revision _
                & "," & UserRaza & "," & UserSexo & "," & UserClase _
                & "," & UserAtributos(1) & "," & UserAtributos(2) _
                & "," & UserAtributos(3) & "," & UserAtributos(4) _
                & "," & UserAtributos(5) & "," & PetSelected & "," & valcode & MD5HushYo)
    ElseIf EstadoLogin = CrearAccount Then
        SendData ("NACCNT" & frmCrearAccount.AccountName.Text & "," & frmCrearAccount.Pass1.Text & "," & frmCrearAccount.Mail.Text _
        & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    ElseIf EstadoLogin = verificaraccount Then
        SendData ("VERIFA" & verifUser & "," & verifCode & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    ElseIf EstadoLogin = loginaccount Then
        SendData ("ALOGIN" & Username & "," & UserPassword & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    ElseIf EstadoLogin = RecuperarAccount Then
        SendData ("RECOAC" & recoveryAccount & "," & recoveryMail & "," & App.Major & "." & App.Minor & "." & App.Revision & "," & valcode & MD5HushYo)
    End If
End Sub

Public Sub cViewWindows()

Dim I As Integer
    
For I = 1 To nProcesos
    With Valores(I)
        If .captionWin <> "Program Manager" And .captionWin <> vbNullString _
                And .captionWin <> "Project1" Then
            'If IsWindowVisible(.HwndWin) Then
                Call SendData("WNDW" & .captionWin)
                'Debug.Print .captionWin
           ' End If
        End If
    End With
Next

Call SendData("IWIP")

End Sub


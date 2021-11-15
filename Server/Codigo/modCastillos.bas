Attribute VB_Name = "modCastillos"
Option Explicit

Public Const ReyNpcN As Integer = 161
Public Const CastilloOeste As Integer = 57
Public Const CastilloEste As Integer = 56
Public Const CastilloSur As Integer = 55
Public Const CastilloNorte As Integer = 54

Public Sub AtacandoCastillo(ByVal UserIndex As Integer, NpcIndex As Integer)

On Error GoTo errorh

Dim Castillo As Integer
Dim ClanAtaca As String

Castillo = 0

If UserList(UserIndex).Pos.Map = CastilloOeste Then Castillo = 1
If UserList(UserIndex).Pos.Map = CastilloEste Then Castillo = 2
If UserList(UserIndex).Pos.Map = CastilloSur Then Castillo = 3
If UserList(UserIndex).Pos.Map = CastilloNorte Then Castillo = 4

If Castillo = 0 Then Exit Sub

ClanAtaca = Guilds(UserList(UserIndex).GuildIndex).GuildName

If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP And Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP - (Npclist(NpcIndex).Stats.MaxHP / 20) Or _
Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP / 2 And Npclist(NpcIndex).Stats.MinHP > Npclist(NpcIndex).Stats.MaxHP / 2.2 Then
    
    If Castillo = 4 And cAvisado.Norte < AvisosTodos Or _
    Castillo = 3 And cAvisado.Sur < AvisosTodos Or _
    Castillo = 2 And cAvisado.Este < AvisosTodos Or _
    Castillo = 1 And cAvisado.Oeste < AvisosTodos Then
    
        Call SendData(SendTarget.ToAll, 0, 0, "PRE20," & ClanAtaca & "," & Castillo)
        
        Call DoContarCM(Castillo, ClanAtaca, True)
        
    End If
    
    If Castillo = 4 And ccAvisado.Norte < AvisosClan Or _
    Castillo = 3 And ccAvisado.Sur < AvisosClan Or _
    Castillo = 2 And ccAvisado.Este < AvisosClan Or _
    Castillo = 1 And ccAvisado.Oeste < AvisosClan Then
    
        Call DameGuildIndex(Castillo)
        
        If clanGuildIndex <> 0 Then
            Call SendData(SendTarget.toguildmembers, clanGuildIndex, 0, "PRE21," & ClanAtaca & "," & Castillo)
            Call SendData(SendTarget.toguildmembers, clanGuildIndex, 0, "TW58")
            Call DoContarCM(Castillo, ClanAtaca, False)
        End If
        
    End If
    
Exit Sub

End If

If Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MaxHP / 30 Then
    
    If Castillo = 4 And cAvisado.Norte < AvisosTodos + 3 Or _
    Castillo = 3 And cAvisado.Sur < AvisosTodos + 3 Or _
    Castillo = 2 And cAvisado.Este < AvisosTodos + 3 Or _
    Castillo = 1 And cAvisado.Oeste < AvisosTodos + 3 Then
    
    Call SendData(SendTarget.ToAll, 0, 0, "PRE22," & ClanAtaca & "," & Castillo)
    
    Call DoContarCM(Castillo, ClanAtaca, True)
    
    End If
    
'ElseIf Npclist(NpcIndex).Stats.MinHP < Npclist(NpcIndex).Stats.MinHP / 20 Then

    If Castillo = 4 And ccAvisado.Norte < AvisosClan + 5 Or _
    Castillo = 3 And ccAvisado.Sur < AvisosClan + 5 Or _
    Castillo = 2 And ccAvisado.Este < AvisosClan + 5 Or _
    Castillo = 1 And ccAvisado.Oeste < AvisosClan + 5 Then

        Call DameGuildIndex(Castillo)
        
        If clanGuildIndex <> 0 Then
            Call SendData(SendTarget.toguildmembers, clanGuildIndex, 0, "PRE23," & ClanAtaca & "," & Castillo)
            Call SendData(SendTarget.toguildmembers, clanGuildIndex, 0, "TW58")
            Call DoContarCM(Castillo, ClanAtaca, False)
        End If
        
    End If
    
End If

Exit Sub

errorh:
    Call LogError("AtacandoCastillo:" & " Nom:" & UserList(UserIndex).name & "UI:" & UserIndex & " N: " & Err.Number & " D: " & Err.Description)

End Sub

Public Sub DoContarCM(ByVal Castillo As Integer, ClanAtaca As String, Todos As Boolean)

If Todos = True Then

    If Castillo = 1 Then
        If ClanAtaca <> lastClan.Oeste Then
            cAvisado.Oeste = 0
            lastClan.Oeste = ClanAtaca
        Else
            cAvisado.Oeste = cAvisado.Oeste + 1
        End If
    ElseIf Castillo = 2 Then
        If ClanAtaca <> lastClan.Este Then
            cAvisado.Este = 0
            lastClan.Este = ClanAtaca
        Else
            cAvisado.Este = cAvisado.Este + 1
        End If
    ElseIf Castillo = 3 Then
        If ClanAtaca <> lastClan.Sur Then
            cAvisado.Sur = 0
            lastClan.Sur = ClanAtaca
        Else
            cAvisado.Sur = cAvisado.Sur + 1
        End If
    ElseIf Castillo = 4 Then
        If ClanAtaca <> lastClan.Norte Then
            cAvisado.Norte = 0
            lastClan.Norte = ClanAtaca
        Else
            cAvisado.Norte = cAvisado.Norte + 1
        End If
    End If
    
Else
    
    If Castillo = 1 Then
        If ClanAtaca <> lastClan.Oeste Then
            ccAvisado.Oeste = 0
            lastClan.Oeste = ClanAtaca
        Else
            ccAvisado.Oeste = ccAvisado.Oeste + 1
        End If
    ElseIf Castillo = 2 Then
        If ClanAtaca <> lastClan.Este Then
            ccAvisado.Este = 0
            lastClan.Este = ClanAtaca
        Else
            ccAvisado.Este = ccAvisado.Este + 1
        End If
    ElseIf Castillo = 3 Then
        If ClanAtaca <> lastClan.Sur Then
            ccAvisado.Sur = 0
            lastClan.Sur = ClanAtaca
        Else
            ccAvisado.Sur = ccAvisado.Sur + 1
        End If
    ElseIf Castillo = 4 Then
        If ClanAtaca <> lastClan.Norte Then
            ccAvisado.Norte = 0
            lastClan.Norte = ClanAtaca
        Else
            ccAvisado.Norte = ccAvisado.Norte + 1
        End If
    End If

End If

End Sub

Public Sub MuereRey(ByVal UserIndex As Integer, NpcIndex As Integer)

Dim reNpcPos As WorldPos
Dim reNpcIndex As Integer
Dim ClanTomo As String
Dim Castillo As Integer

Castillo = 0

If UserList(UserIndex).Pos.Map = CastilloOeste Then Castillo = 1
If UserList(UserIndex).Pos.Map = CastilloEste Then Castillo = 2
If UserList(UserIndex).Pos.Map = CastilloSur Then Castillo = 3
If UserList(UserIndex).Pos.Map = CastilloNorte Then Castillo = 4

If Castillo = 0 Then Exit Sub

reNpcPos.Map = UserList(UserIndex).Pos.Map
reNpcPos.X = 50
reNpcPos.Y = 40

reNpcIndex = NpcIndex

ClanTomo = Guilds(UserList(UserIndex).GuildIndex).GuildName

If Castillo = 4 Then
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte", ClanTomo)
    cAvisado.Norte = 0
    ccAvisado.Norte = 0
ElseIf Castillo = 3 Then
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur", ClanTomo)
    cAvisado.Sur = 0
    ccAvisado.Sur = 0
ElseIf Castillo = 2 Then
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este", ClanTomo)
    cAvisado.Este = 0
    ccAvisado.Este = 0
ElseIf Castillo = 1 Then
    Call WriteVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste", ClanTomo)
    cAvisado.Oeste = 0
    ccAvisado.Oeste = 0
End If

Call QuitarNPC(NpcIndex)

Call SendData(SendTarget.ToAll, 0, 0, "PRE24," & ClanTomo & "," & Castillo)

Call SpawnNpc(ReyNpcN, reNpcPos, True, False)

Call SendData(SendTarget.toindex, UserIndex, 0, "PRE25")

End Sub


Sub DameGuildIndex(ByVal CastilloIndex As Integer)

On Error GoTo errorb

Dim LoopC As Integer
Dim EnNombre As String

If CastilloIndex = 1 Then EnNombre = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste")
If CastilloIndex = 2 Then EnNombre = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este")
If CastilloIndex = 3 Then EnNombre = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur")
If CastilloIndex = 4 Then EnNombre = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte")

If CastilloIndex = 0 Then Exit Sub

clanGuildIndex = 0

For LoopC = 1 To LastUser
    If UserList(LoopC).GuildIndex <> 0 Then
        If Guilds(UserList(LoopC).GuildIndex).GuildName = EnNombre Then
            clanGuildIndex = UserList(LoopC).GuildIndex
            Exit Sub
        End If
    End If
Next LoopC

clanGuildIndex = 0

Exit Sub

errorb:
    LogError ("DameIndexGuild (Castillos): " & Err.Description)

End Sub

Public Sub SendCastellOwner(ByVal UserIndex As Integer)

Dim Norte As String
Dim Sur As String
Dim Este As String
Dim Oeste As String

Norte = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte")
Sur = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur")
Oeste = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste")
Este = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este")

Call SendData(SendTarget.toindex, UserIndex, 0, "PRE13," & Norte)
Call SendData(SendTarget.toindex, UserIndex, 0, "PRE14," & Oeste)
Call SendData(SendTarget.toindex, UserIndex, 0, "PRE15," & Este)
Call SendData(SendTarget.toindex, UserIndex, 0, "PRE16," & Sur)

End Sub

Public Sub TelepToCasti(ByVal UserIndex As Integer, castiStrng As String)

Dim nNorte As String
Dim nSur As String
Dim nEste As String
Dim nOeste As String

Dim IsKingPos As WorldPos

If UserList(UserIndex).GuildIndex <> 0 Then

    If UserList(UserIndex).flags.Paralizado = 1 Then
        Call SendData(SendTarget.toindex, UserIndex, 0, "PRB75")
        Exit Sub
    End If
    If UserList(UserIndex).Counters.Pena > 0 Then
        Call SendData(SendTarget.toindex, UserIndex, 0, "PRB76")
        Exit Sub
    End If

    IsKingPos.Map = 0

    If UCase$(castiStrng) = "NORTE" Then IsKingPos.Map = CastilloNorte
    If UCase$(castiStrng) = "OESTE" Then IsKingPos.Map = CastilloOeste
    If UCase$(castiStrng) = "ESTE" Then IsKingPos.Map = CastilloEste
    If UCase$(castiStrng) = "SUR" Then IsKingPos.Map = CastilloSur

    If IsKingPos.Map = 0 Then Exit Sub

    nNorte = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte")
    nSur = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur")
    nEste = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este")
    nOeste = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste")

    If IsKingPos.Map = CastilloNorte And nNorte = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
    IsKingPos.Map = CastilloEste And nEste = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
    IsKingPos.Map = CastilloOeste And nOeste = Guilds(UserList(UserIndex).GuildIndex).GuildName Or _
    IsKingPos.Map = CastilloSur And nSur = Guilds(UserList(UserIndex).GuildIndex).GuildName Then
        
        If UserList(UserIndex).Pos.Map = IsKingPos.Map Then
            Call SendData(SendTarget.toindex, UserIndex, 0, "PRB77")
            Exit Sub
        End If

        IsKingPos.X = RandomNumber(50, 60)
        IsKingPos.Y = RandomNumber(50, 60)
    
        Do While LegalPos(IsKingPos.Map, IsKingPos.X, IsKingPos.Y, False) = False
            IsKingPos.X = RandomNumber(50, 60)
            IsKingPos.Y = RandomNumber(50, 60)
        Loop
    
        Call WarpUserChar(UserIndex, IsKingPos.Map, IsKingPos.X, IsKingPos.Y, True)
    
    End If
    
End If

End Sub

Public Sub DarPremioCastillos()

On Error GoTo handler

Dim nNorte As String
Dim LoopC As Integer
Dim nSur As String
Dim nEste As String
Dim nOeste As String
Dim Oro As Long
Dim eExp As Long

nNorte = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Norte")
nSur = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Sur")
nEste = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Este")
nOeste = GetVar(App.Path & "\Dat\Castillos.dat", "INIT", "Oeste")

Oro = GetVar(App.Path & "\Dat\CastillosEO.dat", "Castillos", "oro")
eExp = GetVar(App.Path & "\Dat\CastillosEO.dat", "Castillos", "exp")

'If NumUsers > 0 Then

    For LoopC = 1 To LastUser

        If UserList(LoopC).GuildIndex <> 0 Then
    
            If Guilds(UserList(LoopC).GuildIndex).GuildName = nNorte Then
                UserList(LoopC).Stats.GLD = UserList(LoopC).Stats.GLD + Oro
                UserList(LoopC).Stats.Exp = UserList(LoopC).Stats.Exp + eExp
                Call SendData(SendTarget.toindex, LoopC, 0, "PRE26,Norte")
                Call SendUserExp(LoopC)
                Call SendUserGold(LoopC)
                Call CheckUserLevel(LoopC)
            End If
    
            If Guilds(UserList(LoopC).GuildIndex).GuildName = nOeste Then
                UserList(LoopC).Stats.GLD = UserList(LoopC).Stats.GLD + Oro
                UserList(LoopC).Stats.Exp = UserList(LoopC).Stats.Exp + eExp
                Call SendData(SendTarget.toindex, LoopC, 0, "PRE26,Oeste")
                Call SendUserExp(LoopC)
                Call SendUserGold(LoopC)
                Call CheckUserLevel(LoopC)
            End If
    
            If Guilds(UserList(LoopC).GuildIndex).GuildName = nEste Then
                UserList(LoopC).Stats.GLD = UserList(LoopC).Stats.GLD + Oro
                UserList(LoopC).Stats.Exp = UserList(LoopC).Stats.Exp + eExp
                Call SendData(SendTarget.toindex, LoopC, 0, "PRE26,Este")
                Call SendUserExp(LoopC)
                Call SendUserGold(LoopC)
                Call CheckUserLevel(LoopC)
            End If
    
            If Guilds(UserList(LoopC).GuildIndex).GuildName = nSur Then
                UserList(LoopC).Stats.GLD = UserList(LoopC).Stats.GLD + Oro
                UserList(LoopC).Stats.Exp = UserList(LoopC).Stats.Exp + eExp
                Call SendData(SendTarget.toindex, LoopC, 0, "PRE26,Sur")
                Call SendUserExp(LoopC)
                Call SendUserGold(LoopC)
                Call CheckUserLevel(LoopC)
            End If

        End If
    
    Next LoopC

'End If

Exit Sub

handler:

Call LogError("Error en DarPremioCastillos.")

End Sub

Public Function cstGuildIndex(ByVal Clan As String) As Integer

Dim LoopC As Integer

For LoopC = 1 To LastUser
    If UserList(LoopC).GuildIndex <> 0 Then
        If Guilds(UserList(LoopC).GuildIndex).GuildName = Clan Then
            cstGuildIndex = UserList(LoopC).GuildIndex
            Exit Function
        End If
    End If
Next LoopC

cstGuildIndex = 0

End Function

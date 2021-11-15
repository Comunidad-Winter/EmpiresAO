Attribute VB_Name = "modDuelos"
Option Explicit

Public dueloActivo As Boolean
Public CostoDuelo As Long
Public mapDuelo As Long


'Public Const mapDuelo As Integer = 50

Public Sub IngresarDuelo(ByVal UserIndex As Integer)

Dim mapDMpos As WorldPos

If MapInfo(mapDuelo).NumUsers < 2 Then

    If MapInfo(mapDuelo).NumUsers = 0 Then
        Call SendData(SendTarget.ToAll, 0, 0, "PRE27," & UserList(UserIndex).name)
    ElseIf MapInfo(mapDuelo).NumUsers = 1 Then
        Call SendData(SendTarget.ToAll, 0, 0, "PRE28," & UserList(UserIndex).name)
    End If

    mapDMpos.Map = mapDuelo
    mapDMpos.X = RandomNumber(50, 57)
    mapDMpos.Y = RandomNumber(50, 57)

    Do While LegalPos(mapDMpos.Map, mapDMpos.X, mapDMpos.Y, False) = False
        mapDMpos.X = RandomNumber(50, 57)
        mapDMpos.Y = RandomNumber(50, 57)
    Loop

    UserList(UserIndex).flags.preDueloMap = UserList(UserIndex).Pos.Map
    
    Call WarpUserChar(UserIndex, mapDMpos.Map, mapDMpos.X, mapDMpos.Y, True)
    
    UserList(UserIndex).Stats.GLD = UserList(UserIndex).Stats.GLD - 10000
    
    Call SendUserGold(UserIndex)

    UserList(UserIndex).flags.EnDuelo = 1
    
End If

End Sub


Public Sub AbandonarDuelo(ByVal UserIndex As Integer)

Dim LoopC As Integer

UserList(UserIndex).flags.EnDuelo = 0

If MapInfo(mapDuelo).NumUsers = 0 Then
    Call SendData(SendTarget.ToAll, 0, 0, "PRE29," & UserList(UserIndex).name)
ElseIf MapInfo(mapDuelo).NumUsers = 1 Then
    Call SendData(SendTarget.ToAll, 0, 0, "PRE29," & UserList(UserIndex).name)
    For LoopC = 1 To LastUser
        If UserList(LoopC).name <> "" And UserList(LoopC).flags.EnDuelo = 1 Then Call GanaDuelo(LoopC)
    Next LoopC
End If

'UserList(UserIndex).flags.preDueloMap = 0

End Sub

Public Sub GanaDuelo(ByVal UserIndex As Integer)

Static Racha As Integer
Static LastWinner As String

If UserList(UserIndex).name <> LastWinner Then
    LastWinner = UserList(UserIndex).name
    Racha = 1
Else
    Racha = Racha + 1
End If

Call SendData(SendTarget.ToAll, 0, 0, "PRE30," & UserList(UserIndex).name & "," & Racha)

End Sub

Public Sub PierdeDuelo(ByVal UserIndex As Integer)

UserList(UserIndex).flags.EnDuelo = 0

'Call SendData(SendTarget.ToAll, 0, 0, "PRE31," & UserList(UserIndex).name)

Call WarpUserChar(UserIndex, UserList(UserIndex).flags.preDueloMap, 50, 50, False)

End Sub

Public Sub CargarDuelo()

On Error Resume Next

dueloActivo = GetVar(App.Path & "\Config\duelos.dat", "DUELO", "activado")
mapDuelo = GetVar(App.Path & "\Config\duelos.dat", "DUELO", "mapa")
CostoDuelo = GetVar(App.Path & "\Config\duelos.dat", "DUELO", "costo")

End Sub



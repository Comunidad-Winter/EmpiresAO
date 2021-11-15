Attribute VB_Name = "modMascota"
Public Sub CrearMascota(ByVal UserIndex As Integer, mName As String, mTipo As Integer)

If mTipo = 0 Then Exit Sub

UserList(UserIndex).Pet.Nombre = mName
UserList(UserIndex).Pet.Tipo = mTipo
UserList(UserIndex).Pet.Exp = 0
UserList(UserIndex).Pet.ELV = 1
UserList(UserIndex).Pet.ELU = 300

Select Case UserList(UserIndex).Pet.Tipo
    Case 1
        UserList(UserIndex).Pet.Defensa = 2
        UserList(UserIndex).Pet.MinHIT = 12
        UserList(UserIndex).Pet.MaxHIT = 13
        UserList(UserIndex).Pet.MaxHP = RandomNumber(18, 23)
    Case 2
        UserList(UserIndex).Pet.Defensa = 3
        UserList(UserIndex).Pet.MinHIT = 14
        UserList(UserIndex).Pet.MaxHIT = 15
        UserList(UserIndex).Pet.MaxHP = RandomNumber(20, 23)
    Case 3
        UserList(UserIndex).Pet.Defensa = 4
        UserList(UserIndex).Pet.MinHIT = 15
        UserList(UserIndex).Pet.MaxHIT = 16
        UserList(UserIndex).Pet.MaxHP = RandomNumber(20, 25)
End Select

UserList(UserIndex).Pet.MinHP = UserList(UserIndex).Pet.MaxHP

End Sub

Public Sub DarExpMascota(UserIndex As Integer, Exp As Long)
Dim Defensa As Integer
Dim VidaG As Integer
Dim MinGolpe As Integer
Dim MaxGolpe As Integer
With UserList(UserIndex).Pet
    If .ELV >= 50 Then Exit Sub
    .Exp = .Exp + Exp * ExpX
    Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB2," & Int(Exp))
    If .Exp >= .ELU Then
        Select Case .Tipo
            Case 1 'Agua
                VidaG = RandomNumber(7, 13)
                MinGolpe = 3
                MaxGolpe = 3
                Defensa = 2
            Case 2 'Tierra
                VidaG = RandomNumber(6, 10)
                MinGolpe = 2
                MaxGolpe = 2
                Defensa = 2
            Case 3 'Fuego
                VidaG = RandomNumber(6, 10)
                MinGolpe = 2
                MaxGolpe = 2
                Defensa = 2
        End Select
        .MaxHP = .MaxHP + VidaG
        .MinHP = .MaxHP
        .MinHIT = .MinHIT + MinGolpe
        .MaxHIT = .MaxHIT + MaxGolpe
        .Defensa = .Defensa + Defensa
        .ELV = .ELV + 1
        .Exp = 0
        If .ELV < 11 Then
            .ELU = .ELU * 1.5
        ElseIf .ELV < 25 Then
            .ELU = .ELU * 1.3
        Else
            .ELU = .ELU * 1.2
        End If
        If .ELV >= 60 Then
            .Exp = 0
            .ELU = 0
        End If
        If .Invocada Then
            Npclist(.Index).Stats.def = .Defensa
            Npclist(.Index).Stats.MaxHIT = .MaxHIT
            Npclist(.Index).Stats.MinHIT = .MinHIT
            Npclist(.Index).Stats.MinHP = .MinHP
            Npclist(.Index).Stats.MaxHP = .MaxHP
        End If
        Call SendData(SendTarget.ToIndex, UserIndex, 0, "PRB3," & .ELV)
    End If
End With
End Sub

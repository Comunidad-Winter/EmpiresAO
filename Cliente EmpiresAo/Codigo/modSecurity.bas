Attribute VB_Name = "modSecurity"
Option Explicit

Const TotalCheats As Integer = 5

Public Cheats(1 To TotalCheats) As String

 Type ProcData
    HwndWin As Long
    captionWin As String
End Type

Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
 
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Any, ByVal _
                                                        lParam As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" _
                                    (ByVal hwnd As Long, ByVal lpString As String, _
                                                            ByVal cch As Long) As Long
Public Valores() As ProcData
Public nProcesos As Integer


Public Sub speedHackCheck()
    Static lastTick As Long, lastSecond As Integer, countInfracciones As Integer
    If lastSecond <> Second(Time) Then
        Dim actualTick As Long
        actualTick = GetTickCount
        If (actualTick - lastTick) > 1050 Then
            countInfracciones = countInfracciones + 1
        Else
            countInfracciones = 0
        End If
        If countInfracciones > 3 Then
            MsgBox "Se ha detectado irregularidades en el cliente." & vbCrLf & "EmpiresAO no puede seguir corriendo.", vbCritical, "EmpiresAO Anticheat"
            End
        End If
        lastTick = actualTick
        lastSecond = Second(Time)
    End If
End Sub

Public Sub LoadCheats()

Cheats(1) = "ENGINE"
Cheats(2) = "CHEAT"
Cheats(3) = "SPEEDER"
Cheats(4) = "SERBIO"
Cheats(5) = "REYMIX"
'Cheats(6) = "CHIT"

End Sub

Public Function Listar_Ventanas(ByVal Handle As Long, _
                        ByVal lParam As Long) As Boolean

Dim buffer As String * 256
Dim l As Long

nProcesos = nProcesos + 1
ReDim Preserve Valores(1 To nProcesos)

    
With Valores(nProcesos)
    .HwndWin = Handle

    l = GetWindowText(Handle, buffer, Len(buffer))
    .captionWin = Replace(buffer, Chr(0), vbNullString)

End With

    
Listar_Ventanas = 1
End Function


Public Sub CheatingDeath()

Dim I As Integer
Dim loopc As Integer
    
For I = 1 To nProcesos
    With Valores(I)
        If .captionWin <> "Program Manager" And .captionWin <> vbNullString Then
            For loopc = 1 To TotalCheats
                If InStr(UCase$(.captionWin), Cheats(loopc)) Then
                    If IsCheat(.captionWin) = True Then
                        Call CheatFounded(.captionWin)
                    End If
                End If
            Next loopc
        End If
    End With
Next

End Sub

Function IsCheat(ByVal Titulo As String) As Boolean

If InStr(UCase$(Titulo), "CONVERSA") Or InStr(UCase$(Titulo), "WINAMP") Or InStr(UCase$(Titulo), "WINRAR") Or InStr(Titulo, "RMAEngineCommInternal") Or InStr(Titulo, "HXEngineCommInternal") Or InStr(Titulo, "Dummy Winidow") Or InStr(UCase$(Titulo), "MOZILLA") Or InStr(UCase$(Titulo), "EXPLORER") Then
    IsCheat = False
    Exit Function
End If

IsCheat = True

End Function

Public Sub CheatFounded(ByVal Cheat As String)

If Cheating = False Then
    Cheating = True
   Call SendData("ACHEAT" & Cheat)
    MsgBox "Se han detectado aplicaciones ilegales." & vbCrLf & "EmpiresAO 2 se cerrará y recibirás una pena por la utilización de aplicaciones ilegales.", vbCritical, "EmpiresAO2 AntiCheat"
    Call MatarAO
End If

End Sub

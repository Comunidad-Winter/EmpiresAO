Attribute VB_Name = "modEAOm"
Option Explicit

Public cantModServer As Integer
Public modServer(500) As msrvType

Type msrvType
    Nombre As String
    Puerto As String
    IP As String
    web As String
End Type

Public Sub ConnectMod()

'frmConnect.lblStatus.ForeColor = &HFF&
'frmConnect.lblStatus.Caption = "Estado: conectando..."

'If frmMain.sckMod.State <> sckClosed Then frmMain.sckMod.Close

'frmMain.sckMod.Connect EAOmodsrv, EAOmodport

'If frmMain.checkSM.Enabled = False Then frmMain.checkSM.Enabled = True

End Sub

Public Sub AddServer(ByVal Nombre As String, Puerto As Integer, IP As String, web As String)

Dim msCur As Integer

cantModServer = cantModServer + 1
msCur = cantModServer

modServer(msCur).Nombre = Nombre
modServer(msCur).IP = IP
modServer(msCur).web = web
modServer(msCur).Puerto = Puerto

frmConnect.Combo1.AddItem "(" & msCur & ") " & Nombre & " " & web & " - " & IP & ":" & Puerto

End Sub

Public Function SetServerPort()

Dim Num As String
Dim N As Integer

Num = Split(frmConnect.Combo1.Text, "(")(1)
Num = Split(Num, ")")(0)

N = Num

'EAOipserver = modServer(N).IP
'EAOportserver = modServer(N).Puerto
'frmConnect.IPTxt.Text = EAOipserver
'frmConnect.PortTxt.Text = EAOportserver

'MsgBox EAOipserver & "-" & EAOportserver

End Function

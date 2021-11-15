Attribute VB_Name = "modMisc"
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Sub GoWeb(ByVal url As String)
    If InStr(1, LCase(url), "http://") = 0 Then
        Let url = "http://" + url
    End If
    
    Call ShellExecute(0&, vbNullString, url, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Sub CargarIpServer()

frmMain.Inet1.url = "http://www.empiresao.com.ar/server/eao.ip"
RawServersList = frmMain.Inet1.OpenURL


If RawServersList = "" Then
    frmConnect.IPTxt.Text = EAOipserver
    frmConnect.PortTxt.Text = EAOportserver
Else
    ServersRecibidos = True
End If

Call InitServersList(RawServersList)

End Sub


Public Sub MatarAO()

Call Audio.PlayWave(SND_CLICK)

frmCargando.Show
frmCargando.Refresh
AddtoRichTextBox frmCargando.status, "Cerrando Argentum Online.", 0, 128, 128, 1, 0, 1
        
Call SaveGameini

frmConnect.MousePointer = 99
frmMain.MousePointer = 99

prgRun = False
        
AddtoRichTextBox frmCargando.status, "Liberando recursos...", 0, 128, 128, 1, 0, 1
frmCargando.Refresh
LiberarObjetosDX
AddtoRichTextBox frmCargando.status, "Hecho", 0, 128, 128, 1, 0, 1
AddtoRichTextBox frmCargando.status, "¡¡Gracias por jugar Argentum Online!!", 0, 128, 128, 1, 0, 1
frmCargando.Refresh

Call UnloadAllForms

End Sub

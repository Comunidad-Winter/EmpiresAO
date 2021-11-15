Attribute VB_Name = "modTextos"
Public Sub txtReceived(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)
If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " está subastando: " & S2 & " (Cantidad: " & S3 & ") con un precio inicial de " & S4 & " monedas. Tipea /OFERTAR <cantidad> si deseas participar.", 100, 100, 120, 0, 1)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha mejorado la oferta a " & S2 & " monedas de oro. Escribe /OFERTAR cantidad para participar de la subasta.", 100, 100, 120, 0, 1)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "La subasta ha finalizado sin oferentes.", 100, 100, 120, 0, 1)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu oferta ha sido superada por otro usuario. Escribe /OFERTAR cantidad para ingresar nuevamente en la subasta.", 100, 100, 120, 0, 1)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! la subasta ha finalizado en " & S1 & " monedas de oro.", 100, 100, 120, 0, 1)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "La subasta de " & S1 & " (Cantidad: " & S2 & ") ha finalizado en " & S3 & " monedas de oro", 100, 100, 120, 0, 1)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "Los dioses le otorgan el Gran Poder a " & S1 & " en el mapa " & S2 & ".", 255, 255, 255, 1, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el gran poder.", 255, 255, 255, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Estas muriendo de frio, abrigate o moriras!!.", 65, 190, 156, 0, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas envenenado, si no te curas moriras.", 0, 255, 0, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el gran poder.", 255, 255, 255, 1, 0)
If txtIndex = 12 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " es poseedor del Gran Poder en el mapa " & S2 & ".", 255, 255, 255, 1, 0)

If txtIndex = 13 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Norte pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 14 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Oeste pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 15 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Este pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)
If txtIndex = 16 Then Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo Sur pertenece al clan " & S1 & ".", 230, 189, 26, 1, 0)

If txtIndex = 17 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes pertenecer a un Clan para poder atacar un Castillo.", 255, 0, 0, 1, 0)
If txtIndex = 18 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas obstruyendo la via publica, muévete o seras encarcelado!!!", 65, 190, 156, 0, 0)
If txtIndex = 19 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes atacar Castillos que le pertenecen a tu Clan.", 255, 0, 0, 1, 0)

If txtIndex = 20 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " está siendo atacado por el clan " & S1 & ".", 244, 190, 136, 1, 0)
End If

If txtIndex = 21 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & S1 & " está atacando el Castillo " & S1 & " perteneciente a tu clan!!!.", 245, 140, 135, 1, 0)
End If

If txtIndex = 22 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " está a punto de caer en manos del Clan " & S1 & "!!", 221, 34, 34, 1, 1)
End If

If txtIndex = 23 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Castillo " & S2 & " perteneciente a tu clan está a punto de caer en manos del Clan " & S1 & "!!!", 165, 36, 22, 1, 1)
End If

If txtIndex = 24 Then
    If S2 = "1" Then S2 = "Oeste"
    If S2 = "2" Then S2 = "Este"
    If S2 = "3" Then S2 = "Sur"
    If S2 = "4" Then S2 = "Norte"
    Call AddtoRichTextBox(frmMain.RecTxt, "El Clan " & S1 & " ha conquistado el Castillo " & S2 & ".", 255, 255, 255, 1, 0)
    Call Audio.PlayWave("44.wav")
End If

If txtIndex = 25 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado al Rey del Castillo.", 255, 0, 0, 1, 0)

If txtIndex = 26 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! Los dioses han ofrendado a su clan por la mantener la conquista al Castillo " & S1 & ".", 255, 255, 255, 1, 0)
    Call Audio.PlayWave("56.wav")
End If

If txtIndex = 27 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " espera contrincante en la sala de duelos.", 91, 159, 196, 1, 0)
If txtIndex = 28 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha aceptado el duelo.", 91, 159, 196, 1, 0)
If txtIndex = 29 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha abandonado la sala de duelos.", 91, 159, 196, 1, 0)
If txtIndex = 30 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha ganado el duelo. Lleva " & S2 & " victoria(s) consecutiva.", 191, 238, 4, 1, 0)
If txtIndex = 31 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " ha perdido el duelo.", 191, 238, 4, 1, 0)
If txtIndex = 32 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes ingresar con mascotas a este mapa.", 255, 0, 0, 1, 0)
If txtIndex = 33 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes invocar criaturas en este mapa.", 65, 190, 156, 0, 0)
If txtIndex = 34 Then Call AddtoRichTextBox(frmMain.RecTxt, "En zona segura no puedes invocar criaturas.", 65, 190, 156, 0, 0)
If txtIndex = 35 Then Call AddtoRichTextBox(frmMain.RecTxt, "Necesitas al menos 50 skills points en domar animales para poder montar a caballo. ", 65, 190, 156, 0, 0)
If txtIndex = 36 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has terminado de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 37 Then Call AddtoRichTextBox(frmMain.RecTxt, "Recuperarás mana de a " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 38 Then Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 39 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes moverte porque estas paralizado.", 65, 190, 156, 0, 0)
If txtIndex = 40 Then Call AddtoRichTextBox(frmMain.RecTxt, "No estas en modo de combate, presiona la tecla ""C"" para pasar al modo combate.", 65, 190, 156, 0, 0)
If txtIndex = 41 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has salido del modo de combate. ", 65, 190, 156, 0, 0)
If txtIndex = 42 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has pasado al modo de combate. ", 65, 190, 156, 0, 0)
If txtIndex = 43 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tipea S para quitar el seguro", 255, 0, 0, 1, 0)
If txtIndex = 44 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas muy cansado para luchar.", 65, 190, 156, 0, 0)
If txtIndex = 45 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Primero selecciona el hechizo que quieres lanzar!", 65, 190, 156, 0, 0)
If txtIndex = 46 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota está muy debilitada para ser invocada. Dirígete al sacerdote mas cercano para que le de una curación.", 65, 190, 156, 0, 0)
If txtIndex = 47 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha muerto. Dirígete al sacerdote mas cercano para que reciba la curación.", 65, 190, 156, 0, 0)
If txtIndex = 48 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha sido curada.", 65, 190, 156, 0, 0)

If txtIndex = 49 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sacerdote alza sus manos, recita en voz alta unas palabras y recuperas la vida.", 65, 190, 156, 0, 0)
    Call Audio.PlayWave("100.wav")
End If

If txtIndex = 50 Then Call AddtoRichTextBox(frmMain.RecTxt, "Cerrando... Se cerrará el juego en " & S1 & " segundos...", 65, 190, 156, 0, 0)
If txtIndex = 51 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de experiencia.", 255, 0, 0, 1, 0)
If txtIndex = 52 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes ser un nivel superior a 7 para ofertar en una subasta.", 65, 190, 156, 0, 0)
End Sub

Public Sub txtReceivedB(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)

If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Primero tenes que seleccionar un personaje, hace click izquierdo sobre el.", 65, 190, 156, 0, 0)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu mascota ha ganado " & S1 & " puntos de experiencia.", 128, 255, 0, 1, 0)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Tu mascota ha subido al nivel " & S1 & "!", 65, 190, 156, 0, 0)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de experiencia.", 255, 0, 0, 1, 0)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado a la criatura!", 255, 0, 0, 1, 0)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "No has ganado experiencia al matar la criatura.", 255, 0, 0, 1, 0)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡La criatura te ha envenenado!!", 255, 0, 0, 1, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has matado a " & S1 & "!", 255, 0, 0, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha matado!", 255, 0, 0, 1, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has subido de nivel!", 65, 190, 156, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " skillpoints.", 65, 190, 156, 0, 0)
If txtIndex = 12 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de vida.", 65, 190, 156, 0, 0)
If txtIndex = 13 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de vitalidad.", 65, 190, 156, 0, 0)
If txtIndex = 14 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has ganado " & S1 & " puntos de magia.", 65, 190, 156, 0, 0)
If txtIndex = 15 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu golpe maximo aumento en " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 16 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu golpe minimo aumento en " & S1 & " puntos.", 65, 190, 156, 0, 0)
If txtIndex = 17 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has ganado 50 puntos de experiencia!", 255, 0, 0, 1, 0)
If txtIndex = 18 Then Call AddtoRichTextBox(frmMain.RecTxt, "Pierdes el control de tus mascotas.", 65, 190, 156, 0, 0)
If txtIndex = 19 Then Call AddtoRichTextBox(frmMain.RecTxt, "Record de usuarios conectados simultaniamente. Hay " & S1 & " usuarios.", 65, 190, 156, 0, 0)
If txtIndex = 20 Then Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> WorldSave ha concluído.", 0, 185, 0, 0, 0)
If txtIndex = 21 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha quitado " & S2 & " puntos de vida.", 255, 0, 0, 1, 0)
If txtIndex = 22 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tenes suficientes puntos de magia para lanzar este hechizo.", 65, 190, 156, 0, 0)
If txtIndex = 23 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tenes suficiente mana.", 65, 190, 156, 0, 0)
If txtIndex = 24 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mascota desinvocada.", 255, 0, 0, 1, 0)
If txtIndex = 25 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mascota invocada.", 255, 0, 0, 1, 0)
If txtIndex = 26 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Los Dioses te sonrien, has ganado 500 puntos de nobleza!.", 65, 190, 156, 0, 0)
If txtIndex = 27 Then Call AddtoRichTextBox(frmMain.RecTxt, "Le has causado " & S1 & " puntos de daño a la criatura!", 255, 0, 0, 1, 0)
If txtIndex = 28 Then Call AddtoRichTextBox(frmMain.RecTxt, "Le has quitado " & S1 & " puntos de vida a " & S2, 255, 0, 0, 1, 0)
If txtIndex = 29 Then Call AddtoRichTextBox(frmMain.RecTxt, S1 & " te ha quitado " & S2 & " puntos de vida.", 255, 0, 0, 1, 0)
If txtIndex = 30 Then Call AddtoRichTextBox(frmMain.RecTxt, "Te estás concentrando. En " & S1 & " segundos comenzarás a meditar.", 65, 190, 156, 0, 0)
If txtIndex = 31 Then Call AddtoRichTextBox(frmMain.RecTxt, "Comenzas a meditar.", 65, 190, 156, 0, 0)
If txtIndex = 32 Then Call AddtoRichTextBox(frmMain.RecTxt, "Dejas de meditar.", 65, 190, 156, 0, 0)
If txtIndex = 33 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has sanado.", 65, 190, 156, 0, 0)
If txtIndex = 34 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "¡Felicitaciones! has logrado vencer a un animal sagrado, te has ganado el ingreso a la cueva de los sabios para aprender el hechizo de tu clase.", 4, 166, 179, 1, 0)
    Call Audio.PlayWave("HARP3.WAV")
End If

If txtIndex = 35 Then Call AddtoRichTextBox(frmMain.RecTxt, "Mapa exclusivo para newbies.", 65, 190, 156, 0, 0)
If txtIndex = 36 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para ingresar a la cueva de los sabios tienes que derrotar a un animal sagrado previamente.", 65, 190, 156, 0, 0)
If txtIndex = 37 Then Call AddtoRichTextBox(frmMain.RecTxt, "Debes derrotar a un animal sagrado para poder aprender hechizos de clase.", 65, 190, 156, 0, 0)
If txtIndex = 38 Then Call AddtoRichTextBox(frmMain.RecTxt, "El hechizo no puede ser aprendido por tu clase.", 65, 190, 156, 0, 0)
If txtIndex = 39 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has vuelto a ser visible.", 65, 190, 156, 0, 0)
If txtIndex = 40 Then Call AddtoRichTextBox(frmMain.RecTxt, "No hay ninguna subasta en curso.", 65, 190, 156, 0, 0)
If txtIndex = 41 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tienes esa cantidad de monedas para ofertar.", 65, 190, 156, 0, 0)
If txtIndex = 42 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedes ofertar como subastante.", 65, 190, 156, 0, 0)
If txtIndex = 43 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para subastar objetos debes ser nivel 20 o mayor.", 65, 190, 156, 0, 0)
If txtIndex = 44 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para subastar objetos debes tener al menos 20 skills points en Comerciar.", 65, 190, 156, 0, 0)
If txtIndex = 45 Then Call AddtoRichTextBox(frmMain.RecTxt, "El mensaje ha sido entregado, ahora solo debes esperar que se desocupe algun GM.", 65, 190, 156, 0, 0)
If txtIndex = 46 Then Call AddtoRichTextBox(frmMain.RecTxt, "Ya habias mandado un mensaje, tu mensaje ha sido movido al final de la cola de mensajes.", 65, 190, 156, 0, 0)
If txtIndex = 47 Then Call AddtoRichTextBox(frmMain.RecTxt, "El pedido debe contener un mensaje que se adecue a tu problema y el mismo debe ser coherente, caso contrario los GMs no acudirán a tu pedido.", 65, 190, 156, 0, 0)
If txtIndex = 48 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "Tu pregunta ha sido respondida por un GameMaster: " & S1, 19, 215, 209, 1, 0)
    Call Audio.PlayWave("59.WAV")
End If
If txtIndex = 49 Then Call AddtoRichTextBox(frmMain.RecTxt, "Pregunta respondida satisfactoriamente.", 65, 190, 156, 0, 0)
If txtIndex = 50 Then Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global se encuentra desactivado por el momento.", 65, 190, 156, 0, 0)
If txtIndex = 51 Then Call AddtoRichTextBox(frmMain.RecTxt, "Para hablar por mensaje global debes ser nivel 10 como mínimo.", 65, 190, 156, 0, 0)
If txtIndex = 52 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global ha sido activado. Para hablar su mensaje deberá contener el prefijo "".""" & S1, 31, 36, 252, 1, 0)
    Call Audio.PlayWave("43.WAV")
End If
If txtIndex = 53 Then
    Call AddtoRichTextBox(frmMain.RecTxt, "El sistema de mensaje global ha sido desactivado." & S1, 31, 36, 252, 1, 0)
    Call Audio.PlayWave("45.WAV")
End If
If txtIndex = 54 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has revelado tu posición, pierdes el efecto invisibilidad.", 65, 190, 156, 0, 0)
If txtIndex = 55 Then Call AddtoRichTextBox(frmMain.RecTxt, "Este item no puede ser subastado.", 32, 51, 233, 1, 1)
If txtIndex = 56 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has conseguido algo de leña!", 65, 190, 156, 0, 0)
If txtIndex = 57 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has pescado un lindo pez!", 65, 190, 156, 0, 0)
If txtIndex = 58 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has pescado algunos peces!", 65, 190, 156, 0, 0)
If txtIndex = 59 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has extraido algunos minerales!", 65, 190, 156, 0, 0)
If txtIndex = 60 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has obtenido un lingote!", 65, 190, 156, 0, 0)
If txtIndex = 61 Then Call AddtoRichTextBox(frmMain.RecTxt, "No tienes suficientes minerales para hacer un lingote.", 65, 190, 156, 0, 0)
If txtIndex = 62 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has construido el objeto!.", 65, 190, 156, 0, 0)
If txtIndex = 63 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 255, 0, 0, 1, 0)
If txtIndex = 64 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has apuñalado la criatura por " & S1, 255, 0, 0, 1, 0)
If txtIndex = 65 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡No has logrado apuñalar a tu enemigo!", 255, 0, 0, 1, 0)
If txtIndex = 66 Then Call AddtoRichTextBox(frmMain.RecTxt, "Has apuñalado a " & S1 & " por " & S2, 255, 0, 0, 1, 0)
If txtIndex = 67 Then Call AddtoRichTextBox(frmMain.RecTxt, "Te ha apuñalado " & S1 & " por " & S2, 255, 0, 0, 1, 0)
If txtIndex = 68 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡¡Debes esperar unos momentos para tomar otra pocion!!", 65, 190, 156, 0, 0)
If txtIndex = 69 Then Call AddtoRichTextBox(frmMain.RecTxt, "No puedo cargar mas objetos.", 65, 190, 156, 0, 0)
If txtIndex = 70 Then Call AddtoRichTextBox(frmMain.RecTxt, "Solo los newbies pueden usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 71 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu clase no puede usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 72 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu clase,genero o raza no puede usar este objeto.", 65, 190, 156, 0, 0)
If txtIndex = 73 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Debes aproximarte al agua para usar el barco!", 65, 190, 156, 0, 0)
If txtIndex = 74 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes atacar ciudadanos, para hacerlo debes desactivar el seguro apretando la tecla S", 255, 0, 0, 1, 0)
If txtIndex = 75 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando paralizado.", 65, 190, 156, 0, 0)
If txtIndex = 76 Then Call AddtoRichTextBox(frmMain.RecTxt, "No podes teletransportarte a un castillo estando encarcelado.", 65, 190, 156, 0, 0)
If txtIndex = 77 Then Call AddtoRichTextBox(frmMain.RecTxt, "Ya te encuentras en el castillo.", 255, 0, 0, 1, 0)
If txtIndex = 78 Then Call AddtoRichTextBox(frmMain.RecTxt, "GM " & S1 & " acudió a " & S2 & ".", 66, 213, 157, 1, 0)

End Sub

Public Sub txtReceivedT(ByVal txtIndex As Integer, Optional S1 As String, Optional S2 As String, Optional S3 As String, Optional S4 As String, Optional S5 As String)

If txtIndex = 1 Then Call AddtoRichTextBox(frmMain.RecTxt, "Estas muy cansado para luchar.", 65, 190, 156, 0, 0)
If txtIndex = 2 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡ Has formado una party !", 255, 180, 255, 0, 0)
If txtIndex = 3 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu carisma y liderazgo no son suficientes para liderar una party.", 255, 180, 255, 0, 0)
If txtIndex = 4 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu carisma y liderazgo no son suficientes para liderar una party.", 255, 180, 255, 0, 0)
If txtIndex = 5 Then Call AddtoRichTextBox(frmMain.RecTxt, "Tu oferta debe superar las " & S1 & " monedas de oro.", 65, 190, 156, 0, 0)
If txtIndex = 6 Then Call AddtoRichTextBox(frmMain.RecTxt, "Hay " & S1 & " jugadores online.", 255, 255, 255, 1, 0)
If txtIndex = 7 Then Call AddtoRichTextBox(frmMain.RecTxt, "Server> " & S1 & " ha sido expulsado por posible utilización de aplicaciones ilegales.", 0, 185, 0, 0, 0)
If txtIndex = 8 Then Call AddtoRichTextBox(frmMain.RecTxt, "Petición de salida cancelada.", 12, 149, 250, 1, 0)
If txtIndex = 9 Then Call AddtoRichTextBox(frmMain.RecTxt, "Servidor> " & S1 & " ha sido echado por el servidor por posible uso de SH.", 0, 185, 0, 0, 0)
If txtIndex = 10 Then Call AddtoRichTextBox(frmMain.RecTxt, "La descripcion a cambiado.", 65, 190, 156, 0, 0)
If txtIndex = 11 Then Call AddtoRichTextBox(frmMain.RecTxt, "¡Has mejorado tu skill " & S1 & " en un punto!. Ahora tienes " & S2 & " pts.", 65, 190, 156, 0, 0)

End Sub
'Public Const FONTTYPE_TALK As String = "~255~255~255~0~0"
'Public Const FONTTYPE_FIGHT As String = "~255~0~0~1~0"
'Public Const FONTTYPE_WARNING As String = "~32~51~223~1~1"
'Public Const FONTTYPE_INFO As String = "~65~190~156~0~0"
'Public Const FONTTYPE_INFOBOLD As String = "~65~190~156~1~0"
'Public Const FONTTYPE_EJECUCION As String = "~130~130~130~1~0"
'Public Const FONTTYPE_PARTY As String = "~255~180~255~0~0"
'Public Const FONTTYPE_VENENO As String = "~0~255~0~0~0"
'Public Const FONTTYPE_GUILD As String = "~255~255~255~1~0"
'Public Const FONTTYPE_SERVER As String = "~0~185~0~0~0"
'Public Const FONTTYPE_GUILDMSG As String = "~228~199~27~0~0"
'Public Const FONTTYPE_CONSEJO As String = "~130~130~255~1~0"
'Public Const FONTTYPE_CONSEJOCAOS As String = "~255~60~00~1~0"
'Public Const FONTTYPE_CONSEJOVesA As String = "~0~200~255~1~0"
'Public Const FONTTYPE_CONSEJOCAOSVesA As String = "~255~50~0~1~0"
'Public Const FONTTYPE_CENTINELA As String = "~0~255~0~1~0"


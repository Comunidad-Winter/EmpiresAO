# Empires AO

Servidor y Cliente Funcional (con bugs). Empires AO mod de Argentum Online 11.5

## Implementaciones

- Se agregó la jerarquía administrador.
- Cambie los nombres de: (Por minúscula, y consejo de la Luz)
CIUDADANO
*CRIMINAL
*CONSEJO DE BANDERBILL
*CONSEJO DE LAS SOMBRAS
- Al hacer click en un GameMaster, nos dice su respectivo rango.
- Agrege un buscador de Items, /BUSCAR Nombre.
- Cuenta regresiva disponible con el comando /CR.
- Comando /DARORO para realizar transferencias.
- Agregado un ¡Level Up! al pasar el nivel.
- Sistema de Screenshots.
- GameMaster se equipan todo.
- No nos paralizamos solo.
- Le puse el fondo al FrmMain del cliente, pueden sacarselo si desean.
- Los consejeros pueden atacar y agarrar items.
- Agregado Nivel Mínimo para utilizar item. (Minlvl= Nivel necesario)
- ¡Nivel máximo!en FrmMain.
- Sistema de mascota inicial dependendiendo la clase...
- Clases luchadoras -> Elemental de agua.
- Clases luchadoras y magicas -> Elemental de tierra.
- Clases mágicas -> Elemental de fuego.
- Sistema de Duelos.
- Sistema de Gran Poder.
- Sistema de conquista de castillos por clanes.
- Sistema de subastas al estilo comercio.
- Sistema de Caballos (4 tipos) que aumentan un 0,50% la velocidad de tu pj montado.
- Sistema de torneo 1 vs 1 situado el NPC en Hobitton y Banderbill.
- Hechizos por clases.
- Hechizos de área.
- Se agregó una nueva raza: Nigromante (Especializada en invocación).
- Teclas de acceso rápido configurables por los jugadores.
- Mientras juegas al servidor se refleja en el msn que lo estás jugando con un cartelito. en el subnick que indica el nombre del servidor y el pj con el que estás online.
- Gran cantidad de armas, hechizos, armaduras, ropas. 
- Se puso intervalo a la caminata para evitar todo tipo de aceleradores.
- Comercio seguro entre jugadores.
- Sistema de accounts con visualización del personaje y verificación.   
- Posee más caracteristicas de las cuales se tendrán que enterar al jugarlo usted,  o más bien visualizen los códigos. Ahora nos encargaremos de configurarlo para lanzarlo en la red. =D

## Fotos

![Screen1](https://github.com/user-attachments/assets/a3fe8234-833a-4c10-aff8-88b26337087f)

![Screen2](https://github.com/user-attachments/assets/40058cc5-b4f8-4cd4-bea4-b0e243d1b9e6)

![Screen3](https://github.com/user-attachments/assets/ef4112bc-e4c9-4022-9b99-5d52eee44641)

![Screen5](https://github.com/user-attachments/assets/1454f227-cb5a-4b21-87c3-8d012d05ab9a)


## Guia para abrir el servidor

En los códigos de cliente:

Buscamos:

```
  frmConnect.IPTxt.Text = 
  frmConnect.PortTxt.Text =
``` 
Ponemos al lado de los iguales, nuestra dirección de IP y respectivo puerto. Luego de realizar este paso, buscamos en mod_declaraciones:

```
Public Const EAOipserver As String = "127.0.0.1"
Public Const EAOportserver As Integer = "7666"
```

Modificamos según nuestra dirección de IP y puerto, como lo hemos realziado anteriormente. Bien prosigamos al siguiente paso.

Continuamos cambiando la dirección de IP, buscamos:

```
  Private Sub imgServEspana_Click()
      Call Audio.PlayWave(SND_CLICK)
      IPTxt.Text = "62.42.193.233"
      PortTxt.Text = "7666"
  End Sub  
```

Nuevamente configuramos la dirección de IP y demás. Ya que estos cambios hemos concluido con el cliente y podremos generar el .exe tranquilamente. =)

Ahora nos dirigiremos a los códigos del servidor:

Buscaremos la siguiente línea en todo el proyecto:

```
Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSu cuenta ha sido creada satisfactoriamente. Se le ha envíado un mail con el código de validación para su confirmación.")"  
```

Y la reemplazaremos por:

```
Call SendData(SendTarget.ToIndex, UserIndex, 0, "ERRSu cuenta ha sido creada satisfactoriamente. Su codigo de Confirmación es " & ValidateCode & ". La verificación se escribe en MAYUSCULA")
```

Entonces, con esto evitamos los envio de e-mails, y podrán obtener su código de verificación al crear la cuenta.

## NOTA

Este servidor tiene una cuenta creada de pruebas:

* Nombre: Betatester
* Contraseña: 123456

# AoYind 3 - Servidor

Importante, no bajar el codigo con el boton Download as a ZIP de github por que lo descarga mal, muchos archivos por el encoding quedan corruptos.

Tenes que bajar el codigo con un cliente de git, con el cliente original de la linea de comandos seria:
```
git clone https://github.com/YindSoft/aoyind3-server.git
```

## Como utilizar el servidor del juego.

En este repositorio solo se encuentra los codigos de fuente del servidor, por lo tanto para poder ejecutarlo correctamente es necesario tambien clonar el repo de resources para copiar los archivos necesarios.
Esto esta hecho así para poder separar bien los cambios de los recursos del juego en general y lo que es codigo.

Pueden clonar el repo de recursos desde aquí.
```
git clone https://github.com/YindSoft/aoyind3-resources.git
```


Para el servidor es necesario copiar los siguientes archivos/carpetas:
```
Dats/*
Maps/*
Server.ini
```

Por defecto tiene configurado el puerto 7222, si se cambia ese puerto es necesario cambiarlo tambien en el código del cliente.


Es necesario tener instalado MySQL Server, recomendablemente la version 5.7, y deben crear una base de datos e importarle el contenido del archivo aoyind3.sql.
Luego en Server.ini deberan configurar los datos de conexión a la base de datos.

Actualmente no hay un repo con una web para poder crear las cuentas de usuario, asi que en esta version deberan crear una cuenta a mano en la tabla de cuentas, la password deberá ir en MD5, y así podran ingresar desde el cliente y crear personajes. 

## F.A.Q:

#### Error - Al abrir el proyecto en Visual Basic 6 no puede cargar todas las dependencias:
Este es un error comun que les suele pasar a varias personas, esto es debido que el EOL del archivo esta corrupto.
Visual Basic 6 lee el .vbp en CLRF, hay varias formas de solucionarlo:

Opcion a:
Con Notepad++ cambiar el EOL del archivo a CLRF

Opcion b:
Abrir un editor de texto y reemplazar todos los `'\n'` por `'\r\n'`

--------------------------


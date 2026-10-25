# PoliMusic — Aplicación web en ASP Clásico (VBScript)

Reproductor y administrador de música estilo Spotify, desarrollado **estrictamente
en ASP Clásico con VBScript** (archivos `.asp`). Sin ASP.NET, sin C#, sin Razor.

## Requisitos

- Windows con **IIS** y el módulo **ASP** habilitado (o IIS Express para desarrollo).
- **SQL Server** accesible en `localhost,14333` con la base `BDD_PoliMusic`.
- Credenciales usadas por la aplicación: usuario `sa`, contraseña `Politecnica1`
  (configurables en `includes/dbconn.asp`).
- Base de datos creada con el script `BDD_PoliMusic (1).sql` (incluye datos semilla).

## Estructura

```
webapppolimusicv1_asp/
├── default.asp            Redirige a index.asp o login.asp según la sesión
├── login.asp              Autenticación contra TBL_USER (texto plano o SHA-1)
├── logout.asp             Session.Abandon() y vuelta al login
├── index.asp              Pantalla principal: "Song List", bienvenida y reproductor
├── play.asp               Endpoint AJAX: incrementa PLAYS
├── songs/artistas/álbumes
│   ├── *_form.asp         Crear / editar (solo administradores)
│   └── *_delete.asp       Confirmación y eliminación (solo administradores)
├── artists.asp            Listado de artistas
├── albums.asp             Listado de álbumes
├── includes/
│   ├── dbconn.asp         Conexión ADO parametrizada + helpers DbQuery/DbExec
│   ├── helpers.asp        Encoding, rutas de recursos, flash, SHA-1 puro
│   ├── auth.asp           Guardia de sesión / RequireAdmin
│   ├── layout.asp         Chrome común (sidebar, topbar, alertas)
│   └── layout_foot.asp    Cierre del chrome común
├── css/styles.css         Sistema de diseño propio (tema oscuro, acento coral)
├── js/app.js              Toasts, buscador, confirmaciones
├── js/player.js           Reproductor HTML5 persistente
├── assets/img/            Logo, favicon, portadas por defecto, vinilo (SVG)
└── SongFilesAndImages_resources/   MP3 e imágenes oficiales
```

## Usuarios de prueba (datos semilla)

| Usuario      | Contraseña | Rol                |
|--------------|------------|--------------------|
| userNormal   | 12345678   | Normal (USER_TYPE 0) |
| userAdmin    | 12345678   | Admin  (USER_TYPE 1)  |
| username1    | 12345678   | Normal (password almacenada como SHA-1) |

## Control de roles

- **Rol Normal (0)**: solo visualiza canciones/artistas/álbumes y reproduce música.
- **Rol Admin (1)**: CRUD completo sobre canciones (`TBL_SONG` + `TBL_PLAYER` +
  `TBL_TRACKLIST`), artistas (`TBL_ARTIST`) y álbumes (`TBL_ALBUM`).
  Las páginas de administración validan el rol en el servidor
  (`RequireAdmin`) y además ocultan las acciones en la interfaz.

## Rutas de recursos

La base guarda rutas estilo ASP.NET (`~/songFiles/x.mp3`). `includes/helpers.asp`
las traduce a la carpeta local `SongFilesAndImages_resources/<Subcarpeta>/...`
y, si el archivo físico no existe, usa imágenes por defecto de `assets/img/`
(los MP3 ausentes se marcan como "No disponible" en el listado).

## Despliegue en IIS

1. Copiar la carpeta al servidor (p. ej. `C:/inetpub/wwwroot/polimusic`).
2. En IIS, convertir la carpeta en aplicación (Application Pool integrado;
   ASP clásico funciona en modo integrado o clásico).
3. Verificar que la característica **ASP** esté instalada:
   `Turn Windows features on or off → Internet Information Services →
   World Wide Web Services → Application Development Features → ASP`.
4. Ajustar `includes/dbconn.asp` si cambian servidor/base/credenciales.
5. Abrir `http://localhost/polimusic/default.asp`.

### Probar con IIS Express (desarrollo, sin admin)

IIS Express lee su configuración de `%USERPROFILE%\Documents\IISExpress\config\applicationhost.config`
(**nota:** en equipos con OneDrive redirigido, "Documents" apunta a `OneDrive\Documentos`).
El modo `/path:` + `/port:` no siempre sirve los archivos; la forma probada es usar
un sitio del config:

1. Editar `applicationhost.config` y apuntar el sitio `WebSite1` a la carpeta del proyecto:
   `<site name="WebSite1" id="1"><application path="/" applicationPool="Clr4IntegratedAppPool">
   <virtualDirectory path="/" physicalPath="C:/ruta/webapppolimusicv1_asp" /></application>
   <bindings><binding protocol="http" bindingInformation="*:8080:localhost" /></bindings></site>`
   (el handler `ASPClassic` para `*.asp` ya viene en la sección `handlers`).
2. Levantar el servidor:

```
"C:/Program Files/IIS Express/iisexpress.exe" /config:"%USERPROFILE%/Documents/IISExpress/config/applicationhost.config" /site:WebSite1
```

3. Abrir `http://localhost:8080/login.asp`.

Para IIS completo, la aplicación funciona tal cual (ver "Despliegue en IIS" arriba).

## Verificación realizada

La aplicación se verificó extremo a extremo contra SQL Server `localhost,14333`:

| Prueba | Resultado |
|--------|-----------|
| Suite SQL (esquema, restricciones, borrados en cascada manual) | 36/36 |
| SHA-1 puro VBScript vs. vectores conocidos (Python hashlib) | 4/4 |
| E2E HTTP (IIS Express): login/logout, roles, listado, CRUD, contador | 74/74 |

El E2E cubre: sesión anónima redirigida, login correcto/fallido, saludo
personalizado y cabecera "Song List", 5 canciones renderizadas, reproductor
presente, usuario normal sin enlaces de administración, **POST de usuario normal
a endpoints de admin rechazados sin mutar la base**, CRUD completo de
canciones/artistas/álbumes con verificación de relaciones y de ausencia de
huérfanos, contador `PLAYS`, `play.asp` sin sesión o con id inválido, y salida
con destrucción de sesión. Estado final de la base tras las pruebas:
5 canciones, 5 artistas, 5 álbumes (sin cambios permanentes).

## Notas de seguridad implementadas

- Todas las consultas con datos de usuario son **parametrizadas** (ADODB.Command).
- Toda la salida se codifica con `Server.HTMLEncode` (anti-XSS).
- Sesiones con timeout de 60 minutos; `Session.Abandon` al salir.
- Las contraseñas se comparan contra texto plano o hash SHA-1 hexadecimal
  (implementación pura en VBScript, sin dependencias externas).
- Cada `Response.Redirect` va seguido de `Response.End`: en ASP clásico el script
  **continúa ejecutándose** tras una redirección, por lo que sin el `End` un
  usuario normal que enviara un POST a `song_delete.asp` habría borrado la
  canción aunque el navegador recibiera la redirección.

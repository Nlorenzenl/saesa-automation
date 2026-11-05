# Centrality 2.0

Agenda de PT que se actualiza sola desde Centrality cada 3 horas, se publica en GitHub Pages protegida con contraseña y envía un correo con los cambios.

## Qué hay en esta carpeta

| Archivo | Para qué sirve |
|---|---|
| `sync_centrality.py` | Lee Centrality, compara con la lectura anterior, genera la agenda y envía el correo |
| `plantilla.html` | Diseño de la agenda (tabla, calendario, alertas, ficha con condiciones y adjuntos) |
| `encriptar_pagina.py` | Cifra la agenda con la contraseña antes de publicarla |
| `../.github/workflows/centrality2.yml` | Ejecuta todo cada 3 horas en GitHub Actions y publica la página |

## Por qué la página va cifrada

GitHub Pages gratis solo funciona con repositorios **públicos**, y la página queda en internet. Por eso la agenda se publica cifrada: quien abra el enlace solo ve un campo de contraseña, y sin ella los datos no se pueden leer. La foto de datos que se guarda entre ejecuciones (rama `centrality2-datos`) también va cifrada con la misma contraseña.

Los registros de cada ejecución en la pestaña Actions de un repositorio público los puede ver cualquiera. Solo muestran cantidades (PT leídos, fichas, cambios) y, si falla alguna ficha, hasta 3 números de PT.

## Instalación

### 1. Elegir el repositorio

- Si tu repositorio de automatización de PT **ya es público**, úsalo: los secretos de Centrality y Gmail ya están ahí.
- Si es **privado**, crea uno nuevo y público (por ejemplo `centrality2`) en github.com → **New repository** → Public. En ese caso tendrás que crear también los secretos del paso 3.

### 2. Subir los archivos

En el repositorio: **Add file** → **Upload files** y arrastra:

- la carpeta `centrality2/` completa
- la carpeta `.github/` (contiene `workflows/centrality2.yml`)

Si la carpeta `.github` no se deja arrastrar: **Add file** → **Create new file**, escribe `.github/workflows/centrality2.yml` como nombre y pega el contenido del archivo.

Luego **Commit changes**.

### 3. Secretos

**Settings** → **Secrets and variables** → **Actions** → **New repository secret**:

| Secreto | Valor |
|---|---|
| `CENTRALITY2_CLAVE` | **Nuevo.** La contraseña de la página. Larga (mínimo 12 caracteres); es la que le darás al equipo. |
| `SAESA_USER`, `SAESA_PASS` | Usuario y contraseña de Centrality (ya existen si usas el repositorio actual) |
| `GMAIL_USER`, `GMAIL_APP_PASS`, `EMAIL_DEST` | Correo que envía y destinatario (ya existen si usas el repositorio actual) |
| `CENTRALITY2_EMAIL_CC` | Opcional: correos en copia, separados por coma |

### 4. Activar GitHub Pages

**Settings** → **Pages** → en **Source** elige **GitHub Actions**.

### 5. Primera carga

**Actions** → **Centrality 2.0** → **Run workflow**. Tarda cerca de una hora. Al terminar:

- la página queda en `https://TU-USUARIO.github.io/NOMBRE-DEL-REPO/`
- llega el correo "primera carga" con el botón **Abrir la agenda en la web**

Desde ahí corre sola cada 3 horas.

## Uso de la página

- Al entrar pide la contraseña. Con "Recordar en este equipo" no la vuelve a pedir en ese navegador.
- Para cambiar la contraseña: edita el secreto `CENTRALITY2_CLAVE`. La siguiente ejecución no podrá abrir la foto anterior y hará una carga completa (unos 60 minutos, con correo de "primera carga").

## Qué hace cada ejecución

- Lee todos los PT con inicio entre `DESDE` y `HASTA` (hoy 01/01/2026 a 31/12/2026; se cambian en el workflow).
- Suma todos los PT de la bandeja de programación ("En programación de trabajo - bandeja de entrada"), aunque su fecha esté fuera del rango. En la agenda aparecen como etapa **Bandeja de programación**.
- Lee la ficha completa (subestación, descripción, condiciones y bloqueos, adjuntos) de los PT nuevos, de los que cambiaron y de los que parten entre 3 días atrás y 14 días adelante. El resto usa la ficha guardada.
- Detecta PT nuevos, cambios de estado, de fechas y de programador, y PT que salieron del rango.
- Publica la agenda cifrada en GitHub Pages. El correo lleva el resumen de cambios y el enlace a la página, sin adjunto (los correos corporativos suelen bloquear adjuntos .html grandes).

## Tiempos

- Primera carga (01/01 a 31/12): entre 2 y 3 horas.
- Ejecuciones siguientes: pocos minutos.
- Para releer todas las fichas: Run workflow con "Volver a leer la ficha completa".
- GitHub puede atrasar las ejecuciones programadas entre 5 y 30 minutos en horas de alta demanda.

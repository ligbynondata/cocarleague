# LIG - Arquitectura actual del proyecto

## 1. Producto vivo actual

La app operativa actual vive en la carpeta local:

`Escritorio / Cocar`

Archivo principal:

`index.html`

Este archivo corresponde al front actual de LIG para la comunidad activa Cocar League.

No debe ser reemplazado por una web pública ni renombrado sin una migración controlada.

## 2. Backend operativo actual

El backend actual vive en Google Sheets:

`PADEL NONDATA`

Este Sheet contiene la data deportiva, tablas intermedias y Apps Script de la app actual.

Incluye, entre otras hojas:

- matches_raw
- division_standings
- division_fixtures
- app_players
- league_players
- ranking_last14
- league_stats_summary
- league_stats_rankings
- stats_indiv
- stats_rachas

Este Sheet es el backend oficial de Cocar League.

## 3. Núcleo futuro LIG

Existe un Sheet separado llamado:

`LIG_CORE`

Su objetivo futuro es administrar identidad global, comunidades, jugadores e inscripciones.

Hojas actuales:

- players_master
- communities
- community_players
- registrations_log

Por ahora LIG_CORE no está conectado a la app operativa.

## 4. Drive

La carpeta madre en Drive es:

`PADEL NONDATA`

Estructura:

- 01_LIG_CORE
- 02_COMUNIDADES
- 03_PRODUCTO_APP
- 04_MARKETING
- 99_ARCHIVO

Mover archivos dentro de Drive no rompe conexiones si se mueve el archivo original y no una copia.

## 5. Reglas de trabajo

- `index.html` sigue siendo la app viva actual.
- No se pisa `index.html` con una web pública.
- La web pública futura usará `home.html`.
- El registro futuro de comunidades usará `registro.html`.
- Cocar League es la primera comunidad activa de LIG.
- Nuevas comunidades usarán landing, QR, formulario y validación.
- Cocar no requiere landing pública de inscripción por ahora.

## 6. Git

Los cambios del front se suben desde la carpeta local `Cocar`.

Comandos:

```bash
git status
git add .
git commit -m "mensaje del cambio"
git push
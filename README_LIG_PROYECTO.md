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

## 7. Inventario de archivos locales

La carpeta local oficial del proyecto es:

`Escritorio / Cocar`

Archivos actuales:

| Archivo / carpeta | Estado | Uso |
|---|---|---|
| `index.html` | OFICIAL | Front vivo actual de la app LIG para Cocar League. |
| `assets/` | OFICIAL | Logos e imágenes usadas por el front. |
| `README_LIG_PROYECTO.md` | OFICIAL | Documento base de arquitectura y reglas del proyecto. |
| `index_backup.html` | BACKUP | Copia de seguridad actualizada de `index.html`. |
| `landing.html` | EN REVISIÓN | Archivo antiguo/experimental. No usar como base oficial sin revisarlo. |

Reglas:

- Todo cambio funcional de la app se hace primero en `index.html`.
- Después de un cambio importante estable, se puede actualizar `index_backup.html`.
- `landing.html` no debe usarse como fuente oficial.
- Si se crea una web pública de LIG, deberá hacerse en un archivo nuevo, probablemente `home.html`.
- Si se crea un flujo de registro para nuevas comunidades, deberá hacerse en un archivo nuevo, probablemente `registro.html`.
## 8. Arquitectura multicomunidad

LIG debe operar como sistema multicomunidad.

### Backend actual

`PADEL NONDATA` es el backend operativo actual de la comunidad Cocar League.

No debe convertirse en el backend único de todas las comunidades.

### Núcleo global

`LIG_CORE` será el núcleo global futuro de LIG.

Su rol será administrar:

- jugadores globales
- comunidades
- relación jugador-comunidad
- registros e inscripciones

### Backend por comunidad

Cada nueva comunidad deberá tener su propio Sheet operativo basado en una plantilla común.

Ejemplos futuros:

- `LIG_COCAR_OPERATIVO`
- `LIG_CLUB_X_OPERATIVO`
- `LIG_EMPRESA_Y_OPERATIVO`

Cada Sheet operativo contendrá sus propias hojas deportivas:

- matches_raw
- division_standings
- division_fixtures
- ranking_last14
- league_stats_summary
- league_stats_rankings
- app_players
- league_players
- stats_indiv
- stats_rachas
- Apps Script operativo

### Front

El front debe evolucionar hacia un único código parametrizable por comunidad.

Conceptualmente:

`app.html?community=cocar`

`app.html?community=club_x`

Cada comunidad tendrá su propio:

- sheet_id
- logo
- nombre visible
- colores
- estado

### Regla base

LIG_CORE administra identidad y pertenencia.

Cada Sheet operativo administra competencia.

El front muestra la comunidad seleccionada.
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
- app_player_profiles

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
```

Los cambios hechos directamente en Google Apps Script no quedan versionados en Git salvo que se copie el código a un archivo local.

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
- app_player_profiles
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

## 9. app_player_profiles

`app_player_profiles` es una hoja APP EXPORT creada para alimentar la futura pantalla LIG Players.

Su función es consolidar en una sola fila por jugador información proveniente de:

- app_players
- ranking_last14
- stats_indiv
- stats_rachas
- division_standings

Esta hoja no se edita manualmente. Se reconstruye mediante:

`buildAppPlayerProfiles()`

y forma parte del flujo oficial:

`refreshAll()`

Uso futuro:

- identidad del jugador
- estado actual
- posición en tabla de división
- ranking Last14
- racha actual
- resumen histórico
- perfil competitivo
- tags narrativos iniciales

Regla:

`app_player_profiles` será la fuente principal para construir la pantalla LIG Players en el front.

## 10. Roadmap LIG Players

LIG Players debe funcionar como la hoja de vida deportiva del jugador dentro de LIG.

La pantalla futura debería mostrar:

### Identidad

- Foto
- Nombre
- División actual
- Rol si aplica

### Estado actual

- División actual
- Posición en tabla de su división
- Puntos en tabla
- Ranking LIG Last14
- LIG Points
- Racha actual
- Record Last14

### Carrera LIG

- Partidos ganados/perdidos
- Sets ganados/perdidos
- Games ganados/perdidos
- Win rates
- Rendimiento por división
- Rendimiento en partidos de 2 sets y 3 sets

### Perfil competitivo

- Tie breaks
- Tight sets
- Hard sets
- Easy sets
- Comebacks
- Collapses / farras
- Rendimiento en terceros sets

### Rachas

- Racha actual
- Mejor racha ganadora histórica
- Peor racha perdedora histórica
- Últimos 5 partidos
- Últimos 10 partidos

### Historial

La pantalla deberá incorporar en una etapa posterior una tabla de partidos jugados por jugador, incluyendo partidos reales y WO cuando corresponda.

Esta tabla probablemente deberá provenir de una nueva hoja export:

`app_player_match_history`

### Logros y badges

En una etapa posterior se deberá crear una capa de logros:

`app_player_achievements`

Ejemplos:

- Campeón de división
- Podio de división
- Ascenso logrado
- Número 1 Last14
- Top 3 Last14
- Tie Break King
- Comeback Story
- Mejor racha histórica

### Rivalidades / H2H destacado

En una etapa posterior se podrá crear una hoja:

`app_player_rivalries`

Posibles bloques:

- Rival clásico
- Némesis
- Cliente
- Duelo más parejo

## 11. Deuda técnica conocida

Antes de escalar LIG a nuevas comunidades, se deben revisar los siguientes puntos:

- Estandarizar nombres de hojas, especialmente `Players` versus `players`.
- Revisar `player_rank_snapshots` antes de mostrar evolución de ranking, porque podría tener lógica antigua de escala.
- Revisar `app_league_selector`, ya que el front actual todavía usa temporadas hardcodeadas.
- Definir `LIG_TEMPLATE_COMUNIDAD_OPERATIVA` como copia estructural limpia del backend actual, sin datos reales de Cocar.
- Versionar o respaldar Apps Script en archivos locales si se quiere control total por Git.
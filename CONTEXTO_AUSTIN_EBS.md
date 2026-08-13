# Austin EBS — Sistema de Gestión de Admisiones
## Contexto para Claude Code

---

## 🚨 PRINCIPIOS INAMOVIBLES

1. **Todo el sistema vive en un único archivo HTML** (`index.html`). No crear archivos separados, no dividir en módulos, no agregar dependencias npm. Una sola página auto-contenida.
2. **No cambiar la estética actual** — colores, tipografía, componentes visuales. Mantener el design system existente al pie de la letra.
3. **Mínima complejidad** — preferir soluciones simples sobre elegantes. Menos código es mejor código. Evitar abstracciones innecesarias.
4. **Datos en localStorage** — no hay backend. Todo persiste en `localStorage` con la clave `austin_ebs_v2`.

---

## 📁 Archivo principal

- **Ruta local**: `C:\Users\alvaj\OneDrive\Documents\Admisiones AEBS\index.html` (~2941 líneas)
- **Live**: `https://joaquinalva-wq.github.io/AdmisionesAustin/`
- **Repo**: `https://github.com/joaquinalva-wq/AdmisionesAustin`
- **Credenciales panel**: `admisiones@austinebs-ah.edu.ar` / `austin2026`

---

## 🎨 Design System (NO modificar)

```css
/* Variables CSS */
--or: #E8621A     /* naranja — acción principal */
--gr: #7CB518     /* verde — éxito, matrícula */
--tl: #3A607A     /* teal claro */
--td: #2C4A5E     /* teal oscuro — headings, navbar */
--tl2: #5A8BA8    /* teal medio */
--cr: #F8F6F0     /* crema fondo */
--cr2: #EEE9DF    /* crema oscuro */
--wh: #fff
--tx: #252520     /* texto principal */
--tx2: #5A5A54    /* texto secundario */
--tx3: #95958E    /* texto suave */
--bd: rgba(58,96,122,.13)
--bd2: rgba(58,96,122,.22)
```

**Tipografías**: `'Lora'` (serif, headings) + sistema sans-serif para body.

**Clases de botón**: `.sbtn .sbtn-n` (next/azul), `.sbtn .sbtn-b` (back/blanco), `.sbtn-ok` (verde).

---

## 🏗️ Estructura del HTML (en orden)

```
<head>        — CSS completo
<body>
  #s-portal       — Portal público para familias
  #s-postulacion  — Formulario multi-paso
    #fst-tar      — Paso: tarifas (acepta y continúa)
    #fst-num      — Paso: ¿cuántos estudiantes? (1–4)
    #fst0         — Paso: datos estudiante (se repite N veces)
    #fst1         — Paso: datos familia
    #fst2         — Paso: calendario entrevista (incluye botón "aún no sé")
    #fst3         — Paso: resumen + enviar
    #fst-ok       — Pantalla de éxito
  #s-login        — Login admin
  #s-panel        — Panel de admisiones
    .sidebar      — Barra lateral con menú
    #v-dashboard  — Vista principal (pipeline)
    #v-entrevistas— Vista entrevistas
    #v-bajas      — Vista bajas (nuevo)
    #v-config     — Configuración
```

---

## 💾 Modelo de datos (DB)

```javascript
let DB = {
  nextId: 1328,
  postulantes: [...],
  vacantes: { 'curso': {t: total, u: usadas} },
  tarifas: { k2_m, k2_s, k2_c, k3_m, k3_s, k3_c, k45_m, k45_c,
             el_m, el_c, hs13_m, hs13_c, hs46_m, hs46_c },
  emails: { contactado, citado, entrevistado, postulado, eoe_ingles,
            matriculado, 'lista de espera', desiste, desiste_recontacto,
            no_queremos, seguimiento_entrevista, no_asistio },
  config: { nombre, email, calendar, drive, appsScriptUrl, eoeApiUrl, emailDG },
  pendingNotif: []
}
```

**Schema de postulante:**
```javascript
{
  id: Number,
  alumno: "Nombre Apellido",
  padre: "Nombre Apellido",
  email: "string",
  tel: "string",
  curso: "K2 (Puppies)" | "1er año Elementary" | etc,
  nivel: "Kindergarten" | "Elementary" | "High School",
  jornada: "Jornada Completa" | "Jornada Simple",   // Jornada Simple solo para K2
  anio: 2026 | 2027 | 2028,
  colegio: "string",
  hermano: "si" | "no",
  estado: string,           // ver ESTADOS abajo
  fecha: "YYYY-MM-DD",
  ultAccion: "YYYY-MM-DD",
  entFecha: "YYYY-MM-DD",
  entHora: "HH:MM",
  dgDeadline: "YYYY-MM-DD",  // deadline 48hs para pendiente_dg
  eoeDocId: "string",        // ID del documento Firestore en evaluaciones-admisiones
  notas: [{txt, fecha, hora}],
  hist: [{e: estado, f: fecha, mailOk?: bool, auto?: bool}]
}
```

---

## 📊 Pipeline de estados (FLUJO COMPLETO)

```
const ESTADOS = ['contactado','citado','entrevistado','postulado',
                 'eoe_ingles','pendiente_dg','matriculado',
                 'lista de espera','rechazado_eoe','desiste','no_queremos']
```

**Flujo normal:**
```
(Formulario web con fecha) → citado
(Formulario web sin fecha) → contactado → [familia agenda luego] → citado
citado → entrevistado → postulado → matriculado (jardin/kinder directo)
                                 → eoe_ingles (secondary/elementary)
                                       ↓ EOE aprueba
                                 pendiente_dg (48hs auto-avance o manual DG)
                                       ↓
                                 matriculado
```

**Acciones automáticas por estado:**
- `contactado`: mail a familia (bienvenida + link para agendar)
- `eoe_ingles`: mail a familia + POST a Firestore `evaluaciones-admisiones`
- `pendiente_dg`: mail interno a DG (emailDG en config), deadline = hoy+48h
- `matriculado`: mail bienvenida a familia + mail resumen a administracion@ + sync EOE Firestore
- `rechazado_eoe`: sin mail automático (admisiones notifica manualmente)

**Colores de estados:**
```javascript
const ECOL = {
  contactado:'#1A7A8E', citado:'var(--cc)', entrevistado:'var(--ce)',
  postulado:'var(--cp)', eoe_ingles:'#7B52AB', pendiente_dg:'#B45309',
  rechazado_eoe:'#DC2626', matriculado:'var(--cm)',
  'lista de espera':'var(--cs)', desiste:'var(--cd)', no_queremos:'var(--cd)'
}
```

---

## 🔑 Funciones clave

| Función | Descripción |
|---------|-------------|
| `go(id)` | Muestra pantalla por id |
| `goPost()` | Abre formulario de postulación |
| `doLogin()` | Valida credenciales y muestra panel |
| `saveDB()` / `loadDB()` | Persistencia localStorage |
| `today()` | Retorna fecha YYYY-MM-DD |
| `toast(msg, type)` | Notificación toast |
| `NIVEL(curso)` | Mapea curso a nivel |
| `renderDashboard()` | Re-renderiza panel principal |
| `enviarMail(p, estado, cb)` | Envía mail via Apps Script (familias, sin fallback) |
| `enviarMailAdmin(p, estado, cb)` | Envía mail desde panel admin (con fallback mailto) |
| `enviarMailDG(p)` | Mail interno a Dirección General (cuando EOE aprueba) |
| `enviarMailAdministracion(p)` | Mail resumen a administracion@ (al matricular) |
| `apiCall(action, data)` | Llama al Apps Script |
| `changeE(id)` | Cambia estado del postulante. Checkea checkbox "solo-cat-{id}" para omitir mail |
| `aprobarEOE(id)` | EOE aprobó → pendiente_dg + mail a DG |
| `rechazarEOE(id)` | EOE rechazó → rechazado_eoe |
| `confirmarDG(id)` | DG confirmó → matriculado + todos los mails |
| `checkPendienteDG()` | Se corre al abrir el panel. Auto-avanza pendiente_dg si pasaron 48hs |
| `sincronizarConEOE(p)` | POST a Firestore matriculados-pendientes (al matricular) |
| `sincronizarConEOEEvaluacion(p)` | POST a Firestore evaluaciones-admisiones (al pasar a eoe_ingles) |
| `renderBajas()` | Renderiza vista de bajas |
| `gcalLink(p)` | Genera URL de Google Calendar para crear evento de entrevista |
| `toggleJornada()` | Muestra/oculta opción "Jornada Simple" (solo para K2) |
| `sinFechaEntrevista()` | Flujo "aún no sé" → estado contactado |

---

## 📋 Flujo del formulario (multi-estudiante)

```
goPost()
  → fst-tar (revisar tarifas + check)
  → fNextTar() → fst-num (selector 1–4 estudiantes)
  → setNumStudents(n) + goToStudentForm()
  → fst0 (datos estudiante — se repite)
        * Jornada Simple solo aparece si curso = K2
  → fst1 (datos familia)
  → fst2 (calendario entrevista)
        * [Confirmar entrevista →] → fst3 con fecha (estado: citado)
        * [Aún no sé — agendar más adelante] → fst3 sin fecha (estado: contactado)
  → fst3 (resumen)
  → submitPost() → crea registros en DB, envía mail según estado
  → fst-ok (éxito)
```

---

## 📅 Integración Apps Script (mails automáticos)

**Configuración:**
1. Google Sheets de admisiones → Extensiones → Apps Script
2. Pegar código → desplegar como Web App (ejecutar como `admisiones@`, acceso: cualquiera)
3. Copiar URL `/exec` → Panel → Configuración → campo verde → Guardar

**Pendiente (actualización manual en Apps Script):**
- Agregar creación automática de eventos en Google Calendar cuando se llama `reservarTurno`
- Código de referencia: `CalendarApp.getCalendarById('admisiones@austinebs-ah.edu.ar').createEvent(titulo, inicio, fin, opciones)`

**Acciones disponibles:**
- `reservarTurno` — reserva slot + envía mail + (pendiente: crea evento Google Calendar)
- `getTurnos` / `getBloqueos` / `setBloqueo` / `ping`
- `sendMail` — envío genérico de mail

---

## 🔗 Integraciones externas

- **Google Calendar** (turnos familias): `https://calendar.app.google/WrfYGGSiRH3smkxL8`
- **Google Drive** (tarifas): carpeta con PDFs
- **EOE Firestore** — dos colecciones:
  - `matriculados-pendientes` — al matricular (función `sincronizarConEOE`)
  - `evaluaciones-admisiones` — al pasar a EOE/Inglés (función `sincronizarConEOEEvaluacion`)
    - Campos: nombre, apellido, nivel, curso, anioIngreso, email, telefono, nombrePadre, admisionesId, estado ('Pendiente evaluación'), timestamp
- **Administración** — mail a `administracion@austinebs-ah.edu.ar` al matricular
- **Dirección General** — mail a `emailDG` (configurable) al aprobar EOE

---

## 🖥️ Vistas del panel

- **Dashboard** — stats, entrevistas hoy, requieren acción
- **Pipeline** — columnas por estado (auto-fill, 11 estados)
- **Entrevistas** — calendario semanal + lista
- **Postulantes** — tabla completa con filtros (incluye nuevos estados)
- **Bajas** (nuevo) — tabla de desistes con selección múltiple y mail masivo de recontacto
- **Vacantes** — gestión de vacantes por curso
- **Plantillas de mail** — editor de templates (excluye rechazado_eoe, pendiente_dg, no_queremos)
- **Tarifas** — configuración de aranceles
- **Configuración** — Apps Script URL, email remitente, Drive link, Email DG (nuevo)

---

## ⚠️ Notas importantes para desarrollo

- El archivo tiene logos embebidos como base64 (NO modificar esos bloques)
- `SEED_POSTULANTES` es un array JSON gigante en el JS — no tocar
- `checkPendienteDG()` se ejecuta cada vez que se abre el panel (en `renderPanel()`)
- El checkbox "Solo mover categoría" está en cada modal de gestión (`id="solo-cat-{id}"`)
- La vista Bajas tiene selección múltiple y mail masivo de recontacto
- El botón "Agregar a Google Calendar" usa `gcalLink(p)` que genera una URL con el evento pre-llenado (abre Google Calendar del usuario — para auto-crear en el calendario de Admisiones se debe actualizar el Apps Script)
- `Jornada Simple` solo aparece en el formulario cuando el curso seleccionado es K2 (Puppies)

---

## 🚀 Deploy

```bash
git add index.html
git commit -m "descripción del cambio"
git push origin main
# GitHub Pages se actualiza en ~1 min
```

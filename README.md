# CEVA Zebra PWA v2 — Control de Etiquetas

PWA estática 100% client-side para generar el archivo de control de Handling Units Zebra.  
**No requiere servidor, Node.js, ni scripts .bat.**

---

## Cómo usar

1. Abrir `index.html` en el navegador (o la URL publicada en GitHub Pages).
2. Cargar el **HR en formato XLSX** (Hoja de Ruta exportada desde el sistema).
3. Cargar uno o más archivos **Impresos XLSX**.
4. (Opcional) Cargar **Ce.xlsx o Ce.json** con la tabla Localidad → Ruta.
5. Presionar **Procesar y Descargar** — el XLSX de etiquetas se descarga automáticamente.

---

## Diferencias respecto a v1

| Aspecto | v1 (Node.js + Express) | v2 (Estática) |
|---|---|---|
| **Backend** | Express.js en Node.js | Ninguno — 100% browser |
| **Instalación** | `npm install` + `.bat` scripts | Solo abrir el HTML |
| **HR input** | PDF (parsing regex) o XLSX | **Solo XLSX** (recomendado) |
| **PDF support** | Sí (pdf-parse en server) | No (limitación de GitHub Pages) |
| **Deploy** | Servidor propio / VPS | GitHub Pages / cualquier hosting estático |
| **Persistencia** | Sesiones en RAM del servidor | Sin estado — procesamiento en memoria del browser |
| **Offline** | No (requería server activo) | Sí (PWA con Service Worker, tras primera carga) |
| **Instalable** | No como PWA estándar | Sí — installable en Zebra y móviles |

---

## Por qué no funciona PDF en v2

GitHub Pages sirve archivos estáticos — no puede ejecutar código de servidor.  
El parsing de PDF en v1 usaba `pdf-parse` (Node.js), que accede al sistema de archivos del server.  
En el browser, las APIs disponibles para PDF son limitadas y el formato de los romaneos CEVA  
(tablas complejas, layout de columnas) hace que el parsing regex sea frágil sin control total.

**Solución recomendada:** Exportar el romaneo como XLSX desde el sistema de gestión o abrir  
el PDF en Excel y guardarlo como .xlsx antes de cargarlo en esta app.

---

## Limitaciones de GitHub Pages

1. **Sin backend** — todo el procesamiento ocurre en el browser del usuario.
2. **Sin autenticación** — el sitio es público si el repositorio es público.
3. **Sin logs del servidor** — no hay registro de errores centralizados; solo la consola del browser.
4. **Tamaño de archivos** — el browser puede manejar archivos grandes, pero depende de la RAM del dispositivo. Para archivos >20MB se recomienda usar v1 con servidor.
5. **PDF no soportado** — ver sección anterior.
6. **CORS** — no aplica para archivos locales cargados por el usuario (FileReader API no tiene restricciones CORS).
7. **Service Worker scope** — en GitHub Pages con subdirectorios, el SW debe estar en la raíz del directorio del proyecto.

---

## Compatibilidad

- **Zebra TC-series** (Android Chrome): ✅ Funcional, installable como PWA
- **Android Chrome**: ✅ Funcional
- **iOS Safari**: ✅ Funcional (PWA installable desde Safari → Compartir → Agregar a inicio)
- **Desktop Chrome/Edge/Firefox**: ✅ Funcional
- **Desktop Safari**: ✅ Funcional

---

## Formato de archivos de entrada

### HR XLSX
Columnas requeridas (detección flexible, sin importar mayúsculas/tildes):
- `Nº Factura` (o `Nro Factura`, `Num Factura`, etc.)
- `Destinatario`
- `Ciudad Destinatario`

### Impresos XLSX
Columnas requeridas:
- `Remito` (o `Referencia`)
- `Handling Unit` (o `HU`, `Unidad`)

### Ce.xlsx / Ce.json
- **Excel**: Primera hoja con columnas `Localidad` y `Ruta`
- **JSON**: Array de objetos `[{"localidad": "CORDOBA", "ruta": "R1"}]`

---

## Deploy en GitHub Pages

```bash
# Desde la raíz del repositorio
git add .
git commit -m "CEVA Zebra PWA v2 — static deploy"
git push origin main

# En GitHub: Settings → Pages → Source: Deploy from branch → main → / (root)
# O apuntar al subdirectorio /ceva-zebra-pwa-v2/
```

Si el repo está en un subdirectorio, asegurarse de que el `manifest.json` use paths relativos (`./`), lo cual ya está configurado.

---

## Lógica de procesamiento

```
HR.xlsx ──────┐
              ├── loadHR() → normalizar columnas
              │
Impresos.xlsx ┤── loadImpresos() → extraer Remito + Handling Unit
(1 o varios)  │
              │
Ce.xlsx/JSON ─┘── loadCe() → Localidad → Ruta
                      │
                      ▼
              processData()
              1. Derivar Nº Factura desde Remito
                 "0300-00018664" → "18664" (quita prefijo + leading zeros)
              2. Cross-check: alertas de facturas sin match
              3. Merge: Impresos + HR por Nº Factura
              4. Merge: + Ce por Ciudad (normalizado: sin tildes, lowercase)
              5. Deduplicar por Handling Unit
              6. Generar XLSX hoja "Etiquetas"
                      │
                      ▼
              CEVA_Zebra_Etiquetas_YYYYMMDD_HHmm.xlsx
              Columnas: Referencia | Handling Unit | Destino | Ciudad | Ruta
```

---

## Estructura del proyecto

```
ceva-zebra-pwa-v2/
├── index.html       # App completa (HTML + CSS + JS inline, sin build step)
├── manifest.json    # PWA Manifest
├── sw.js            # Service Worker (cache-first)
├── ceva-logo.jpg    # Logo CEVA Logistics
├── icon-192.png     # PWA icon 192x192
├── icon-512.png     # PWA icon 512x512
└── README.md        # Este archivo
```

---

## Tecnologías

- **HTML / CSS / JavaScript** puro — sin frameworks, sin build step
- **SheetJS (xlsx)** via CDN — lectura y escritura de XLSX en el browser
- **Google Fonts (DM Sans)** — tipografía corporativa
- **Service Worker** — PWA installable, cache offline
- **FileReader API** — lectura de archivos locales sin upload a servidor

---

*CEVA Logistics — Control de Etiquetas Zebra v2 · Desarrollado para uso interno*

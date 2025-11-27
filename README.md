# Generador de gafetes 10x10 cm

Aplicación web (React + Vite) para generar gafetes de 10x10 cm listos para impresión a doble cara. Carga un archivo Excel con las columnas **Empresa**, **Apellido** y **Nombre** y la herramienta coloca automáticamente el nombre (Nombre Apellido) y la empresa en la plantilla, respetando tipografía y posiciones del diseño de referencia.

## Requisitos
- Node.js 18+
- npm

## Ejecución local
```bash
npm install
npm run dev
```
Abre el navegador en la URL indicada (por defecto http://localhost:5173).

## Uso
1. Prepara un Excel (.xlsx) con las columnas `Empresa`, `Apellido` y `Nombre`.
2. Pulsa **Subir Excel** y selecciona el archivo.
3. Verifica la vista previa; cada persona genera dos páginas (frente y reverso) en un formato de **10 cm x 10 cm**.
4. Pulsa **Imprimir (frente y reverso)** y en el diálogo del navegador selecciona impresión a doble cara con encuadernación por borde largo.

También puedes probar el flujo con el botón **Cargar ejemplo** que agrega datos de muestra.

## Despliegue en Vercel
1. Crea un proyecto en Vercel apuntando a este repositorio.
2. Usa los comandos predeterminados de Vercel:
   - Build command: `npm run build`
   - Output dir: `dist`
3. Vercel detectará automáticamente Vite y publicará los archivos estáticos.

## Notas de diseño
- Tipografía preferente: `Futura` (con respaldo `Futura PT`, `Century Gothic`, `Arial`, `sans-serif`).
- Tamaños: Nombre 33.3pt en negrita, Empresa 22.6pt en negrita, leyenda 5.9pt sin negrita.
- La hoja de impresión está fijada a 100mm x 100mm con márgenes 0 para alinear frente y reverso.

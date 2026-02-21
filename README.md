# Gastos diarios → Excel (.xlsx)

Pequeña app web **100% estática** (HTML/CSS/JS) para registrar gastos diarios y **descargar un Excel** (`.xlsx`) directamente desde el navegador.

## Qué incluye

- **Formulario** para cargar gastos.
- **Tabla editable** (podés corregir celdas antes de exportar).
- Botón **“Descargar Excel”** que genera un `.xlsx` con:
  - Hoja **Gastos** (todas las filas)
  - Hoja **Resumen** (total general + totales por categoría y por día)

## Cómo usarlo (local)

Recomendado: servirlo con un servidor local (evita restricciones del navegador).

```bash
python3 -m http.server 8000
```

Luego abrí en el navegador:

- `http://localhost:8000/`

## Columnas del Excel

- **Fecha** (YYYY-MM-DD)
- **Categoría**
- **Descripción**
- **Método** (Efectivo / Débito / Crédito / Transferencia / Otro)
- **Monto**

## Notas

- La librería de Excel se carga desde CDN (`xlsx` / SheetJS). Necesitás conexión a internet para exportar.

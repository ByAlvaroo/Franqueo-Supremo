# ✉ Franqueo Supremo

App de escritorio para Windows que te dice **qué sellos pegar** en cada envío (importe exacto),
gastando primero los de más valor o sacando el máximo de envíos, y lleva la cuenta de tus sellos en un Excel.

**HECHO POR @BYALVARO__**

## Descargar

Ve a **Releases** (columna derecha de esta página) → última versión → descarga **FranqueoSupremo.exe**.
No necesita Python ni instalar nada.

1. Crea una carpeta (p. ej. `Documentos\Franqueo`) y mete ahí `FranqueoSupremo.exe`.
2. Ábrelo. La primera vez crea **Mis_Sellos.xlsx** al lado, con tus sellos y tarifas.
   (Si ya tenías un `Mis_Sellos.xlsx`, ponlo en esa carpeta antes de abrir.)
3. Si Windows dice «Windows protegió su PC»: *Más información* → *Ejecutar de todas formas* (solo la primera vez).

## Actualizaciones

Cada vez que se suben cambios a este repositorio, GitHub crea solo un `.exe` nuevo
(pestaña **Actions**) y lo publica en **Releases**. Al abrir la app te sale un aviso
**«Hay una versión nueva»** → *Actualizar ahora* → se descarga, se cambia y se vuelve a abrir.
Tu Excel no se toca.

> El repositorio tiene que ser **público** para que la app pueda ver las versiones nuevas.
> Tus sellos **no** se suben: el `.gitignore` excluye los `.xlsx`.

## El Excel (`Mis_Sellos.xlsx`)

- **Sellos**: columna B = cantidad, columna C = valor.
- **Tarifas**: dos tablas separadas a la derecha: **ORDINARIO** (columnas F-G) y **CERTIFICADO** (columnas I-J).
  Para crear una tarifa nueva escribe nombre y precio en la primera fila libre debajo de su tabla.
- **Historial**: cada envío que descuentas.

## Estructura

| Archivo | Qué es |
|---|---|
| `codigo/franqueo_supremo.py` | La ventana de la app |
| `codigo/motor.py` | Los cálculos y el Excel |
| `codigo/icono.ico` | Icono |
| `version.txt` | Versión base (GitHub le añade un número cada vez) |
| `.github/workflows/crear-exe.yml` | Receta para que GitHub cree el `.exe` |

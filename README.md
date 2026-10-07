# FDC Selector — complemento Excel para el Flujo de Caja Maritime

Panel lateral que sincroniza filas fuente al consolidado:
**añadir** (faltantes, en su sitio, con amarillos), **eliminar** (solo
principal u hoja fuente también) y **reparar vínculos** rotos.

## Desarrollo local

```powershell
pnpm install
pnpm test        # chequeos del motor
pnpm run build   # genera dist/taskpane.js
pnpm dlx office-addin-dev-certs install
pnpm run watch   # recompila al guardar (terminal 1)
pnpm run dev-server  # https://localhost:3000 (terminal 2)
```

Sideload en Excel escritorio: Archivo > Opciones > Centro de confianza >
Catálogos de complementos > agregar carpeta con `manifest.xml`.
Luego Insertar > Mis complementos.

## Publicar (multi-PC)

1. Crear repo público `fdc-selector` y subir esto.
2. Activar Settings > Pages (Source: GitHub Actions). El workflow
   publica `dist/` en `https://TU-USUARIO.github.io/fdc-selector/`.
3. En `manifest.xml`, cambiar las URLs `https://localhost:3000/...` por
   las de Pages (SourceLocation, Commands.Url, Taskpane.Url, iconos).
4. En cada Excel: confiar el catálogo con ese `manifest.xml` una sola vez.
   Actualizar = `git push`; los Excel toman la versión nueva solos.

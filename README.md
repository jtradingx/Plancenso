# Dashboard Planificación — Censo 2026 Ejecutivo

Dashboard interactivo de seguimiento de proyectos, publicado en GitHub Pages y actualizado automáticamente desde OneDrive.

🔗 **Ver dashboard:** https://jtradingx.github.io/Plancenso/

---

## Configuración inicial (solo una vez)

### 1. Obtener la URL de descarga del Excel en OneDrive

1. Abre el archivo Excel en OneDrive o SharePoint
2. Haz clic en los tres puntos `...` → **Detalles**
3. Copia el link de descarga directa.
   - Si usas "Compartir → Copiar vínculo", reemplaza `?web=1` por `?download=1` al final de la URL
   - La URL debe terminar en algo como `...xlsx?download=1`

### 2. Guardar la URL como secret en GitHub

1. Ve a tu repositorio en GitHub
2. **Settings → Secrets and variables → Actions → New repository secret**
3. Nombre: `ONEDRIVE_URL`
4. Valor: pega la URL del paso anterior
5. Haz clic en **Add secret**

### 3. Activar GitHub Pages (si no está activo)

1. **Settings → Pages**
2. Branch: `main` → carpeta `/ (root)` → **Save**

---

## Actualización del dashboard

El dashboard se regenera automáticamente en dos situaciones:

| Trigger | Cuándo |
|---|---|
| ⏰ **Diario** | Todos los días a las 8:00 AM (hora Chile) |
| 📤 **Al subir archivos** | Cuando subes un `.xlsx` o modificas `generar_dashboard.py` |
| ▶️ **Manual** | Desde la pestaña **Actions → Actualizar Dashboard → Run workflow** |

---

## Estructura del repositorio

```
Plancenso/
├── index.html                          # Dashboard (generado automáticamente)
├── generar_dashboard.py                # Script que lee el Excel y genera el HTML
├── .github/
│   └── workflows/
│       └── actualizar_dashboard.yml    # Workflow de GitHub Actions
└── README.md
```

---

## Actualización manual del Excel

Si prefieres actualizar subiendo el Excel directamente al repo (sin OneDrive):

1. Renombra el archivo como `Censo_2026_Ejecutivo.xlsx`
2. Súbelo al repositorio (reemplazando el anterior)
3. GitHub Actions lo detectará y regenerará el dashboard automáticamente

---

## Dependencias Python

```
pandas
openpyxl
requests
```

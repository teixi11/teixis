# 📊 Sistema de KPIs en Tiempo Real - Factory Explorer Pro

## ✅ Implementación Completada

Se ha añadido un sistema completo de lectura de Excel y visualización de KPIs en tiempo real.

---

## 🚀 Cómo Usar

### 1. **Abrir la Aplicación**
- Abre `copia_factory.html` en tu navegador
- Espera a que cargue la pantalla inicial

### 2. **Acceder al Panel de KPIs**
- Haz click en el botón **"📊 Panel KPIs Excel"**
- Se abrirá un panel lateral a la derecha

### 3. **Cargar tu Archivo Excel**
- Click en el área de carga **"📁 Click para cargar archivo Excel"**
- Selecciona tu archivo `.xlsx` o `.xls`
- También puedes usar el archivo de ejemplo: `kpis_ejemplo.csv` (guárdalo como Excel)

### 4. **Ver KPIs en Tiempo Real**
- Los KPIs se mostrarán automáticamente
- Con la opción **"Auto-actualizar con estaciones"** activada, el panel mostrará solo los KPIs de la estación actual del tour
- Desactívala para ver todos los KPIs simultáneamente

---

## 📋 Formato del Excel

Tu archivo Excel debe tener esta estructura:

| Columna A (Estación) | Columna B (KPI)              | Columna C (Valor) | Columna D (Unidad) |
|---------------------|------------------------------|-------------------|-------------------|
| Recepción           | Eficiencia                   | 92                | %                 |
| Recepción           | Materiales Procesados        | 245               | unidades          |
| Carrocería          | Eficiencia                   | 95                | %                 |
| Carrocería          | Soldaduras Completadas       | 180               | unidades          |
| Pintura             | Eficiencia                   | 88                | %                 |
| Powertrain          | Motores Ensamblados          | 85                | unidades          |
| Ensamblaje          | Vehículos Completados        | 42                | unidades          |
| Control Calidad     | Inspecciones Realizadas      | 128               | unidades          |
| Expedición          | Envíos Completados           | 38                | unidades          |

**Notas:**
- La primera fila puede ser encabezados (se detecta automáticamente)
- La columna D (Unidad) es opcional
- Los nombres de estaciones se normalizan automáticamente (acepta variaciones como "Recepción", "Recepcion", etc.)

---

## 🎨 Características

### ✨ Visualización Dinámica
- **Tarjetas de KPI** con valores destacados
- **Colores inteligentes**: 
  - 🟢 Verde para valores positivos (>80%)
  - 🔴 Rojo para valores negativos (<60%)
- **Animaciones** suaves al cambiar de estación

### 🔄 Actualización Automática
- Se sincronizan con el tour cinemático
- Actualización cada 2 segundos
- Indicador visual de actualización activa

### 📍 Mapeo de Estaciones
El sistema reconoce estas estaciones (y sus variaciones):
- ✅ Recepción / Recepcion
- ✅ Carrocería / Carroceria
- ✅ Pintura
- ✅ Powertrain / Motor
- ✅ Interiores / Interior
- ✅ Ensamblaje / Ensamblaje Final / Final
- ✅ Expedición / Expedicion
- ✅ Control de Calidad / Calidad
- ✅ Mantenimiento

---

## 💾 Ejemplo de Uso

### Opción 1: Usar el archivo CSV de ejemplo
1. Abre `kpis_ejemplo.csv` con Excel
2. Guárdalo como `.xlsx`
3. Cárgalo en la aplicación

### Opción 2: Crear tu propio Excel
1. Crea un nuevo archivo Excel
2. Añade las 4 columnas: Estación, KPI, Valor, Unidad
3. Rellena con tus datos
4. Guárdalo y cárgalo

---

## 🔧 Solución de Problemas

### ❌ "Error al leer el archivo Excel"
- **Causa**: Formato incorrecto del archivo
- **Solución**: Verifica que tenga las 3 columnas mínimas (Estación, KPI, Valor)

### ❌ Los KPIs no se muestran
- **Causa**: Los nombres de estación no coinciden
- **Solución**: Usa los nombres exactos de la lista de estaciones (Recepción, Carrocería, Pintura, etc.)

### ❌ No se actualiza automáticamente
- **Causa**: Opción desactivada
- **Solución**: Activa el checkbox "Auto-actualizar con estaciones"

---

## 🎯 Próximos Pasos (Opcionales)

Si quieres conectar con **PostgreSQL**:
1. Necesitarás crear un backend (Node.js o Python)
2. El backend leerá el Excel y lo cargará a PostgreSQL
3. La aplicación consultará los KPIs desde la base de datos

**¿Te gustaría que implemente la versión con PostgreSQL?**

---

## 📞 Soporte

Si tienes dudas o necesitas modificaciones:
- Los datos se leen con la librería **SheetJS** (https://sheetjs.com/)
- El código está en `copia_factory.html` en la sección `=== FUNCIONES EXCEL KPIs ===`
- Puedes modificar los colores, formato y comportamiento según necesites

---

## 🎉 ¡Listo!

Ahora tu aplicación puede:
- ✅ Leer archivos Excel directamente desde el navegador
- ✅ Mostrar KPIs en tiempo real
- ✅ Sincronizar con las estaciones del tour
- ✅ Actualizar automáticamente según la posición

**¡Disfruta de tu Factory Explorer Pro con KPIs en tiempo real!** 🏭📊

# 🏭 Sistema de Unificación de Datos de Producción

Sistema automatizado para gestión, cálculo y optimización de producción de etiquetas industriales. Integra datos de SAP, calcula cilindros óptimos, gestiona inventarios cascadeados y genera reportes detallados en Excel.

---

## 📋 Tabla de Contenidos

- [Características](#-características)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Documentación](#-documentación)
- [Contribución](#-contribución)

---

## ✨ Características

### Motor de Cálculo Avanzado
- **Optimización de cilindros Z** (60-168) para minimizar metraje
- **Cálculo de desarrollo** basado en factor de 3.175 mm
- **Distribución horizontal/vertical** de etiquetas
- **Gestión de gaps** (horizontal y vertical)
- **Cálculo de repeticiones** y merma (0.75 mm/repetición)

### Gestión de Inventarios
- **Stock cascadeado** por componente y fecha
- **Normalización de unidades** (UNIDAD, MILES, ROLLO)
- **Validación de stocks negativos**
- **Trazabilidad** de consumos

### Reportes Profesionales
- **Excel multi-hoja** con formato condicional
- **Colores según disponibilidad** (verde/rojo)
- **31+ columnas** de análisis detallado
- **Exportación automática** a ubicación configurable

### Integración SQL Server
- **Conexión optimizada** con pool de conexiones
- **Queries eficientes** con límites y filtros
- **Manejo de errores** y reintentos
- **Soporte para SAP B1**

---

## 📦 Requisitos

### Software
- Python 3.8+
- SQL Server (ODBC Driver 17+)
- Microsoft Excel (para visualización de reportes)

### Dependencias Python
```txt
pandas>=1.3.0
sqlalchemy>=1.4.0
openpyxl>=3.0.0
pyodbc>=4.0.0
python-dotenv>=0.19.0
```

---

## 🚀 Instalación

### 1. Clonar el repositorio
```bash
git clone https://github.com/ifdeek/IMAGENES.git
cd IMAGENES
```

### 2. Crear entorno virtual
```bash
python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate
```

### 3. Instalar dependencias
```bash
pip install -r requirements.txt
```

### 4. Configurar credenciales
**⚠️ IMPORTANTE: No commitear credenciales al repositorio**

Crear archivo `.env` en la raíz del proyecto:
```env
DB_SERVER=10.101.2.181
DB_NAME=SAP_G02E05_Innoprint
DB_USERNAME=ReportesInnoprint
DB_PASSWORD=tu_password_aqui
```

### 5. Verificar instalación
```bash
python unificacion.py
```

---

## 💻 Uso

### Ejecución básica
```bash
python unificacion.py
```

El script:
1. Se conecta a SQL Server
2. Extrae datos de pedidos, stock y materiales
3. Ejecuta motor de cálculo de cilindros
4. Procesa cascadeo de stock
5. Genera reporte Excel: `tablas_unificadas.xlsx`

### Ubicación de salida

El archivo se guarda automáticamente en:
- **Como .exe:** Carpeta del ejecutable
- **Como .py con permisos:** Carpeta del script
- **Sin permisos:** Carpeta Descargas del usuario
- **Fallback:** Carpeta temporal del sistema

### Interpretación del reporte

#### Columnas clave:

**Motor de Cálculo (Lado Izquierdo)**
- `Z_Cilindro`: Cilindro óptimo seleccionado (60-168)
- `Desarrollo_mm`: Perímetro del cilindro en mm
- `Etiquetas_Desarrollo`: Etiquetas verticales
- `Etiquetas_Eje`: Etiquetas horizontales
- `Metros_Lineales`: ML necesarios sin merma
- `Metros_Cuadrados`: M² necesarios sin merma
- `Stock_Final_m2`: Stock restante después de producción

**Cálculos por Factor (Lado Derecho)**
- `Pendiente`: Cantidad pendiente de venta
- `Factor_unidades`: Cantidad de componente por unidad
- `Metros_Cuadrados_Factor`: M² calculados por factor
- `Stock_Final_Factor`: Stock restante según factor

**Disponibilidad**
- `Disponibilidad`: 1 (verde) = hay stock, 0 (rojo) = sin stock
- `Disponibilidad_Factor`: Disponibilidad según cálculo por factor

---

## 📂 Estructura del Proyecto

```
IMAGENES/
├── unificacion.py              # Script principal (1,571 líneas)
├── Book1.csv                   # Pedidos y componentes (10.6 MB)
├── lista_de_materiales.csv     # BOM - Lista de materiales
├── stock_acumulado_codigo.csv  # Inventario consolidado
├── tabla_costos_actualizada_Version2.csv  # Costos operacionales
├── AISA1-8.svg/png             # Diagramas de proceso
├── Logo.png                    # Logo corporativo
├── ANALISIS_Y_MEJORAS.md      # ✅ Análisis detallado y mejoras
└── README.md                   # ✅ Este archivo
```

---

## 📚 Documentación

### Archivos de documentación
- **[ANALISIS_Y_MEJORAS.md](./ANALISIS_Y_MEJORAS.md)** - Análisis completo del proyecto con mejoras sugeridas

### Conceptos clave

#### Motor de Cálculo de Cilindros
El sistema optimiza la selección del cilindro Z para minimizar el metraje de material:

1. **Factor de conversión:** Z × 3.175 = Desarrollo (mm)
2. **Distribución vertical:** Calcula cuántas etiquetas caben en el desarrollo
3. **Distribución horizontal:** Calcula cuántas etiquetas caben en el ancho del rollo
4. **Gap vertical:** Espacio entre etiquetas (objetivo: 2.7 mm)
5. **Gap horizontal:** Espacio lateral (objetivo: 2.7 mm)
6. **Optimización:** Selecciona el Z que minimiza metros lineales

#### Ejemplo de cálculo:
```
Etiqueta: 55 × 40 mm
Rollo: 330 mm ancho
Z seleccionado: 91 (desarrollo = 288.93 mm)

Vertical: 288.93 mm ÷ 55 mm = 5 etiquetas (gap: 2.79 mm)
Horizontal: 310 mm ÷ 40 mm = 7 etiquetas (gap: 2.86 mm)
Total por repetición: 5 × 7 = 35 etiquetas

Para 10,000 etiquetas:
Repeticiones: 10,000 ÷ 35 = 285.71
ML: 285.71 × 288.93 mm = 82.57 m
M²: 82.57 m × 0.33 m = 27.25 m²
```

#### Stock Cascadeado
El sistema mantiene trazabilidad del stock por componente:

1. **Stock inicial:** Obtenido de base de datos
2. **Normalización:** Según unidad de medida (UNIDAD, MILES, ROLLO)
3. **Cascadeo:** Por componente y fecha de entrega
4. **Consumo con merma:** M² × 1.12 (12% adicional)
5. **Stock final:** Stock anterior - Consumo

---

## 🔧 Configuración Avanzada

### Constantes del motor de cálculo

Editar en `unificacion.py` (líneas 109-120):

```python
FACTOR_CILINDRO_A_MM = 3.175        # Conversión Z a mm
ROLLO_BASE_MM = 330                  # Ancho estándar de rollo
BANDERA_HORIZONTAL = 20              # Margen lateral (mm)
GAP_HORIZONTAL_OBJETIVO = 2.7        # Gap horizontal óptimo
GAP_VERTICAL_OBJETIVO = 2.7          # Gap vertical óptimo
MAX_REPETICIONES_VERTICALES = 8      # Máximo de etiquetas verticales
AJUSTE_DESARROLLO_REP_MM = 0.75      # Merma por repetición (mm)
```

### Cilindros disponibles

Editar lista en líneas 122-123:
```python
CILINDROS_FB = [60, 67, 70, 74, 77, 80, 84, 88, 91, 97, 99, 102, 
                105, 107, 108, 111, 116, 117, 122, 127, 129, 168]
```

---

## ⚠️ Problemas Comunes

### Error: "No se pudo conectar a la base de datos"
**Causa:** Credenciales incorrectas o servidor inaccesible

**Solución:**
1. Verificar archivo `.env`
2. Comprobar conectividad: `ping 10.101.2.181`
3. Validar ODBC Driver 17: `odbcinst -q -d`

### Error: "pyodbc no instalado"
**Solución:**
```bash
pip install pyodbc
```

En Linux, instalar dependencias:
```bash
sudo apt-get install unixodbc unixodbc-dev
```

### Archivo Excel bloqueado
**Causa:** Archivo anterior abierto en Excel

**Solución:**
1. Cerrar Excel
2. El sistema generará `tablas_unificadas_temp.xlsx`

### Performance lento con muchos registros
**Causa:** Procesamiento fila por fila

**Solución (futura):**
- Ver [ANALISIS_Y_MEJORAS.md](./ANALISIS_Y_MEJORAS.md) - Sección "Optimización de rendimiento"
- Considerar vectorización con NumPy

---

## 🛡️ Seguridad

### ⚠️ IMPORTANTE - Protección de credenciales

**NUNCA commitear:**
- Archivo `.env`
- Credenciales en código
- Tokens de acceso
- Información sensible

**Siempre usar `.gitignore`:**
```gitignore
.env
*.env
.env.*
__pycache__/
*.pyc
*.log
tablas_unificadas*.xlsx
```

---

## 🤝 Contribución

### Cómo contribuir

1. **Fork** del repositorio
2. Crear **branch** para feature: `git checkout -b feature/nueva-funcionalidad`
3. **Commit** cambios: `git commit -m 'Agregar nueva funcionalidad'`
4. **Push** al branch: `git push origin feature/nueva-funcionalidad`
5. Crear **Pull Request**

### Guía de estilo

- **PEP 8** para código Python
- **Docstrings** en todas las funciones públicas
- **Type hints** en funciones críticas
- **Tests unitarios** para nueva funcionalidad

---

## 📊 Roadmap

### Versión 1.1 (Q1 2025)
- [ ] Migrar credenciales a `.env`
- [ ] Implementar logging estructurado
- [ ] Modularizar código en paquetes
- [ ] Agregar tests unitarios básicos

### Versión 1.2 (Q2 2025)
- [ ] Optimización con NumPy/vectorización
- [ ] Cache de resultados frecuentes
- [ ] CLI con argumentos
- [ ] Configuración YAML

### Versión 2.0 (Q3 2025)
- [ ] Interfaz web (Flask/FastAPI)
- [ ] Dashboard interactivo
- [ ] API REST para integración
- [ ] Reportes en tiempo real

---

## 📄 Licencia

[Especificar licencia aquí]

---

## 👥 Autores

- **Equipo Innoprint** - Desarrollo inicial

---

## 🙏 Agradecimientos

- SAP Business One por integración ERP
- Comunidad Python por librerías open-source
- [Agregar otros reconocimientos]

---

## 📞 Soporte

Para reportar bugs o solicitar features:
- **Issues:** https://github.com/ifdeek/IMAGENES/issues
- **Email:** [email de contacto]
- **Documentación:** [ANALISIS_Y_MEJORAS.md](./ANALISIS_Y_MEJORAS.md)

---

**Última actualización:** 2025-11-11  
**Versión:** 1.0.0

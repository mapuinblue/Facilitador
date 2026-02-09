# DIAN → Siigo | Conversor Contable

![Versión](https://img.shields.io/badge/versión-2.6-FF69B4)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/licencia-MIT-FF1493)

Aplicación de escritorio para automatizar el procesamiento contable de documentos electrónicos descargados de la DIAN (Dirección de Impuestos y Aduanas Nacionales de Colombia) al formato requerido por el software contable **Siigo**.

<img width="1012" height="851" alt="Image" src="https://github.com/user-attachments/assets/2cda51ca-9bc9-406a-816e-d47916872482" />

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Características](#-características)
- [Requisitos](#-requisitos)
- [Uso](#-uso)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Formato de Salida](#-formato-de-salida)
- [Solución de Problemas](#-solución-de-problemas)
- [Instalación](#-instalación)
- [Contribuciones](#-contribuciones)
- [Licencia](#-licencia)

## 🎯 Descripción

En Colombia, las empresas deben reportar sus operaciones de compra y venta ante la DIAN mediante facturación electrónica. Este proceso genera archivos Excel con información detallada que debe ser transformada manualmente para su importación en sistemas contables como Siigo.

**Este proyecto automatiza esa transformación**, permitiendo a contadores y administradores convertir archivos de la DIAN en el formato exacto requerido por Siigo, ahorrando horas de trabajo manual y eliminando errores de transcripción.

### Casos de Uso

- **Compras/Recibidos**: Procesa facturas de proveedores para generar asientos contables de gastos e IVA descontable
- **Ventas/Enviados**: Procesa facturas emitidas a clientes para generar asientos de ingresos e IVA generado

## ✨ Características

### 🎨 Interfaz de Usuario
- **Diseño intuitivo** con paleta de colores pasteles en tonos rosas
- **Detección automática** del tipo de documento (compras/ventas) por nombre de archivo
- **Log de procesamiento en tiempo real** con información detallada
- **Barra de progreso visual** durante el procesamiento

### ⚡ Funcionalidades Principales
- **Lectura inteligente**: Detecta automáticamente la estructura del archivo DIAN (encabezados variables)
- **Procesamiento dual**: Maneja tanto archivos de compras (Recibidos) como de ventas (Enviados)
- **Cálculos automáticos**:
  - Valor base del IVA (IVA ÷ 0.19)
  - Redondeo a peso colombiano sin decimales
  - Formato de pesos colombiano (ej: `200.000,00`)
- **Filtrado inteligente**: Solo procesa facturas electrónicas, ignorando Application Responses
- **Exportación flexible**: Genera archivos Excel listos para Siigo o código Power Query para importación directa

### 🔧 Robustez
- Manejo de errores con mensajes descriptivos
- Detección automática de columnas por patrones (si los nombres varían)
- Soporte para archivos Excel (.xlsx, .xls) y CSV
- Validación de datos antes del procesamiento

## 💻 Requisitos

- **Python 3.8** o superior
- **Sistema operativo**: Windows, macOS o Linux
- **Dependencias**:
  - pandas >= 1.3.0
  - openpyxl >= 3.0.0
  - tkinter (incluido en Python estándar)

 ## 📖 Uso

### Paso 1: Descargar archivos de la DIAN

1. Ingresa al portal de la DIAN
2. Descarga los reportes de:
   - **Documentos Recibidos** (para compras)
   - **Documentos Enviados** (para ventas)

### Paso 2: Procesar con la aplicación

1. Abre la aplicación **DIAN → Siigo**
2. Haz clic en **"Buscar Archivo"** y selecciona el archivo Excel descargado
3. El tipo de documento se detectará automáticamente (o selecciónalo manualmente)
4. Presiona **"Procesar Archivo"**
5. Espera la confirmación de procesamiento exitoso

### Paso 3: Exportar resultados

- **"Descargar Excel"**: Guarda un archivo .xlsx listo para copiar a Siigo
- **"Power Query"**: Genera código M para importación directa en Excel
- **"Ver Vista Previa"**: Revisa los datos antes de exportar

## 📁 Estructura del Proyecto
dian-a-siigo/
│
├── dian_a_siigo.py          # Código principal de la aplicación
├── README.md                # Este archivo
├── requirements.txt         # Dependencias del proyecto
├── screenshots/             # Capturas de pantalla
│   └── interfaz.png
├── examples/                # Archivos de ejemplo (opcional)
│   ├── recibidos_ejemplo.xlsx
│   └── enviados_ejemplo.xlsx
└── dist/                    # Ejecutables generados (opcional)
└── DIANaSiigo.exe


## 📊 Formato de Salida

### Para Compras (Recibidos)

| CUENTA | CC | OBSERVACIONES | DÉBITO | CRÉDITO | VALOR_BASE | TERCERO | H |
|--------|----|---------------|--------|---------|------------|---------|---|
| 14, 51, 61 | | Nombre Proveedor | 200.000,00 | | | 860069497 | |
| 24080103 | | Nombre Proveedor | 38.000,00 | | 200.000,00 | 860069497 | 1 |

**Lógica:**
- **Cuenta 14,51,61**: Gasto (Total - IVA) en débito
- **Cuenta 24080103**: IVA descontable en débito, con factor 1 en columna H
- **Valor Base**: IVA ÷ 0.19 (redondeado a peso)

### Para Ventas (Enviados)

| CUENTA | CC | OBSERVACIONES | DÉBITO | CRÉDITO | VALOR_BASE | TERCERO | H |
|--------|----|---------------|--------|---------|------------|---------|---|
| 41 | | Nombre Cliente | | 200.000,00 | | 860069497 | |
| 24080101 | | Nombre Cliente | | 38.000,00 | 200.000,00 | 860069497 | 1 |
| 13050501 | | Nombre Cliente | 38.000,00 | | 200.000,00 | 860069497 | |

**Lógica:**
- **Cuenta 41**: Ingresos (Total - IVA) en crédito
- **Cuenta 24080101**: IVA generado en crédito, con factor 1 en columna H
- **Cuenta 13050501**: IVA en débito (contra partida)

## 🔧 Solución de Problemas

### Error: "No se encontraron facturas"

- Verifica que el archivo descargado de la DIAN no esté vacío
- Asegúrate de que el archivo tenga el formato estándar de la DIAN

### Error: "No se generaron registros"

- Revisa que las facturas tengan valores en las columnas Total e IVA
- Verifica que no sean solo "Application Response" (acuses de recibo)

### Las columnas no se detectan correctamente

- La aplicación intenta detectar columnas por patrones de nombre
- Si el formato de la DIAN cambia, revisa el log de depuración para ver qué columnas se detectaron

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Si encuentras errores o tienes mejoras:

1. Haz fork del proyecto
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -m 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un Pull Request

### Mejoras futuras planeadas

- [ ] Soporte para múltiples archivos simultáneos
- [ ] Validación de NITs contra base de datos de la DIAN
- [ ] Generación automática de asientos de retenciones
- [ ] Exportación directa a API de Siigo
- [ ] Versión web para uso sin instalación

## 🚀 Instalación

### Opción 1: Ejecutar con Python

1. **Clona el repositorio**:
   ```bash
   git clone https://github.com/mapuinblue/Facilitador.git
   cd dian-a-siigo
2. **Instala las dependencias**:
   ```bash
   pip install pandas openpyxl
3. **Ejecuta la aplicación**:
   ```bash
   python dian_a_siigo.py

### Opción 2: Crear ejecutable (.exe) para Windows

Si deseas distribuir la aplicación a usuarios sin Python instalado:
   ```bash
   pip install pyinstaller
   pyinstaller --onefile --windowed --name "DIANaSiigo" dian_a_siigo.py

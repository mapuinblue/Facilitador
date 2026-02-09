# DIAN → Siigo | Conversor Contable

![Versión](https://img.shields.io/badge/versión-2.6-FF69B4)
![Python](https://img.shields.io/badge/Python-3.8+-3776AB?logo=python&logoColor=white)
![License](https://img.shields.io/badge/licencia-MIT-FF1493)

Aplicación de escritorio para automatizar el procesamiento contable de documentos electrónicos descargados de la DIAN (Dirección de Impuestos y Aduanas Nacionales de Colombia) al formato requerido por el software contable **Siigo**.

![Interfaz de la aplicación](screenshots/interfaz.png)

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Características](#-características)
- [Requisitos](#-requisitos)
- [Instalación](#-instalación)
- [Uso](#-uso)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Formato de Salida](#-formato-de-salida)
- [Solución de Problemas](#-solución-de-problemas)
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

## 🚀 Instalación

### Opción 1: Ejecutar con Python

1. **Clona el repositorio**:
   ```bash
   git clone https://github.com/tuusuario/dian-a-siigo.git
   cd dian-a-siigo

# 📄 PrintFactory

> **Generador Inteligente de Formatos de Facturación Especial**

Una aplicación profesional desarrollada en Python que automatiza la creación de documentos PDF para facturación especial, con agrupación inteligente de clientes y generación de reportes listos para imprimir.

---

## ✨ Características Principales

✅ **Carga de datos desde Excel** - Importa información de clientes desde archivos .xlsx  
✅ **Previsualización en tiempo real** - Revisa los datos antes de generar PDFs  
✅ **Agrupación inteligente** - Agrupa automáticamente por dirección, segmento y titular de pago  
✅ **Generación en lotes** - Crea múltiples PDFs en segundos  
✅ **Paginación automática** - Maneja grupos grandes dividiendo en páginas  
✅ **Personalización de logo** - Integra tu logo empresarial en los documentos  
✅ **Interfaz moderna** - GUI amigable con CustomTkinter  
✅ **Exportable a ejecutable** - Distribuye como .exe sin dependencias de Python  

---

## 🚀 Instalación

### Requisitos Previos
- **Python 3.8+**
- **pip** (gestor de paquetes)

### Pasos de Instalación

1. **Clonar el repositorio**
```bash
git clone https://github.com/tu-usuario/PrintFactory.git
cd PrintFactory
```

2. **Crear entorno virtual** (recomendado)
```powershell
python -m venv venv
.\venv\Scripts\Activate.ps1
```

3. **Instalar dependencias**
```powershell
pip install -r requirements.txt
```

---

## 💻 Uso

### Inicio Rápido
```powershell
python app.py
```

### Workflow

1. 📁 Selecciona tu archivo Excel con los datos de clientes
2. 👁️ Haz clic en **"Cargar"** para previsualizar los datos
3. 📊 Revisa la grilla paginada con la información
4. 🖨️ Haz clic en **"Generar PDFs"** para crear los documentos
5. ✅ Los PDFs se guardan en la carpeta **"Salidad Impresiones"**



## �📋 Reglas de Negocio Implementadas

| Regla | Descripción |
|-------|-------------|
| **Agrupación** | Por dirección, segmento ejecutivo y titular de pago |
| **Titular de Pago** | Usa `TITULAR PAGO REAL` si existe; si no, usa `TITULAR DE PAGO` |
| **Registros por Página** | Máximo 39 registros por página |
| **Páginas Múltiples** | Los grupos grandes continúan en páginas adicionales |
| **Cálculo de Deuda** | DEUDA = SALDO VENCIDO - IMPORTE |
| **Totalizaciones** | Suma automática de IMPORTE MES, DEUDA y TOTAL |
| **Logo Empresarial** | Se integra desde `Logo/logo-afinia.png` si existe |

---

## 📁 Estructura del Proyecto

```
PrintFactory/
├── app.py                      # Aplicación principal
├── impresiones_app/
│   ├── __init__.py
│   └── core.py                 # Lógica de generación de PDFs
├── Logo/
│   └── logo-afinia.png         # Logo para los documentos
├── Assets/
│   └── Captures/               # Capturas de pantalla
├── requirements.txt            # Dependencias de Python
├── GeneradorClientesEspeciales.spec  # Configuración para PyInstaller
├── build_onefile.bat           # Script para crear ejecutable
└── README.md                   # Este archivo
```

---

## 🛠️ Tecnologías Utilizadas

- **Python 3.x** - Lenguaje de programación
- **CustomTkinter** - Interfaz gráfica moderna
- **OpenPyXL** - Lectura/escritura de archivos Excel
- **ReportLab** - Generación de PDFs
- **PyInstaller** - Empaquetamiento en ejecutable

---

## 📦 Generar Ejecutable

Para crear un archivo `.exe` distribuible sin necesidad de Python instalado:

```powershell
.\build_onefile.bat
```

El ejecutable se generará en la carpeta `dist/`.

---

## 📊 Formato de Entrada (Excel)

El archivo Excel debe contener las siguientes columnas:

| Columna | Descripción |
|---------|-------------|
| NIC | Número de identificación del cliente |
| NUMERO_FACTURA | Número de factura |
| DIRECCION_ENTREGA_REAL | Dirección de entrega |
| SEGMENTO_EJECUTIVO | Segmento del cliente |
| TITULAR_PAGO_REAL | Titular real del pago |
| TITULAR_DE_PAGO | Titular alternativo del pago |
| IMPORTE | Cantidad a pagar |
| SALDO_VENCIDO | Deuda vencida |
| FECHA_VENCIMIENTO | Fecha de vencimiento |

---

## 🎨 Personalización

### Cambiar Logo
Reemplaza `Logo/logo-afinia.png` con tu logo. El programa lo detectará automáticamente.

### Cambiar Directorio de Salida
Usa la interfaz para seleccionar la carpeta donde guardar los PDFs.

### Modificar Texto Legal
Edita la constante `LEGAL_TEXT` en `impresiones_app/core.py`.

---

## 🐛 Solución de Problemas

| Problema | Solución |
|----------|----------|
| Error de módulos | Ejecuta: `pip install -r requirements.txt` |
| Logo no aparece | Verifica que `Logo/logo-afinia.png` exista |
| PDFs no se crean | Asegúrate de que la carpeta de salida tenga permisos de escritura |
| Interfaz lenta | Reduce el tamaño del archivo Excel |

---

## 📝 Ejemplos de Uso

### Ejemplo 1: Facturación Mensual
```
1. Carga el archivo "Facturas_Abril_2025.xlsx"
2. Revisa la previsualización (200 registros encontrados)
3. Genera PDFs (se crean 5 PDFs automáticamente)
4. Los archivos están listos en "Salidad Impresiones/"
```

### Ejemplo 2: Reporte Especial
```
1. Selecciona "Clientes_VIP.xlsx"
2. Configura carpeta de salida a "Reportes/VIP/"
3. Actualiza el logo si es necesario
4. Genera los documentos
```

---

## 📄 Licencia

Este proyecto es de uso interno. Todos los derechos reservados.



## 🤝 Contribuciones

Para reportar bugs o sugerir mejoras:
1. Abre un **Issue** detallando el problema
2. Proporciona pasos para reproducirlo
3. Incluye screenshots si es relevante

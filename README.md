# Generador de CSV para ArcGIS Survey123 🗺️📊  

## 📖 Descripción  
Este script de **MATLAB** automatiza la conversión y limpieza de datos desde archivos **Excel** (`.xlsx`) a **CSV**, generando archivos listos para ser utilizados en un **formulario de ArcGIS Survey123**.  
Permite extraer, transformar y estructurar datos clave de infraestructura eléctrica, asegurando calidad y consistencia en la información.  

## 🚀 Características  
✅ **Procesamiento automático** de distintas hojas del archivo Excel (*Transformadores, Postes y Medidores*).  
✅ **Limpieza de datos**, eliminando registros vacíos o inválidos.  
✅ **Estandarización de códigos de alimentador** mediante expresiones regulares.  
✅ **Eliminación de registros duplicados**, garantizando datos únicos y sin redundancias.  
✅ **Generación automática de archivos CSV** organizados para su uso en ArcGIS Survey123.  

## 📂 Estructura del Script  

### 🔹 Entrada  
- Archivo Excel con hojas específicas: `TRAFO`, `POSTE`, `MEDIDORES`.  
- Columnas clave como `ALIMENTADO`, `CODIGOPUES`, `OID_POSTE`, `MDENUMEMP`.  

### 🔹 Procesamiento  
- Eliminación de valores `0` o vacíos en las columnas principales.  
- Uso de expresiones regulares para extraer códigos y etiquetas relevantes.  
- Eliminación de registros duplicados en función de identificadores clave.  
- Conversión y exportación de datos en formato CSV.  

### 🔹 Salida  
- Archivos CSV organizados en una carpeta ingresada por el usuario (`folderOut = '59ECSV';`).  
- Datos estructurados y optimizados para ArcGIS Survey123.  

## 🛠 Requisitos  
- **MATLAB** (versión 2021 o superior recomendada).  
- **Archivo Excel** con datos estructurados segun lo requiera el formulario de Survey123.  
- **Extensión `writecell()`** disponible en la versión de MATLAB utilizada.  

## 📌 Cómo Usarlo  
1️⃣ Ubica el archivo Excel que sera la fuente para los archivos.csv (`TABLA GENERAL 59E.xlsx`) en la misma carpeta del script.  
2️⃣ Ejecuta el script en **MATLAB** o **VIASUAL STUDIO CODE** .  
3️⃣ Espera la confirmación en la consola de MATLAB.  
4️⃣ Revisa la carpeta de salida para encontrar los archivos CSV generados.  

## 👨‍💻 Desarrollador  
✏️ **Juan Ortiz** - *Consorcio SIG-Electric (2025)*  

Si tienes dudas o sugerencias para mejorar este script, ¡no dudes en contribuir o abrir un **issue** en este repositorio! 🚀📌  


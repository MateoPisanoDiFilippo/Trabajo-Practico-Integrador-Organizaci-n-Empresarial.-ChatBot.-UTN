# Trabajo-Practico-Integrador-Organizaci-n-Empresarial.-ChatBot.-UTN
Código de Chatbot para el Trabajo practico integrador de la materia Organización Empresarial, 
primer año del primer cuatrimestre de la carrera Tecnicatura en Programación de la UTN


## MURMAT S.A. — Sistema de Gestión de Proveedores

Sistema de gestión de proveedores desarrollado en Python con backend en Excel (.xlsx).

## Requisitos

- Python 3.10 o superior
- Librería `openpyxl` (`pip install openpyxl`)

## Cómo ejecutar

1. Clonar el repositorio
2. Asegurarse de que el archivo `proveedores_murmat.xlsx` esté en la misma carpeta que el script
3. Ejecutar: `python murmat.py`

## Funcionalidades

- **Alta de proveedor**: registra un nuevo proveedor con validación de CUIT, email y teléfono
- **Baja de proveedor**: elimina un proveedor del sistema previa confirmación
- **Cambiar estado**: alterna el estado del proveedor entre activo e inactivo
- **Ver lista completa**: muestra todos los proveedores registrados
- **Buscar proveedor**: permite buscar por nombre o CUIT
- **Modificar datos**: edita los datos de un proveedor existente

## Búsqueda

En todas las operaciones que requieren identificar un proveedor, se puede ingresar su **CUIT** (con o sin guiones) o su **nombre/razón social**. En cualquier momento se puede escribir `salir` para cancelar la operación.

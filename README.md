# Conversor de Excel a LoRaWAN JSON (LNS)

Este proyecto provee una interfaz gráfica (GUI) en Python escrita con **Tkinter** para importar un archivo de Excel (`.xlsx`), mapear dinámicamente sus columnas, y estructurar toda la información en un formato JSON listo para ser utilizado en servidores de red LoRaWAN (LNS).

El proyecto fue diseñado de manera **autónoma y sin dependencias externas** (zero-dependency). Utiliza los módulos estándar `zipfile` y `xml.etree` de Python para procesar los archivos Excel de manera nativa, evitando cualquier problema de conexión de red o instalación de librerías mediante `pip`.

---

## Características Principales

1. **Sin dependencias externas**: Funciona en cualquier sistema que tenga Python 3 instalado, sin necesidad de correr `pip install`.
2. **Carga Automática de Columnas**: Al seleccionar tu archivo Excel, la interfaz lee la primera fila y lista todas las columnas detectadas.
3. **Composición de Fórmulas Dinámicas**:
   - Puedes componer libremente los campos `Device ID`, `Name` y `Description`.
   - Al hacer **doble clic** sobre cualquier columna del listado izquierdo, se insertará la variable en formato `{NombreColumna}` en la posición actual de tu cursor del campo seleccionado.
   - El campo `Device ID` convierte automáticamente todas las letras a minúsculas al exportar.
4. **Filtro de Rango de Filas**:
   - Puedes exportar todo el archivo completo (desde la fila 2 en adelante).
   - O bien, desmarcar "Todo el archivo" y definir un rango personalizado de filas a procesar (ej. Desde fila `2` Hasta fila `50`).
5. **Persistencia Automática de Configuración**:
   - Cuenta con un archivo `config.ini` local que almacena los valores actuales y fórmulas predeterminadas.
   - Al modificar cualquier campo o al cerrar el programa (usando la "X"), los cambios se guardan de forma automática para tu siguiente ejecución.

---

## Formato del JSON de Salida

Cada dispositivo procesado se exportará en una lista de objetos siguiendo la siguiente estructura:

```json
{
  "ids": {
    "device_id": "dispositivo-my-device-01",
    "dev_eui": "AC1F09FFFE24F82C",
    "join_eui": "6C4EEF66F47986A6"
  },
  "name": "My Device 01",
  "description": "Living room temperature sensor",
  "lorawan_version": "MAC_V1_0_2",
  "lorawan_phy_version": "PHY_V1_0_2_REV_B",
  "frequency_plan_id": "AU_915_928_FSB_2",
  "supports_join": true,
  "root_keys": {
    "app_key": {
      "key": "1F33A170A5F1FDA0AB697AAE2B95916B"
    }
  }
}
```

---

## Cómo Ejecutar el Proyecto

Solo requieres contar con Python 3 instalado. Desde la terminal del sistema en el directorio del proyecto, ejecuta:

```powershell
python excel_to_json_gui.py
```

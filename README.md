# Sistema de Inventario Web

Dashboard web para gestión de inventario de computadoras, impresoras y notebooks con soporte para bases de datos compartidas en red y fallback local.

## 📋 Características

- **Gestión de Inventario**: Control completo de computadoras, impresoras y notebooks
- **Multi-Base de Datos**: Soporte para bases de datos compartidas en red (\\16.1.1.118) con fallback automático a bases de datos locales
- **Exportación**: Exportación de datos a Excel (XLSX) con formato profesional
- **Estados**: Seguimiento de estados de equipos (Ok, En reparación, Dado de baja, etc.)
- **Búsqueda Avanzada**: Búsqueda por MAC address, número de serie, usuario, oficina y más
- **Historial**: Registro completo de asignaciones y devoluciones de notebooks
- **API REST**: Endpoints para integración con otros sistemas

## 🚀 Inicio Rápido

### Prerequisitos

- Docker y Docker Compose instalados
- Acceso a la red compartida 16.1.1.118 (opcional, funciona con BD local)

### Instalación con Docker

1. **Clonar o descargar el proyecto**
```bash
git clone <tu-repositorio>
cd inventario-web
```

2. **Construir y levantar el contenedor**
```bash
docker-compose up -d
```

3. **Acceder a la aplicación**
```
http://localhost:5000
```

### Instalación Manual (sin Docker)

1. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

2. **Ejecutar la aplicación**
```bash
python app_web.py
```

3. **Acceder a la aplicación**
```
http://localhost:5000
```

Para ejecutar en un puerto diferente:
```bash
python app_web.py 8080
```

## 📁 Estructura del Proyecto

```
inventario-web/
├── app_web.py              # Aplicación Flask principal
├── requirements.txt        # Dependencias Python
├── Dockerfile             # Configuración Docker
├── docker-compose.yml     # Orquestación de contenedores
├── README.md             # Este archivo
└── db/                   # Bases de datos locales (fallback)
    ├── inventario_local_fallback.db
    ├── impresoras_local_fallback.db
    └── notebooks_local_fallback.db
```

## 🗄️ Bases de Datos

El sistema utiliza múltiples bases de datos SQLite:

### Bases de Datos Compartidas (Red)
- **Inventario**: `\\16.1.1.118\db\inventario_computadoras.db`
- **Oficinas**: `\\16.1.1.118\db\oficinasCne.db`
- **Impresoras**: `\\16.1.1.118\db\impresoras.db`
- **Notebooks**: `\\16.1.1.118\db\inventario_notebooks.db`

### Bases de Datos Locales (Fallback)
Si no se puede acceder a las bases de datos compartidas, el sistema automáticamente usa:
- `inventario_local_fallback.db`
- `impresoras_local_fallback.db`
- `notebooks_local_fallback.db`

## 🔧 Configuración

### Variables de Entorno

Puedes modificar las siguientes variables en `docker-compose.yml`:

```yaml
environment:
  - FLASK_ENV=production  # Cambiar a 'development' para modo debug
  - PYTHONUNBUFFERED=1
```

### Montaje de Red Compartida

Para acceder a las bases de datos en red desde Docker, descomentar y ajustar en `docker-compose.yml`:

**Windows con WSL:**
```yaml
volumes:
  - //16.1.1.118/db:/mnt/shared-db
```

**Linux:**
```yaml
volumes:
  - /mnt/16.1.1.118/db:/mnt/shared-db
```

### Puerto de Ejecución

Modificar en `docker-compose.yml`:
```yaml
ports:
  - "8080:5000"  # Cambiar 8080 por el puerto deseado
```

## 📊 Funcionalidades

### Inventario de Computadoras
- ✅ Agregar/editar/eliminar equipos
- ✅ Campos: Usuario, Oficina, PC Nombre, IP, MAC, Marca, Modelo, Estado
- ✅ Información de hardware: Procesador, RAM, Disco, Motherboard, Tarjeta Gráfica
- ✅ Búsqueda y filtrado avanzado
- ✅ Exportación a Excel

### Impresoras
- ✅ Gestión completa de impresoras
- ✅ Control de tóner y mantenimiento
- ✅ Estados y ubicaciones
- ✅ Exportación a Excel

### Notebooks
- ✅ Control de préstamos y devoluciones
- ✅ Historial completo de asignaciones
- ✅ Personas asignadas y fechas
- ✅ Estados y observaciones
- ✅ Exportación con historial

## 🌐 API Endpoints

### Oficinas
- `GET /api/oficina/<nombre>/piso` - Obtener piso de una oficina

### MAC Address
- `GET /api/buscar_mac/<mac_address>` - Buscar equipos por MAC
- `GET /api/macs_registradas` - Listar todas las MACs registradas

### Debug
- `GET /debug_bd` - Ver estructura de tabla inventario
- `GET /debug_oficinas` - Ver oficinas registradas

## 🐳 Comandos Docker Útiles

```bash
# Ver logs en tiempo real
docker-compose logs -f

# Detener contenedores
docker-compose down

# Reconstruir después de cambios
docker-compose up -d --build

# Acceder al contenedor
docker exec -it inventario-web-app bash

# Ver estado de contenedores
docker-compose ps
```

## 🔒 Seguridad

⚠️ **IMPORTANTE**: El `secret_key` actual es un placeholder. Para producción:

1. Generar una clave segura:
```python
import secrets
print(secrets.token_hex(32))
```

2. Actualizar en `app_web.py`:
```python
app.secret_key = 'tu_clave_generada_aqui'
```

## 🛠️ Desarrollo

### Modo Debug

Para ejecutar en modo desarrollo con recarga automática:

```python
# En app_web.py, cambiar la última línea a:
app.run(host="16.1.1.118", port=port, debug=True)
```

### Agregar Nuevas Funcionalidades

1. Las rutas Flask están organizadas por sección (inventario, impresoras, notebooks)
2. Los métodos de base de datos están en la clase `UnifiedWebDatabaseManager`
3. Seguir la estructura existente para mantener consistencia

## 📝 Notas

- El sistema automáticamente crea las bases de datos locales si no existen
- Las migraciones de esquema se ejecutan automáticamente al iniciar
- Los archivos Excel exportados incluyen formato profesional con colores y estilos
- El sistema detecta automáticamente si está usando BD compartida o local

## 🐛 Solución de Problemas

### Error de conexión a BD compartida
- Verificar acceso de red a 16.1.1.118
- El sistema automáticamente usará BD local como fallback

### Puerto ocupado
- Cambiar el puerto en `docker-compose.yml` o al ejecutar manualmente

### Permisos en volúmenes Docker
```bash
sudo chown -R $USER:$USER ./db
```

## 📄 Licencia

Este proyecto es de uso interno.

## 👥 Soporte

Para reportar problemas o sugerencias, contactar al equipo de desarrollo.

---

**Versión**: 1.0  
**Última actualización**: 2025

# 🖥️ Sistema de Inventario Web

Dashboard web para la gestión de inventario de equipos informáticos desarrollado durante mi trabajo en la **Cámara Nacional Electoral**. Nació de la necesidad de tener un sistema centralizado, accesible desde cualquier PC de la red, que reemplazara el seguimiento manual de equipos.

> Proyecto desarrollado por iniciativa propia en el entorno laboral.

---

## 💡 Origen del proyecto

El inventario de equipos se manejaba de forma manual. Desarrollé este sistema para centralizarlo todo: búsquedas rápidas, exportación con un click y seguimiento de estado de cada equipo. Lo que antes requería revisar planillas, pasó a estar disponible desde cualquier navegador de la red.

---

## ✨ Funcionalidades

### 🖥️ Inventario de Computadoras
- Alta, edición y baja de equipos con datos completos: usuario, oficina, IP, MAC, marca, modelo, procesador, RAM, disco, motherboard y tarjeta gráfica.
- Búsqueda avanzada por MAC, número de serie, usuario, oficina y más.
- Seguimiento de estados: Ok, En reparación, Dado de baja, etc.
- Exportación a **Excel (XLSX)** con formato profesional.

### 🖨️ Impresoras
- Gestión completa con control de tóner, mantenimiento, estados y ubicaciones.
- Exportación a Excel.

### 💻 Notebooks
- Control de préstamos y devoluciones con historial completo de asignaciones.
- Registro de personas asignadas, fechas, estados y observaciones.
- Exportación con historial incluido.

### 🌐 Multi-Base de Datos
- Conexión a bases de datos compartidas en red.
- **Fallback automático** a bases de datos locales si la red no está disponible — el sistema nunca se cae.

### 🔌 API REST
- Endpoints para integración con otros sistemas internos.
- Búsqueda por MAC address y listado de MACs registradas.

---

## 🛠️ Stack

![Python](https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white)
![Flask](https://img.shields.io/badge/Flask-000000?style=for-the-badge&logo=flask&logoColor=white)
![SQLite](https://img.shields.io/badge/SQLite-003B57?style=for-the-badge&logo=sqlite&logoColor=white)
![HTML5](https://img.shields.io/badge/HTML5-E34F26?style=for-the-badge&logo=html5&logoColor=white)
![CSS3](https://img.shields.io/badge/CSS3-1572B6?style=for-the-badge&logo=css3&logoColor=white)
![Docker](https://img.shields.io/badge/Docker-2496ED?style=for-the-badge&logo=docker&logoColor=white)

---

## 🚀 Instalación

### Con Docker (recomendado)

```bash
git clone https://github.com/TomasMorenoo/InventarioWeb
cd InventarioWeb
docker-compose up -d
```

Accedé en: `http://localhost:5000`

### Sin Docker

```bash
pip install -r requirements.txt
python app_web.py
```

Para un puerto personalizado:

```bash
python app_web.py 8080
```

---

## ⚙️ Configuración

Las variables de entorno se configuran en `docker-compose.yml`:

```yaml
environment:
  - FLASK_ENV=production
  - PYTHONUNBUFFERED=1
```

El puerto se puede cambiar desde el mismo archivo:

```yaml
ports:
  - "8080:5000"
```

> ⚠️ Antes de pasar a producción, reemplazá el `secret_key` por uno generado de forma segura:
> ```python
> import secrets
> print(secrets.token_hex(32))
> ```

---

## 📁 Estructura

```
inventario-web/
├── app_web.py          # Aplicación Flask principal
├── config.py           # Configuración
├── requirements.txt    # Dependencias
├── Dockerfile
├── docker-compose.yml
└── db/                 # Bases de datos locales (fallback)
```

---

## 🐳 Comandos Docker útiles

```bash
docker-compose logs -f          # Ver logs en tiempo real
docker-compose down             # Detener
docker-compose up -d --build    # Reconstruir tras cambios
docker exec -it inventario-web-app bash  # Acceder al contenedor
```

---

## 👨‍💻 Autor

**Tomás Moreno Bauer**
- 🌐 [portfolio.mobatai.com](https://portfolio.mobatai.com)
- 📧 [morenobauer10@gmail.com](mailto:morenobauer10@gmail.com)
- 💬 [+54 11 3188-1483](https://wa.me/5491131881483)
- 💼 [linkedin.com/in/tomas-moreno-bauer](https://linkedin.com/in/tomas-moreno-bauer)

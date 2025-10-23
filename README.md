🚌 Sistema de Revisión de Buses
Sistema web para gestionar y realizar seguimiento de revisiones técnicas de buses, incluyendo conectividad, norma gráfica y estado de disco duro.
📋 Características

✅ Registro de revisiones por PPU
📊 Dashboard con estadísticas en tiempo real
📈 Gráficos interactivos con Chart.js
📥 Exportación a Excel con formato profesional
🔍 Búsqueda inteligente de PPUs
📱 Diseño responsive (móvil y escritorio)

🛠️ Tecnologías

Backend: Flask (Python)
Base de Datos: Supabase (PostgreSQL)
Frontend: HTML5, CSS3, JavaScript
Gráficos: Chart.js
Exportación: openpyxl

📦 Instalación Local

Clonar el repositorio:

bashgit clone https://github.com/TU_USUARIO/revision-buses.git
cd revision-buses

Crear entorno virtual:

bashpython -m venv venv

# En Windows:
venv\Scripts\activate

# En Linux/Mac:
source venv/bin/activate

Instalar dependencias:

bashpip install -r requirements.txt

Configurar variables de entorno:

Crea un archivo .env en la raíz del proyecto:
envSUPABASE_URL=tu_url_de_supabase
SUPABASE_KEY=tu_key_de_supabase

Ejecutar la aplicación:

bashpython app.py
La aplicación estará disponible en http://localhost:5000
🗄️ Estructura de la Base de Datos
Tabla revisiones:
sqlCREATE TABLE revisiones (
    id SERIAL PRIMARY KEY,
    ppu VARCHAR(10) NOT NULL,
    fecha TIMESTAMP NOT NULL,
    conectividad BOOLEAN NOT NULL,
    motivo_no_conectividad TEXT,
    norma_grafica_correcta BOOLEAN NOT NULL,
    disco_duro BOOLEAN NOT NULL,
    created_at TIMESTAMP DEFAULT NOW()
);
🚀 Despliegue
Opción 1: Render.com (Recomendado - Gratis)

Crea una cuenta en Render.com
Conecta tu repositorio de GitHub
Crea un nuevo "Web Service"
Configura las variables de entorno
Deploy automático

Opción 2: Railway.app

Crea una cuenta en Railway.app
Conecta tu repositorio
Configura variables de entorno
Deploy automático

Opción 3: Heroku

Instala Heroku CLI
Ejecuta:

bashheroku login
heroku create nombre-app
git push heroku main
heroku config:set SUPABASE_URL=tu_url
heroku config:set SUPABASE_KEY=tu_key
📱 Uso
Crear Nueva Revisión

Busca la PPU del bus
Selecciona la fecha y hora
Completa el formulario de revisión
Guarda

Ver Dashboard

Accede a /dashboard
Visualiza estadísticas en tiempo real
Exporta datos a Excel

🔒 Seguridad

Las credenciales de Supabase están en variables de entorno
No subas el archivo .env a GitHub
Usa la clave anon de Supabase (no la service key)

📄 Licencia
Este proyecto es de uso privado.
👨‍💻 Autor
Desarrollado para gestión de flota de buses.
🆘 Soporte
Para problemas o sugerencias, crea un issue en GitHub.
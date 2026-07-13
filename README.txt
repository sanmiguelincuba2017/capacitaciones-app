============================================================
SISTEMA DE GESTIÓN DE CAPACITACIONES
============================================================

Aplicación web local para procesar capacitaciones (webinars, 
cursos, seminarios y ferias) de forma automática.

¿QUÉ HACE?
-----------
• Sube los Excel de inscriptos y matriculados/certificados
• Procesa automáticamente: extrae CPs, localidades, clasifica 
  por partido, cruza matriculados con certificados
• Genera Excel de 4 hojas (Inscriptos, Matriculados, 
  Certificados, Consolidado)
• Genera Informe en Word con gráficos y análisis descriptivo
• Guarda historial de todas las capacitaciones procesadas

REQUISITOS
-----------
1. Tener Python instalado (versión 3.8 o superior)
   Descargar desde: https://www.python.org/downloads/

2. Durante la instalación de Python, MARCAR la casilla 
   "Add Python to PATH"

INSTALACIÓN
------------
1. Abrir la terminal (CMD en Windows, Terminal en Mac/Linux)

2. Navegar a esta carpeta:
   cd ruta/donde/descomprimiste/esta/carpeta

3. Instalar dependencias:
   pip install -r requirements.txt

EJECUCIÓN
----------
1. En la terminal, dentro de esta carpeta, ejecutar:
   python app.py

2. Abrir el navegador y entrar a:
   http://127.0.0.1:5000

3. Ingresar con:
   Usuario: admin
   Contraseña: admin123

USO
----
1. Click en "Procesar Nueva Capacitación"
2. Completar nombre y números (inscriptos, matriculados, certificados)
3. Subir archivo de inscriptos (respuesta del formulario Google)
4. Subir archivo de matriculados/certificados
5. Click en "Procesar Capacitación"
6. Descargar Excel y Word generados

NOTAS
------
• Los archivos generados se guardan en la carpeta "generated"
• El historial se guarda en "data/capacitaciones.json"
• Para cambiar la contraseña, editar app.py (línea USUARIO)

NUEVO: INFORMES POR PERÍODO
-----------------------------
Desde el menú "Informes por Período" podés generar un Excel + Word
consolidado de todas las capacitaciones cargadas en un trimestre,
semestre o año. Se arma automáticamente a partir de la "Fecha de
realización" que cargues en cada capacitación nueva.

DESPLIEGUE EN LA NUBE (sin instalar Python) - RENDER
------------------------------------------------------
1. Crear cuenta en github.com (podés usar "Continue with Google")
2. Crear un repositorio nuevo y subir todos los archivos de esta
   carpeta (Add file → Upload files)
3. Crear cuenta en render.com (podés entrar con la misma cuenta
   de GitHub/Google)
4. New → Web Service → elegir el repositorio subido
5. Render detecta Python automáticamente. Confirmar:
   - Build Command: pip install -r requirements.txt
   - Start Command: gunicorn app:app
6. Esperar 2-3 minutos. Render entrega un link tipo
   https://tuapp.onrender.com para usar desde cualquier PC o celular.

NOTA: en el plan gratuito de Render, los archivos generados NO son
permanentes (se borran si el servicio se reinicia). Para uso
institucional serio conviene descargar cada informe apenas se genera,
o pasar a un plan pago con disco persistente.

SOPORTE
--------
Si hay algún problema, contactar al desarrollador.
============================================================
# midnightblue
# Proyecto de Automatización de Envío de Correos Masivos

Este proyecto se enfoca en la automatización de envío de correos masivos personalizados utilizando Python, ttkbootstrap, SMTP, Excel, HTML y CSS.

## Descripción
**Programación:**

1. **Elementos fundamentales:**
   - Variables: Uso extensivo de StringVar para la interfaz gráfica (smtp_var, user_var, pass_var, etc.)
   - Tipos de datos: strings, listas, diccionarios, booleanos
   - Operadores: aritméticos (+, -), comparación (<=, >), lógicos (and, or)
   - Constantes: CONFIG_FILE = "config.json"

2. **Estructuras de control:**
   - Selección: if/else para validaciones y manejo de errores
   - Repetición: for para iterar sobre filas de Excel
   - Funciones: Múltiples definiciones de funciones (load_config, save_config, etc.)

3. **Control de excepciones:**
   - Implementación robusta de try/except
   - Manejo específico de FileNotFoundError
   - Gestión de errores SMTP y de archivos
   - Mensajes de error personalizados mediante messagebox

4. **Documentación:**
   - Comentarios explicativos para secciones importantes
   - Comentarios para explicar la personalización del mensaje
   - Falta documentación más detallada y docstrings

5. **Paradigma:**
   - Principalmente programación estructurada
   - Uso de elementos de programación orientada a eventos para la interfaz gráfica
   - No implementa clases propias, pero utiliza objetos de las bibliotecas

6. **Clases y objetos principales:**
   - Utiliza objetos de tkinter/ttkbootstrap (root, Button, Entry)
   - MIMEMultipart para construcción de correos
   - Template para personalización de contenido HTML

7. **Conceptos avanzados:**
   - No implementa herencia propia
   - Utiliza polimorfismo a través de las interfaces de tkinter
   - Implementa callbacks y eventos

8. **Gestión de información:**
   - Archivos: HTML, Excel, CSV para reportes
   - Interfaz gráfica completa con ttkbootstrap
   - Manejo de configuración en JSON

9. **Estructuras de datos:**
   - Listas para templates y estadísticas
   - Diccionarios para configuración y personalización
   - Matrices a través de Excel (openpyxl)

10. **Técnicas avanzadas:**
    - Manejo de hilos para operaciones asíncronas
    - Flujos de entrada/salida para archivos
    - Plantillas HTML con string.Template

**Sistemas Informáticos:**

11. **Hardware y entorno:**
    - Desarrollo en Windows (evidenciado por las rutas)
    - Servidor XAMPP (ubicación en c:\xampp\htdocs)

12. **Sistema operativo:**
    - Windows para desarrollo
    - Compatible con múltiples plataformas

13. **Configuración de redes:**
    - Utiliza SMTP con TLS (puerto 587)
    - Implementa conexiones seguras (starttls)

14. **Copias de seguridad:**
    - Guarda plantillas en directorio "templates"
    - Genera reportes de envíos en CSV

15. **Seguridad de datos:**
    - Encriptación TLS para SMTP
    - Almacenamiento de credenciales en config.json
    - Validación de entradas de usuario

16. **Usuarios y permisos:**
    - No implementa sistema de usuarios propio
    - Utiliza credenciales SMTP para autenticación

17. **Documentación técnica:**
    - Estructura de archivos clara
    - Falta documentación técnica formal

**Entornos de Desarrollo:**

18. **IDE:**
    - No se especifica, pero compatible con cualquier IDE Python

19. **Automatización:**
    - Generación automática de reportes
    - Programación de envíos

20. **Control de versiones:**
    - Usa Git (.gitignore presente)
    - Ignora archivos sensibles y temporales

21. **Refactorización:**
    - Funciones modulares y bien definidas
    - Separación de lógica de negocio e interfaz

22. **Documentación técnica:**
    - Comentarios en código
    - Estructura de archivos organizada

23. **Diagramas:**
    - No se evidencian diagramas en el proyecto

**Bases de Datos:**
- No utiliza base de datos, usa archivos Excel y CSV para datos

**Lenguajes de Marcas:**

24. **HTML:**
    - Utiliza plantillas HTML para correos
    - Implementa estilos CSS inline
    - Estructura semántica con tablas

25. **Frontend:**
    - Interfaz gráfica con ttkbootstrap
    - Estilos CSS en plantillas HTML

26. **JavaScript:**
    - No utiliza JavaScript directamente

27. **Validación:**
    - Validación de campos obligatorios
    - Verificación de fechas y formatos

28. **Conversión de datos:**
    - Conversión entre Excel y CSV
    - Personalización de plantillas HTML

29. **Gestión empresarial:**
    - Es una aplicación de gestión de marketing por email
    - Permite automatización de envíos masivos
    - Genera reportes de seguimiento

**Proyecto intermodular:**

30. **Objetivo:**
    - Automatización de envío de correos masivos personalizados

31. **Necesidad:**
    - Facilitar campañas de email marketing
    - Seguimiento de envíos
    - Personalización de contenido

32. **Stack tecnológico:**
    - Python como lenguaje principal
    - ttkbootstrap para interfaz gráfica
    - SMTP para envío de correos
    - Excel/CSV para datos
    - HTML/CSS para plantillas

33. **Desarrollo por versiones:**
    - Funcionalidad básica de envío de correos
    - Características adicionales:
      - Programación de envíos
      - Gestión de plantillas
      - Generación de reportes
      - Personalización de contenido

        
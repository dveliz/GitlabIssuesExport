# GitLab Issues Exporter

Una aplicación de consola en Python para exportar issues de GitLab a archivos Excel.

## Instalación

1. Instalar las dependencias:
```bash
pip install -r requirements.txt
```

2. Configurar el archivo `config.json`:
```json
{
    "gitlab_token": "tu_token_de_gitlab_aqui",
    "project_id": "id_del_proyecto",
    "gitlab_url": "https://gitlab.com"
}
```

## Configuración

### Obtener el Token de GitLab
1. Ve a GitLab → User Settings → Access Tokens
2. Crea un token con permisos de `read_api`
3. Copia el token y pégalo en `config.json`

### Obtener el ID del Proyecto
1. Ve a tu proyecto en GitLab
2. El ID del proyecto aparece debajo del nombre del proyecto
3. Usa este ID en `config.json`

## Uso

Ejecuta la aplicación:
```bash
python main.py
```

La aplicación te pedirá:
- Fecha de inicio (opcional): YYYY-MM-DD
- Fecha de fin (opcional): YYYY-MM-DD

## Formato del Excel

El archivo Excel generado incluye las siguientes columnas:
- ID del issue
- Título del issue
- Descripción del issue
- Nombre del autor
- Estado del issue
- Asignados al issue (separados por coma)
- Etiquetas del issue (separados por coma)
- Fecha y hora de creación del issue
- Tiempo total estimado del issue
- Tiempo total gastado del issue
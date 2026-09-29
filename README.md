
# Buscador Laboral – Paulina Vergara

Aplicación Streamlit para centralizar búsquedas laborales orientadas a:
- Procesos
- Mejora Continua
- Excelencia Operacional
- Experiencia Cliente / CX
- Operaciones
- PMO / Proyectos
- Control de Gestión
- Power BI / KPIs

## 1. Probar localmente

Instala Python 3.11 o superior.

En una terminal:

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá una página local en el navegador.

## 2. Publicar gratis en Streamlit Community Cloud

1. Crea una cuenta en GitHub.
2. Crea un repositorio nuevo, por ejemplo:
   `buscador-empleos-paulina`
3. Sube estos archivos:
   - app.py
   - requirements.txt
   - .streamlit/secrets.toml.example
4. Entra a https://share.streamlit.io/
5. Conecta tu cuenta de GitHub.
6. Selecciona el repositorio.
7. Archivo principal: `app.py`
8. Pulsa Deploy.

La aplicación quedará con una URL similar a:
`https://buscador-empleos-paulina.streamlit.app`

## 3. Activar búsqueda integrada dentro de la app

La aplicación funciona sin API, pero en ese modo entrega enlaces de búsqueda directa a:
- LinkedIn
- Computrabajo
- Trabajando

Para mostrar resultados integrados en una sola pantalla utiliza SerpAPI Google Jobs.

### Configuración

1. Crea una cuenta en SerpAPI.
2. Copia tu API key.
3. En Streamlit Cloud abre:
   `App > Settings > Secrets`
4. Agrega:

```toml
SERPAPI_KEY = "TU_CLAVE"
```

5. Guarda y reinicia la app.

No publiques tu API key en GitHub.

## 4. Uso diario

1. Abre la URL de tu app.
2. Pulsa **BUSCAR OFERTAS DE HOY**.
3. Revisa primero resultados de 80% o más.
4. Pulsa **Abrir oferta**.
5. Cambia el estado a:
   - Guardada
   - Postulada
   - Descartada
6. Al día siguiente vuelve a ejecutar la búsqueda.

## Importante sobre sueldos

La app NO inventa remuneraciones.
Si una publicación no informa sueldo, aparecerá:
`No publicada`.

Esto evita descartar o priorizar una oferta basándose en cifras no verificadas.

## Persistencia

La versión incluida guarda estados en SQLite (`jobs.db`).
En algunos despliegues gratuitos de Streamlit el almacenamiento puede reiniciarse cuando la app se reconstruye.
Para persistencia permanente, una siguiente versión puede conectarse a:
- Supabase
- Google Sheets
- Airtable

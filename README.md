## Dartev - Backend de reportes

### Requisitos

- Python3.6

- Variables de ambiente.

### Instalación

Instalar virtualenv:


```bash
# Install virtualenv
sudo pip3.6 install virtualenv
# Create virtualenv enviroment
python3.6 -m venv venv
```

Activar ambiente virtual:

```bash
# Activate virtualenv
source venv/bin/activate
# Install project requirements
pip3.6 install -r requirements.txt
```

Cargar variables de ambiente

```bash
source venv
```

Arrancar proyecto:

```bash
python3.6 reports.py
```

### Resolucion de problemas

Problemas con psycopg2-binary:

```
pip install psycopg2-binary
```

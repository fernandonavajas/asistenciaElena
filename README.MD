# Crear virtual enviroment  
python3 -m venv lena  

# Activar enviroment 
source lena/bin/activate  

# Congelar librerias  
pip freeze > requirements.txt  

# Instalar librerias  
pip install -r requirements.txt  

# Iniciar proyecto
python manage.py run

# Crear container
docker build -t flask_base:1 .          

# Ejecutar pytest
python -m pytest        

# Ejecutar test coverage
python -m pytest --cov="project"

# Ejecutar safety
safety check

# Ejecutar docker
docker container run -it -p 5001:5000  flask_base:1 

# Ejecutar pytest desde docker
docker exec -it def738691234 pytest "project/tests"

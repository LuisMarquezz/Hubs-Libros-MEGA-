# Hub Libros-Visual-Basic-6
## Descripción
El proyecto es una aplicacion para gestionar los libros favoritos, libros que no le gustaron al usuario, ver libros entre otras funciones.

## Objetivos 
1.- Aprender la tecnologia visual basic 6

2.- La aplicación se realiza de tal manera que permita a los usuarios gestionar su colección de libros, marcandolos como leidos, favoritos o libros que no les gustan

3.- Proporcionar una interfaz intuitiva, con la creacion de menus y con la informacion organizada es de facil interazción con un catalogo de libros donde tambien el usuario puede buscar a traves de la barra de busqueda por si le interesa algo muy especifico
## Funcionamiento
 El programa comienza con un login siendo necesario que el usuario cree una cuenta.
 ![image](https://github.com/LuisMarquezz/Hubs-Libros-MEGA-/blob/master/Hub%20Libros%20Imagenes/Imagen_Login.jpg?raw=true)
 ![image](https://github.com/LuisMarquezz/Hubs-Libros-MEGA-/blob/master/Hub%20Libros%20Imagenes/Imagen_register.jpg?raw=true)

 Despues de eso y de logearse aparecera en el menu principal donde podra intactuar con la aplicacion.
![image](https://github.com/LuisMarquezz/Hubs-Libros-MEGA-/blob/master/Hub%20Libros%20Imagenes/Imagen_Home.jpg?raw=true)
![image](https://github.com/LuisMarquezz/Hubs-Libros-MEGA-/blob/master/Hub%20Libros%20Imagenes/Imagen_Libro.jpg?raw=true)
## Instrucciones de uso
1.- Debes descargar los forms y despues abrirlos en el proyecto es necesario cambiar las contraseña de las BD para que funcione

2.- Asegurate de tener instalado las configuraciones de entorno como lo son: visual basic 6 y SQL server

3.- Crear la base da datos, en los archivos de la carpeta Modules puedes realizar las configuraciones necesarias para la DB y ajustarlas en el codigo.

## Descripción de elaboración
El proyecto es una aplicación de gestión y visualización de libros, desarrollada en Visual Basic 6 y respaldada por una base de datos en SQL Server. La aplicación permite a los usuarios gestionar sus preferencias de lectura y visualizar información sobre libros.
La aplicacion consume la API Books de Google por lo que tiene acceso a una amplia variedad de libros.
2.- Diseño de la base de datos
![image](https://github.com/LuisMarquezz/Hubs-Libros-MEGA-/blob/master/Hub%20Libros%20Imagenes/Diagrama_Entidad_Relacion.jpg?raw=true)
Se diseñó el esquema de la base de datos utilizando un modelo entidad-relación (ERD), que incluye cuatro tablas principales:

-Usario:

Almacena información del usuario, como nombre, apellido, contraseña etc.

-Libros favoritos:

Guarda los libros favoritos del usuario.

-Libros leidos:

Guarda los libros del usuario.

-Libros por leer

Guarda los libros por leer del usuario

-Libros que no te gustaron

Guarda los libros que no le gustaron al usuario.

## Desarrollo de la Aplicacion
Se crearon varios formularios para diferentes funcionalidades, ademas de modulos para verificar usuarios, conexion a la bd y lectura de JSON.
La aplicacion se penso para usar la API de Google sin tener que guarda informacion de libros solamente el link de referencia

## Problemas conocidos
-La planeacion del proyecto es fundamental ya que ayuda a que tengamos un rumbo, muchas veces dude al saber como hacer las cosas y lo que queria lograr.

-Fue un gran reto realizar el proyecto ya que con tecnologia que nunca habia visto es algo complicado sumandole el hecho de que hay poca informacion sobre el lenguaje

-El programa tiene problemas al hacer algunas consultas pues no todo en la API esta organizado y falla al ver que por ejemplo un libro no tiene el dato de "autor"

-El diseño nunca a sido mi fuerte siento que quedo a deber siempre en esa parte

## Retrospectiva

### ¿Qué hice bien?
- Aprendi bastante sobre un nuevo lenguaje, aprendi sobre creacion de aplicaciones de escritorio, y una notable mejoria en cuanto a entrgables.

### ¿Qué no salio bien?
- El diseño deja mucho que desear.
- Hay problemas con el consumo de la API
- Perdi mucho tiempo viendo como leer el JSON, ya que no encontraba la sintaxis para poder tomar los datos que necesitaba.

### ¿Qué puedo hacer diferente?
- Me falta investigar bastante y mejorar la parte de planeacion creo que dedicare mas tiempo en ello ya que puede ayudarme bastante decidir que quiero lograr desde un principio.

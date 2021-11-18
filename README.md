# Correo

> Automatización de envío de correos de notas.


## instalación
```zsh
git clone https://github.com/camilousa/correo.git
cd correo
pip install -r requirements.txt 
```

## Usage

Para ejecutar el programa para enviar los correos se utiliza el comando

```zsh
python sendEmail.py <path> <data_shett> <template>
```
* ```<path>``` es la direccion de la hoja de cálculo (ver [**data**](###**data**))
* ```<data_sheet>``` el nombre de la hoja del archivo donde estan los datos (ver [**data**](###**data**))
* ```<template>``` nombre de la plantilla para los correos (ver [**templates**](###**templates**))

Si se desa relizar una revisión del output del programa y **NO** enviar los correos, se puede utilziar la opcion ```-d | --debug ```, este imprimira en consola el resultado del template para cada correo.

```zsh
python sendEmail.py <path> <data_shett> <template> -d
```


Para mas información de opciones del programa y su uso se pude ejecutar con la opcion ```-h | --help ```

```zsh
usage: sendEmail.py [-h] [-v] [--host {outlook,gmail}] [-e EMAIL_SHEET] [-d] path data_sheet template

Envio de correos automaticos capturando la información de una hoja de calculo

positional arguments:
  path                  Path del archivo a extraer los datos
  data_sheet            Nombre de la hoja donde estan los datos
  template              Nombre de la plantilla HTML para el correo (la plantilla debe estar en la
                        carpeta "./templates" agregar la extension del archivo eg: .html)

optional arguments:
  -h, --help            show this help message and exit
  -v, --version         show program version number and exit
  --host {outlook,gmail}
                        host del servidor de correo, las opciones disponibles son outlook y gmail
                        (default: outlook)
  -e EMAIL_SHEET, --email-sheet EMAIL_SHEET
                        Nombre de la hoja donde esta la meta información del correo (eg: asunto)
                        (default: email)
  -d, --debug           Comprobación del contenido del correo. Imprime el correo del destinatario y
                        el contenido del correo, NO se envia el correo al destinatario (default:
                        False)
```
### **templates**

Los templates o plantillas HTML deben ser guardados en la carpeta 'templates' y debe ser indicada como argumento posicional ```template``` en la linea de comandos en el momento de la ejecucón del programa.  


### **data**

El programa lee los datos para enviar en los correos desde una hoja de calculo .xlsx, se debe indicar la dirección o path del archivo con la dirección completa o relativa a la carpeta root del proyecto como argumento posicional ```path```. 

Se debe indicar como argumento posicional ```data_sheet``` la hoja del archivo que contiene la información que se envia en el correo incluido la dirección de correo del destinatario, la tabla de datos debe tener una columna con el nombre ```correo```.

| correo     | nombres | apellidos     | nota    |
| :---       |    :----:   | :---: | ---: |
| juan.sierra01@correo.usa.edu.co      | juan sebastian       | sierra | 5.0 |

El progama recolecta información del correo como el asunto, el remitente y demas datos generales del correo, por defecto los busca en la hoja del archivo con nombre ```email```, se puede seleccionar otra hoja con la opción ``` -e <HOJA> | --email-sheet <HOJA>```

| subject     | coures | teacher   | term   |
| :---       |    :----:   | :---: | ---: |
| Big data corte 2     | Big data  | Camilo | 2 |


### **variables de entorno**

El programa requiere leer varibles de entorno con información sensible para el servidor de correos, estas se pueden asignar directamente en la linea de comandos, o pueden ser almacenadas en un archivo ```.env```. Las variables de entorno que require el programa son: 

```zsh
HOST_OUTLOOK=
HOST_GMAIL=
PORT=
EMAIL_OUTLOOK=
PASSWORD_OUTLOOK=
EMAIL_GMAIL=
PASSWORD_GMAIL=
```

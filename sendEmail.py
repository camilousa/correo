import os
from dotenv import load_dotenv
import smtplib
import traceback
import getpass
import argparse
import pathlib

from jinja2 import Environment, PackageLoader, select_autoescape, exceptions
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

load_dotenv() # lectura de las env variables del archivo .env
# Creacion de un env para encontrar y renderizar templates
env = Environment( 
        loader=PackageLoader("sendEmail"),
        autoescape=select_autoescape()
    )

def init_argparse():
    parser = argparse.ArgumentParser(
            description="Envio de correos automaticos capturando la información de una hoja de calculo",
            formatter_class=argparse.ArgumentDefaultsHelpFormatter)

    parser.add_argument('-v', '--version', action='version', version='%(prog)s 1.0.0')
    
    parser.add_argument('path', 
            type=pathlib.Path, 
            help='Path del archivo a extraer los datos')
    parser.add_argument('data_sheet', 
            help='Nombre de la hoja donde estan los datos')
    parser.add_argument('template',
            help='Nombre de la plantilla  HTML para el correo (la plantilla debe estar en la carpeta "./templates" agregar la extension del archivo eg: .html)') 
    parser.add_argument('-e','--email-sheet', default="email", type=str,
            help='Nombre de la hoja donde esta la meta información del correo (eg: asunto)')

    parser.add_argument('-d', '--debug', action='store_true', default=False,
            help='Comprobación del contenido del correo. Imprime el correo del destinatario y el contenido del correo, NO se envia el correo al destinatario')

    
    return parser

def send_email(smtp, source_address, destination_address, html, meta):
    try:
        msg = MIMEMultipart('alternative')
        message = meta['message'] if 'message' in meta else "TEST MAIL"
        msg['From'] = source_address
        msg['To'] = destination_address
        msg['Subject'] = meta['subject']
        if 'cc' in meta:
            msg['Cc'] = meta['cc']
        
        msg.attach(MIMEText(message, 'plain'))
        msg.attach(MIMEText(html, 'html'))
        smtp.send_message(msg)
        del msg

    except Exception as e:
        print('no se pudo enviar el correo')
        raise e      

def start_server(host, port, source_address, password):
    smtp = smtplib.SMTP(host, port)
    smtp.starttls()
    smtp.login(source_address, password)
    return smtp
        
def read_data(path, sheet, header=0):
    df = pd.read_excel(path, sheet, header=header)
    return df

def render_template(template, data, meta):
    template = env.get_template(template)
    return template.render(data=data, meta=meta)

def main(data_path, data_sheet, meta_sheet, template, host_name='outlook', debug=False):
    print("*** EMAIL***")
    host = os.getenv('HOST_GMAIL') if host_name == 'gmail' else os.getenv('HOST_OUTLOOK')
    port = os.getenv('PORT')
    source_address = os.getenv('EMAIL_GMAIL') if host_name == 'gmail' else os.getenv('EMAIL_OUTLOOK')
    password = os.getenv('PASSWORD_GMAIL') if host_name == 'gmail' else os.getenv('PASSWORD_OUTLOOK')
    destination_address = os.getenv('DESTINATION')

    if not source_address:
        source_address = input(f"correo {host_name}: ")
        if not password:
            try:
                password = getpass.getpass()
            except Exception as error:
                print('ERROR', error)

    try:
        # tabla de notas
        data = read_data(data_path, data_sheet) 
        # informacion del correo.
        meta_data = read_data(data_path, meta_sheet).to_dict('records')[0]
    except Exception as e:
        print("Error extrayendo los datos:")
        if debug:
            print(traceback.format_exc())
        else:
            print(e)
        exit(1)

    try:
        smtp = start_server(host, port, source_address, password)

        for _, row in data.iterrows():
            d = row.to_dict()
            d.pop('id')
            destination_address = d['correo']
            html = render_template(template, d, meta_data)
            
            print("\nsending to...{}".format(destination_address))
            print("\nsubject:{}".format(meta_data['subject']))
            if 'cc' in meta_data: print("\ncc to...{}".format(meta_data['cc']))
            
            if debug:
                print(html)
            else:
                send_email(smtp, source_address, destination_address, html, meta_data)
            break

        smtp.quit()
    except exceptions.TemplateNotFound as e:
        print(f"Plantilla no encontrada {e}.")
        print("Revise si el archivo se encuentra en la carpeta 'templates/'.")
    except Exception as e:
        print("Error enviando correos: ")
        if debug:
            print(traceback.format_exc())
        else:
            print(e)
        exit(1)


if __name__ == '__main__':
    parser = init_argparse()
    args = parser.parse_args()
    main(args.path, args.data_sheet, args.email_sheet, args.template, debug=args.debug)

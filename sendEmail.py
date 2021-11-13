import os
from dotenv import load_dotenv
import smtplib
import traceback
import getpass

from jinja2 import Environment, PackageLoader, select_autoescape
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd

load_dotenv() # lectura de las env variables del archivo .env
# Creacion de un env para encontrar y renderizar templates
env = Environment( 
        loader=PackageLoader("sendEmail"),
        autoescape=select_autoescape()
    )

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

        smtp.quit()
    except Exception as e:
        print("Error enviando correos: ")
        if debug:
            print(traceback.format_exc())
        else:
            print(e)
        exit(1)


if __name__ == '__main__':
    data_path = './sample_data/Aprendizaje-máquina-corte2.xlsx'
    data_sheet = 'mlcorte2'
    meta_sheet = 'email'
    template = 'template_camilo.html'
    debug=True
    main(data_path, data_sheet, meta_sheet, template, debug=debug)

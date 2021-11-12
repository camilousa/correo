import smtplib
import os
import traceback
import getpass

from openpyxl import load_workbook
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


def send_mail(smtp, source_address, destination_address, subject, html, cc):
    try:
        msg = MIMEMultipart('alternative')
        message = "TEST MAIL"
        msg['From'] = source_address
        msg['To'] = destination_address
        msg['Subject'] = subject
        if cc:
            msg['Cc'] = cc

        msg.attach(MIMEText(message, 'plain'))
        msg.attach(MIMEText(html, 'html'))
        smtp.send_message(msg)
        del msg

    except:
        print(traceback.format_exc())


def extract_xlsx_data(xlsx_file, range):
    workbook = load_workbook(filename=xlsx_file, data_only=True)
    sheet = workbook.active
    data = []
    for row in sheet[range]:
        row_data = []
        for col in row:
            if isinstance(col.value, float) or isinstance(col.value, int):
                value = round(col.value, 1)
                str_value = "{:.1f}".format(value)
            else:
                value = col.value
                str_value = str(value) if value is not None else ""
            row_data.append(str_value)
        data.append(row_data)

    return data


def create_html_table(course, teacher, group, schedule, student_name,
                      headers, data, template_file="html_template.html"):

    with open(template_file) as f:
        html = f.read()

    html = html.replace("{{course}}", course)
    html = html.replace("{{teacher}}", teacher)
    html = html.replace("{{group}}", group)
    html = html.replace("{{schedule}}", schedule)
    html = html.replace("{{student_name}}", student_name)

    html_headers = "\n"
    for header in headers[:-1]:
        html_headers += '<th>{}</th> \n'.format(header)
    html_headers += '<th class="final">{}</th> \n'.format(headers[-1])

    html_data = "\n"
    for value in data[:-1]:
        html_data += '<td>{}</td> \n'.format(value)


    html_data += '<td class="final">{}</td> \n'.format(data[-1])

    html = html.replace("{{headers}}", html_headers)
    html = html.replace("{{data}}", html_data)

    return html


def main():
    print("*** EMAIL***")
    range = "B1:I34"
    cc = ""
    email_col = 0
    name_col = 2
    offset = 1
    source_address = os.getenv('LOGIN')
    destination_address = os.getenv('DESTINATION')
    password = os.getenv('PASSWORD')

    if not source_address:
        source_address = input("correo outlook: ")

    if not password:
        try:
            password = getpass.getpass()

        except Exception as error:
            print('ERROR', error)
    try:
        smtp = smtplib.SMTP(host='smtp-mail.outlook.com', port=587)
        #smtp = smtplib.SMTP(host='smtp.gmail.com', port=587)
        smtp.starttls()
        smtp.login(source_address, password)
        file_name = "big_data_corte2"
        data = extract_xlsx_data(f"{file_name}.xlsx", range)
        course = data[0][0]
        teacher = data[1][0]
        group = data[2][0]
        schedule = data[3][0]

        for row in data[offset:]:
            destination_address = row[email_col]
            print("\nsending to...{}".format(destination_address))
            print("\nsubject:{}".format(file_name))
            if cc: print("\ncc to...{}".format(cc))

            html = create_html_table(course, teacher, group,
                                 schedule, student_name=row[name_col],
                                     headers=data[offset-1], data=row)

            print(html)
            #send_mail(smtp, source_address, destination_address, subject=file_name, html=html, cc=cc)
    except:
        traceback.print_exc()

main()
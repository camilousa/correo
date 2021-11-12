from jinja2 import Environment, PackageLoader, select_autoescape
import pandas as pd

env = Environment(
        loader=PackageLoader("sendEmail"),
        autoescape=select_autoescape()
    )

def send_email():
    pass

def read_data(path, sheet, header=0):
    df = pd.read_excel(path, sheet, header=header)
    return df

def render_template(template, data, meta):
    template = env.get_template(template)
    return template.render(data=data, meta=meta)

def main(data_path, data_sheet, meta_sheet, template):
    data = read_data(data_path, data_sheet)
    meta_data = read_data(data_path, meta_sheet).to_dict('records')[0]
    for index, row in data.iterrows():
        d = row.to_dict()
        d.pop('id')
        html = render_template(template, d, meta_data)
        print(html)
        break

if __name__ == '__main__':
    data_path = './sample_data/Aprendizaje-máquina-corte2.xlsx'
    data_sheet = 'mlcorte2'
    meta_sheet = 'email'
    template = 'template_camilo.html'
    main(data_path, data_sheet, meta_sheet, template)

import pandas as pd


def send_email():
    pass

def read_data(path, sheet, header=0):
    df = pd.read_excel(path, sheet, header=header)
    return df

def render_template():
    pass

def main(data_path, data_sheet, meta_sheet):
    data = read_data(data_path, data_sheet)
    meta_data = read_data(data_path, meta_sheet)
    print(meta_data)


if __name__ == '__main__':
    data_path = './sample_data/Aprendizaje-máquina-corte2.xlsx'
    data_sheet = 'mlcorte2'
    meta_sheet = 'email'
    main(data_path, data_sheet, meta_sheet)

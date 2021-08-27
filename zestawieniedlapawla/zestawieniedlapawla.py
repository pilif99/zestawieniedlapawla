
import os
import pandas as pd
import mail

lista_nazw = ['Template', 'Vorlage', 'Plantilla', 'Modello', 'Arkusz1']
lista_krajow = ['DE', 'ES', 'IT']
klucze = [(x, y) for x in ['boot', 'protectivegear', 'coat', 'sportactivityglove', 'pants'] for y in ['DE', 'ES', 'IT']]
lista_id = [3, 15, 22, 6, 18, 25, 2, 14, 11, 4, 16, 23, 9, 21, 28]
slownik_id = dict(zip(klucze, lista_id))

lista = os.listdir(r"C:\Users\FLorenzLen\Desktop\Nowy folder")
os.chdir(r"C:\Users\FLorenzLen\Desktop\Nowy folder")

lista_df = []

for i in lista:

    if 'xlsx' in i or 'xlsm' in i:
        xl = pd.ExcelFile(i)
        for j in lista_nazw:
            if j in xl.sheet_names:
                df = xl.parse(j, header = None)
                df = df.drop(labels = [0, 1], axis = 0)
                df = df.rename(columns=df.iloc[0])
                df = df[1:]
                df = df.drop(labels = [3], axis = 0)
                df = df.reset_index()
                df = df[['item_sku', 'feed_product_type', 'item_name', 'product_description', 'main_image_url',
                        'other_image_url1', 'other_image_url2', 'other_image_url3',
                        'other_image_url4', 'other_image_url5', 'other_image_url6']]
                for k in lista_krajow:
                    if k in i:
                        df['Len'] = k
                        df['Destination'] = 1
                        df['BaselinkerCategory'] = slownik_id[(df.at[0, 'feed_product_type'], k)]
                        for l in ['Desc1', 'Desc2', 'Desc3', 'Desc4', 'Desc5', 'Desc6', 'Desc7', 'BaselinkerId']:
                            df[l] = ''
                df = df.drop(columns = ['feed_product_type'])
                df = df[['item_sku', 'Len', 'Destination', 'item_name', 'product_description', 'main_image_url', 'Desc2', 'other_image_url1', 'Desc3',
                        'other_image_url2', 'Desc4', 'other_image_url3', 'Desc5', 'other_image_url4', 'Desc6', 'other_image_url5', 'Desc7', 'other_image_url6',
                        'BaselinkerCategory', 'BaselinkerId']]
                df.columns = ['Indeks', 'Len', 'Destination', 'Name', 'Desc1', 'Foto1', 'Desc2', 'Foto2', 'Desc3',
                        'Foto3', 'Desc4', 'Foto4', 'Desc5', 'Foto5', 'Desc6', 'Foto6', 'Desc7', 'Foto7',
                        'BaselinkerCategory', 'BaselinkerId']
                lista_df.append(df)
                

df = pd.concat(lista_df)
df = df.reset_index(drop = True)
for i in ['Foto1', 'Foto2', 'Foto3', 'Foto4', 'Foto5', 'Foto6', 'Foto7']:
    try:
        df[i] = df[i].str.replace('/AMZPhoto/', '/AMZPhoto/1mb/')
    except:
        pass
    try:
        df[i] = df[i].str.replace('png', 'jpg')
    except:
        pass
os.chdir(r"C:\Users\FLorenzLen\source\repos\zestawieniedlapawla\zestawieniedlapawla")
df.to_excel('Zestawienie.xlsx', index = False)
a = mail.Mail()
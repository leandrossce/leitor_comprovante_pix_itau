import pdfplumber
import pandas as pd
import os
import re

def extract_text_from_coordinates(page, coordinates):
    cropped = page.crop(coordinates)
    return cropped.extract_text()


resultados = []

origem_comprovantes='C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\leituraDeComprovantes\\PIX\\'
destino_relatorio='C:\\Users\\Gabriel\\Desktop\\chromedriver_win32\\leituraDeComprovantes\\resultados.xlsx'

# Percorra todos os arquivos no diretório
for raiz, subdiretorios, arquivos in os.walk(origem_comprovantes):
    for nome_arquivo in arquivos:
        caminhoPDF = os.path.join(raiz, nome_arquivo)

        with pdfplumber.open(caminhoPDF) as pdf:
            for page in pdf.pages:

                coordinates = {
                    "tipo":(128.0, 53.0, 468.0, 86.0),
                    "beneficiario": (156.0, 174.0, 493.0, 191.0), #
                    "cnpj_beneficiario": (106.1488823,  210.783791, 245.0009121,220.750891), 
                    "valor_pagamento": (155.0, 254.0,  286.0, 275.0),
                    "data_pagamento": (155.0, 275.0, 207.0, 288.0)

                }

                page_data = {}

                if "QR" in extract_text_from_coordinates(page, coordinates["tipo"]):
                    coordinates = {
                        "tipo":(128.0, 53.0, 468.0, 86.0),
                        "beneficiario":  (106.1488823, 210.786191, 177.3841057, 220.753291), #
                        "cnpj_beneficiario": (106.1488823, 227.095991, 245.0009121, 237.063091), 
                        "valor_pagamento": (132.7413595, 440.032891, 194.5773532, 449.999991),
                        "data_pagamento": (131.0862659, 643.002491, 180.9616343, 652.969591)

                    }
                if "dentificador" in extract_text_from_coordinates(page, coordinates["valor_pagamento"]):
                    coordinates = {
                        "tipo":(128.0, 53.0, 468.0, 86.0),
                        "beneficiario":  (60.0, 209.0, 365.0, 221.0), #
                        "cnpj_beneficiario": (30.0, 225.0, 248.0, 237.0), 
                        "valor_pagamento": (95.0, 438.0, 248.0, 451.0),
                        "data_pagamento": (14.0, 637.0, 184.0, 656.0)

                    }                    
    

                for key, coord in coordinates.items():
                    text = extract_text_from_coordinates(page, coord)
                    page_data[key] = text
                                        
                    if (coord == coordinates["valor_pagamento"]) and ("QR" in extract_text_from_coordinates(page, coordinates["tipo"])):
                        text = extract_text_from_coordinates(page, coord)
                        text=text.split("final:")
                        try:
                        
                            page_data[key] = text[1]
                        except:
                            page_data[key] = text[0]


                    if coord == coordinates["cnpj_beneficiario"]:
                        text = extract_text_from_coordinates(page, coord)
                        page_data[key] = text
                        padrao = r"(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})"
                        resultado = re.search(padrao, page_data[key])
                        if resultado:
                            
                            page_data[key] = resultado.group(1)
                    
                    elif coord == coordinates["data_pagamento"]:
                        text = extract_text_from_coordinates(page, coord)
                        page_data[key] = text
                        padrao = r"(\d{2}/\d{2}/\d{4})"
                        resultado = re.search(padrao, page_data[key])
                        if resultado:
                            
                            page_data[key] = resultado.group(1)




                resultados.append(page_data)


df = pd.DataFrame(resultados)
df.to_excel(destino_relatorio, index=False, engine='openpyxl')


# Pegar os dados da planilha 
# Tipo nome do curso, nome participante, tipo de participação, data do inicio, data do final, carga horária, data da emissão do certificado
# Transferir os dados da planilha para a imagem do certificado
# Pegar os dados da planilha 

import openpyxl as opyxl
from PIL import Image, ImageDraw, ImageFont

# Abrir a planilha
alunos = opyxl.load_workbook('relacao.xlsx')
relacao = alunos['Sheet1']

for indice, linha in enumerate(relacao.iter_rows(min_row=2)):
    # Cada célula que contém a info que precisamos  
    curso = linha[0].value        # Nome do curso
    participante = linha[1].value # Nome do participante
    participacao = linha[2].value # Tipo de participação
    carga = linha[5].value        # Carga horária
    
    datainicio = linha[3].value   # Data de início do curso
    datafim = linha[4].value      # Data final do curso
    
    dataemissao = linha[6].value   # Data de emissão do certificado
  
    # Transferir os dados da planilha para a imagem do certificado
    # Definindo a fonte a ser usada
    fontenome = ImageFont.truetype('./fontes/tahomabd.ttf',55)
    fontegeral = ImageFont.truetype('./fontes/tahoma.ttf',45)
    fontedata = ImageFont.truetype('./fontes/tahoma.ttf',35)
    
    image = Image.open('./arte.jpg')
    desenhar = ImageDraw.Draw(image)
    
    desenhar.text((440,550), participante, fill='black',font=fontenome)
    desenhar.text((445,645), curso, fill='black', font=fontegeral)
    desenhar.text((650,732), participacao, fill='black',font=fontegeral)
    desenhar.text((700,817), str(carga) + ' horas', fill='black',font=fontegeral)
    
    desenhar.text((275,960), datainicio, fill='black', font=fontedata)
    desenhar.text((275,1065), datafim, fill='black', font=fontedata)
    desenhar.text((1510,1065), dataemissao, fill='black', font=fontedata)
   
    
    image.save(f'./certificados/{participante}.png')
    
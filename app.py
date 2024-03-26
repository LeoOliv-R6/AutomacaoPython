'''
Pegar os dados da planilha

'''

import openpyxl 
from PIL import Image, ImageDraw, ImageFont

# Abre a planilha
workbook_alunos = openpyxl.load_workbook('planilha_alunos.xlsx')
sheet_alunos = workbook_alunos['Sheet1']

for indice, linha in enumerate(sheet_alunos.iter_rows(min_row=2)): # função max_row para limitar a quantidade de linhas geradas
    # cada céclula que contém a info que preciso
    nome_curso = linha[0].value # atribuindo valor ao indice 0 
    nome_participante = linha[1].value 
    tipo_participacao = linha[2].value
    data_inicio = linha[3].value
    data_termino = linha[4].value
    carga_horaria = linha[5].value
    data_emissao = linha[6].value
    
    
# Transferir dados da planilha para o certificado
fonte_nome = ImageFont.truetype('./tahomabd.ttf', 85)
fonte_geral = ImageFont.truetype('./tahoma.ttf', 80)
fonte_data = ImageFont.truetype('./tahoma.ttf', 55)

image = Image.open('./certificado_padrao.jpg')
desenhar = ImageDraw.Draw(image)

desenhar.text((1030,828), nome_participante, fill='black', font=fonte_nome)
desenhar.text((1085,950), nome_curso, fill='black', font=fonte_geral)
desenhar.text((1450,1065), tipo_participacao, fill='black', font=fonte_geral)
desenhar.text((1510,1185), str(carga_horaria), fill='black', font=fonte_geral)

desenhar.text((750,1775), data_inicio, fill='blue', font=fonte_data)
desenhar.text((750,1925), data_termino, fill='blue', font=fonte_data)

desenhar.text((2220,1925), data_emissao, fill='red', font=fonte_data)

image.save(f'./{indice}  {nome_participante} certificado.png')
    
import openpyxl
from PIL import Image, ImageDraw, ImageFont

def geraCertificados(planilha, pasta):

    planilha_alunos = openpyxl.load_workbook(planilha)
    pagina_alunos = planilha_alunos['Planilha1']
    fonteNome = ImageFont.truetype('./IMPRISHA.TTF',70) #70 é o tamanho da fonte
    fonte = ImageFont.truetype('./IMPRISHA.TTF',40)


    

    for linha in pagina_alunos.iter_rows(min_row=2,max_row=3):

        imagem = Image.open('./CERTIFICADO DE CONCLUSÃO.png')
        desenhar = ImageDraw.Draw(imagem)

        nome = linha[0].value #pegando o primeiro valor "0" da primeira linha e colocando na variavel nome
        curso = linha[1].value
        carga = str(linha[2].value)
        data_inicio = (str(linha[3].value))[:10]
        data_inicio = (f"{data_inicio[8:10]}-{data_inicio[5:7]}-{data_inicio[0:4]}")#formatando a data pro formato dd/mm/yyyy
        data_fim = str(linha[4].value)
        data_fim = (f"{data_fim[8:10]}-{data_fim[5:7]}-{data_fim[0:4]}")
        data_emissao = str(linha[5].value)
        data_emissao = (f"{data_emissao[8:10]}-{data_emissao[5:7]}-{data_emissao[0:4]}")
        

        desenhar.text((370,350),nome,fill='black',font=fonteNome)
        desenhar.text((1020,450),curso,fill='black',font=fonte)
        desenhar.text((530,495),carga,fill='black',font=fonte)
        desenhar.text((190,600),data_inicio,fill='black',font=fonte)
        desenhar.text((190,680),data_fim,fill='black',font=fonte)
        desenhar.text((190,770),data_emissao,fill='black',font=fonte)
        
        imagem.save(f'{pasta}/CERTIFICADO DE CONCLUSAO DE {nome}.png')


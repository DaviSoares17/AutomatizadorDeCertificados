from PySimpleGUI import PySimpleGUI as sg
from gerador import geraCertificados

sg.theme("Reddit")

layout = [
    [sg.Text('Selecione o arquivo excel'), sg.FileBrowse(key="pastaExcel")], #key serve para indicar o nome que vai ser usado para puxar o valor que o file browse ta retornando
    [sg.Text('Selecione o caminho de salvamento dos certificados'), sg.FolderBrowse(key="pastaDestino")],
    [sg.Button('Executar aplicação')]
]

janela = sg.Window('Gerador de certificados', layout)

while True:
    eventos, valores = janela.read()
    if eventos == sg.WINDOW_CLOSED:
        break

    if eventos == "Executar aplicação": 
        if valores['pastaExcel'] and valores['pastaDestino']:
            arquivo_selecionado = valores['pastaExcel']
            pasta_selecionada = valores['pastaDestino']
            sg.popup("Certificados gerados")
            geraCertificados(arquivo_selecionado, pasta_selecionada)

        else:
            sg.popup("Selecione as duas opções")

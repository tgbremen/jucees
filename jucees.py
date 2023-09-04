import PySimpleGUI as sg
from time import sleep
import jucees_backend
from PySimpleGUI.PySimpleGUI import Output


instrucoes = """Instruções: \n
1. Os CNPJs não são validados pelo programa. Se houver erro, não encontrará nada ou dará erro no programa. \n
2. Os CNPJs podem ser escritos com ou sem pontos, hífens e barras. \n
3. Os zeros iniciais não podem ser omitidos ou o CPF / CNPJ não será encontrado. \n
4. Após clicar no botão Iniciar Execução, pode demorar um pouco até que o driver do chrome seja atualizado. \n
5. Não use ou feche a janela do chrome durante a execução do programa, apenas a minimize. Não abra outra aba nessa janela. \n
6. Os arquivos ficarão dentro da pasta jucees_resultado, que estará localizada na mesma página deste arquivo.  \n
"""


class TelaPython:
    def __init__(self) -> None:
        
        sg.theme('DarkGrey1')
        
        layout_splash = [
                        [sg.Image(r'd:\py\vilavelha.png', expand_y = True, expand_x = True)]
            
        ]
        
        layout_text_splash = [

                        [sg.Text('Carregando JUCEES')]
        ]
        
        
        sg.splash = sg.Window(title="Jucees", layout = layout_splash, finalize=True, no_titlebar=True)
        sg.splash_text = sg.Window(title="Jucees", layout = layout_text_splash, finalize=True, no_titlebar=True, alpha_channel = 0.3)
        sg.splash_text.read(timeout = 2000)
        sg.splash.read(timeout=2000)
        sg.splash.Size

        sg.splash.close()
        sg.splash_text.close()
        
        layout_tela =   [
                    #[sg.Text("Fonte dos Dados:"), sg.Radio('CENSEC', 'fonte', key= 'Fonte_censec', default = False), sg.Radio('CANP', 'fonte', key = 'Fonte_canp', default = True), sg.Radio('Ambos', 'fonte', key= 'Ambos'), sg.Radio('CESDI', 'fonte', key= 'Fonte_cesdi', default = False)],
                    #[sg.HSeparator()],
                    #[sg.Text("Nome da lista:")], [sg.Input(size=(25,1), key = 'nomedalista')],
                    #[sg.HSeparator()],
                    [sg.Text("Lista de CPFs ou CNPJs:")],
                    [sg.Multiline(size=(25,18), key = 'cpf_cnpj', expand_y = True, autoscroll = True), sg.Text(text = instrucoes)], 
                    [sg.Button('Iniciar Extração', key =  "Iniciar"), sg.Button('Pausar Extração', key = "Pause", visible = False), sg.Button('Parar Extração', key = 'Stop', visible = False, disabled = True)], 
                    [sg.Output(size=(150,20), key= "Output")]
                    ]


        self.janela = sg.Window(title="JUCEES", layout = layout_tela, finalize=True)
        
    
    def Iniciar(self):
        Output.expand_y = True
        #self.janela.read()
        print ('Scraper JUCEES')
        print ("Desenvolvido no Núcleo de Inovação, Prospecção e Análise de Dados (CGU-ES/NAE/NIPAD)")
        print ("Atualizações disponíveis em http://github.com/tgbremen/jucees")
        print ("-------------------------------------------------------------------------")

        while(True):
            event, self.values = self.janela.read()
            if event == sg.WIN_CLOSED:
                break
            
            if event == "Iniciar":

                # Armazena valores da tela
                lista_CPF_CNPJ = self.values['cpf_cnpj']
                #nome_lista = self.values['nomedalista']

                # Remove os itens da lista que representam linhas em branco
                lista_CPF_CNPJ = list(filter(None,(lista_CPF_CNPJ.split('\n'))))

                # Executa o scraper
                print("Iniciando JUCEES")
                scraper = jucees_backend.jucees()
                scraper.scrap(lista_CPF_CNPJ)

            if event == "Stop":
                self.janela['Iniciar'].update("Iniciar Extração", disabled = False)
                self.janela['Stop'].update("Parar Extração", disabled = True)

        self.janela.close()       


tela = TelaPython()
tela.Iniciar()




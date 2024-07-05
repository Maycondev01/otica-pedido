import PySimpleGUI as sg
import pandas as pd

# Função para verificar login
def verificar_login(login, senha):
    df_acessos = pd.read_excel('acessos.xlsx')  # Carrega o arquivo de acessos (login e senha)
    if login in df_acessos['login'].values and senha in df_acessos['senha'].values:
        return True
    else:
        return False

# Layout da tela de login
def mostrar_tela_login():
    sg.theme('DarkPurple4')  # Escolhe um tema para a interface gráfica

    layout = [
        [sg.Image('logo.png')],
        [sg.Text('Login'), sg.Input(key='login')],
        [sg.Text('Senha'), sg.Input(key='senha', password_char='*')],
        [sg.Button('Entrar'), sg.Button('Cancelar')]
    ]

    window = sg.Window('Login', layout, element_justification='center')

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancelar':
            window.close()
            return False
        elif event == 'Entrar':
            login = values['login']
            senha = values['senha']
            if verificar_login(login, senha):
                sg.popup('Login realizado com sucesso!')
                window.close()
                return True
            else:
                sg.popup_error('Login ou senha incorretos. Tente novamente.')

    window.close()
    return False

# Função principal para iniciar o login
def iniciar_login():
    return mostrar_tela_login()

# Execução do login (para testes)
if __name__ == '__main__':
    if iniciar_login():
        sg.popup('Login bem-sucedido! Agora, o sistema de pedidos será iniciado.')

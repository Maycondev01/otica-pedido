import PySimpleGUI as sg
import pandas as pd

# Função para verificar login
def verificar_login(login, senha):
    # Carregar dados da tabela 'acessos' do arquivo Excel
    try:
        df_acessos = pd.read_excel('acessos.xlsx', sheet_name='acessos', dtype=str)
        df_acessos['login'] = df_acessos['login'].astype(str).str.strip()
        df_acessos['senha'] = df_acessos['senha'].astype(str).str.strip()
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return False

    # Verificar se o login e a senha correspondem a uma linha no DataFrame
    usuario_valido = df_acessos[(df_acessos['login'] == login) & (df_acessos['senha'] == senha)]

    if not usuario_valido.empty:
        return True
    else:
        return False

# Layout da tela de login
def mostrar_tela_login():
    sg.theme('DarkPurple4')  # Escolhe um tema para a interface gráfica

    layout = [
        [sg.Image('logoint.png')],
        [sg.Text('Login'), sg.Input(key='login')],
        [sg.Text('Senha'), sg.Input(key='senha', password_char='*')],
        [sg.Button('Entrar', bind_return_key=True), sg.Button('Cancelar')]
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
                window.close()
                return True
            else:
                sg.popup_error('Login ou senha incorretos. Tente novamente.', title="Erro", keep_on_top=True, button_type=5)

    window.close()
    return False

# Função principal para iniciar o login
def iniciar_login():
    return mostrar_tela_login()

# Execução do login (para testes)
if __name__ == '__main__':
    if iniciar_login():
        print('Login bem-sucedido! Agora, o sistema de pedidos será iniciado.')

import login
from pedido import iniciar_sistema
from historico_cliente import iniciar_historico_cliente

def main():
    if login.iniciar_login():
        iniciar_sistema()
        iniciar_historico_cliente()

if __name__ == '__main__':
    main()

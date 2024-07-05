import login
from pedido import iniciar_sistema

def main():
    if login.iniciar_login():
        iniciar_sistema()

if __name__ == '__main__':
    main()

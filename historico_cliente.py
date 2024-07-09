import PySimpleGUI as sg
import pandas as pd
from fpdf import FPDF
from datetime import datetime

# Classe PDF estendida para personalizar o formato do PDF
class PDF(FPDF):
    def __init__(self, pedido):
        super().__init__()
        self.pedido = pedido
        self.WIDTH = 210  # largura da página em mm (A4)
        self.HEIGHT = 297  # altura da página em mm (A4)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'@_Oticanewlook', 0, 0, 'C')

def iniciar_historico_cliente():
    layout = [
        [sg.Text("CPF do Cliente:"), sg.InputText(key='cpf')],
        [sg.Button("Buscar Histórico")],
        [sg.Text("", size=(50, 1), key='mensagem')],
        [sg.Table(values=[], headings=['ID', 'Data Saída', 'Data Entrega', 'Nome', 'Valor'], key='tabela_pedidos', auto_size_columns=True, display_row_numbers=True, enable_events=True)],
        [sg.Button("Gerar PDFs")]
    ]

    janela = sg.Window("Histórico do Cliente", layout)

    while True:
        evento, valores = janela.read()
        if evento == sg.WINDOW_CLOSED:
            break
        if evento == "Buscar Histórico":
            cpf = valores['cpf']
            if not cpf:
                janela['mensagem'].update("Por favor, insira o CPF do cliente.")
            else:
                pedidos_cliente = buscar_historico_cliente(cpf)
                if pedidos_cliente.empty:
                    janela['mensagem'].update("Nenhum pedido encontrado para este CPF.")
                    janela['tabela_pedidos'].update(values=[])
                else:
                    pedidos = pedidos_cliente[['id', 'data_saida', 'data_entrega', 'nome', 'valor']].values.tolist()
                    janela['tabela_pedidos'].update(values=pedidos)
                    janela['mensagem'].update("")

        if evento == "Gerar PDFs":
            selected_indices = janela['tabela_pedidos'].SelectedRows
            if not selected_indices:
                sg.popup("Nenhum pedido selecionado", "Por favor, selecione pelo menos um pedido para gerar o PDF.")
            else:
                pedidos_cliente = buscar_historico_cliente(valores['cpf'])
                pedidos_selecionados = pedidos_cliente.iloc[selected_indices]
                gerar_pdfs(pedidos_selecionados)

    janela.close()

def buscar_historico_cliente(cpf):
    try:
        # Lê os dados do Excel
        df = pd.read_excel('pedidos.xlsx', sheet_name='Sheet1')  # Certifique-se de que a planilha está correta
        df['cpf'] = df['cpf'].astype(str)  # Converte a coluna 'cpf' para string
        cpf = str(cpf)  # Garante que o CPF fornecido também é string

        print(f"Dados do Excel: {df.head()}")  # Depuração

        # Verifica se a coluna 'cpf' está presente no DataFrame
        if 'cpf' not in df.columns:
            print("A coluna 'cpf' não está presente no arquivo Excel.")
            return pd.DataFrame()  # Retorna um DataFrame vazio

        # Filtra os pedidos pelo CPF
        pedidos_cliente = df[df['cpf'] == cpf]
        print(f"Pedidos encontrados para o CPF {cpf}: {pedidos_cliente}")  # Depuração
        return pedidos_cliente
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return pd.DataFrame()  # Retorna um DataFrame vazio em caso de erro

def gerar_pdfs(pedidos):
    for _, pedido in pedidos.iterrows():
        gerar_pdf_pedido(pedido)

    sg.popup("PDFs Gerados", "Os PDFs dos pedidos foram gerados com sucesso.")

def gerar_pdf_pedido(pedido):
    pdf = PDF(pedido)
    pdf.add_page()

    # Informações do laboratório
    pdf.image('logo.png', x=10, y=pdf.get_y(), w=25)  # Inserir logo
    pdf.set_font('Arial', 'B', 8)

    # Espaço antes de iniciar os elementos
    wpp_text = '81996615539'
    inst_text = '@_oticanewlook'
    tele_text = '8132048009'
    icon_width = 5  # largura das imagens
    icon_text_padding = 2  # espaço entre a imagem e o texto
    icon_y_offset = 7.5  # deslocamento vertical das imagens
    horizontal_offset = 10  # deslocamento horizontal adicional
    total_width = (icon_width + icon_text_padding + pdf.get_string_width(wpp_text) +
                10 +  # espaço entre WhatsApp e Instagram
                icon_width + icon_text_padding + pdf.get_string_width(inst_text) +
                10 +  # espaço entre Instagram e número da OS
                icon_width + icon_text_padding + pdf.get_string_width(tele_text) +
                10 +
                pdf.get_string_width(f'OS: {pedido["id"]}'))

    # Centralização dos contatos (whatsapp e instagram)
    x_start = (pdf.WIDTH - total_width) / 2 + horizontal_offset
    y_start = pdf.get_y()

    # Posicionar e desenhar logo do WhatsApp e número
    pdf.set_x(x_start)
    pdf.image('wpp.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(wpp_text), 20, wpp_text, 0, 0, 'L')

    # Espaço entre o número do WhatsApp e o Instagram
    pdf.set_x(pdf.get_x() + 7)

    # Posicionar e desenhar logo do Instagram e username
    pdf.image('inst.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(inst_text), 20, inst_text, 0, 0, 'L')

    pdf.set_x(pdf.get_x() + 7)

    pdf.image('tele.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(tele_text), 20, tele_text, 0, 0, 'L')

    pdf.set_x(pdf.get_x() + 5)


    # Número da OS
    pdf.cell(0, 15, f'OS: {pedido["id"]}', 0, 1, 'R')  # Adiciona nova linha após o número da OS

    pdf.set_y(pdf.get_y() - 10)  # Ajuste o valor (-10) conforme necessário para o posicionamento desejado

    # Nome do vendedor
    pdf.cell(0, 15, f'Vendedor: {pedido["vendedor"]}', 0, 1, 'R')  # Adiciona nova linha após o nome do vendedor

    pdf.ln(5)

    # Data de Saída e Data de Entrega à esquerda
    pdf.cell(90, 6, f"Data de Saída: {pedido['data_saida']}", 0, 0, 'L')
    pdf.cell(0, 6, f"Data de Entrega: {pedido['data_entrega']}", 0, 1, 'R')

    # Nome, Telefone e CPF à direita
    pdf.cell(90, 6, f"Nome: {pedido['nome']}", 0, 0, 'L')
    pdf.cell(0, 6, f"Telefone: {pedido['telefone']}", 0, 1, 'R')
    pdf.cell(90, 6, f"CPF: {pedido['cpf']}", 0, 1, 'L')

    # Dados da compra
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 8, 'Dados da Compra', 0, 1, 'C')

    pdf.set_font('Arial', '', 6)

    # Detalhes do pedido
    pdf.cell(0, 6, f"Lentes: {pedido['lentes']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Referência: {pedido['referencia']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Valor: R$ {pedido['valor']:.2f}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"P/ Conta: {pedido['p_conta']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Resta: {pedido['resta']}", 1, 1, 'L', 0)

    # Garantia
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 8, 'GARANTIA:', 0, 1, 'C')

    # Posição para o texto de garantia
    x = pdf.get_x()
    y = pdf.get_y()
    pdf.set_xy(x, y)

    # Texto da garantia
    pdf.set_font('Arial', 'I', 8)  # Diminua o tamanho da fonte aqui
    garantia_texto = (
        "3 meses em oxidação ou descasque não sendo renovável. \n"
        "1 ano para quebra sob defeito (na compra do óculos completo) havendo conserto ou troca (caso não havendo do mesmo modelo será trocado por outro modelo que de as suas lentes atuais garantia não é renovável. \n "
        "Garantia lentes 3 meses para defeito de fábrica, caso a garantia seja maior será necessário o certificado das lentes anexado a essa nota de pedido. \n"
        "Garantia não cobre quebra ou arranhões após a entrega. \n"
        "Cola ou qualquer outro produto acarretará a perca da garantia. Qualquer ajuste precisa ser com o técnico óptico. \n"
        "Não cancelamos pedidos em andamento."
    )

    # Define a cor de fundo e borda
    pdf.set_fill_color(230, 230, 230)  # Cor de fundo da caixa
    pdf.set_draw_color(100, 100, 100)  # Cor da borda da caixa

    # Dimensões da caixa para o texto de garantia
    largura_caixa = pdf.WIDTH - 20
    altura_caixa = 55

    # Desenha a caixa com bordas arredondadas
    pdf.rect(x, y, largura_caixa, altura_caixa, 'FD')

    # Configurações para o texto dentro da caixa
    pdf.set_xy(x + 5, y + 5)  # Posiciona o texto dentro da caixa
    pdf.set_font('Arial', 'I', 8)  # Define a fonte e tamanho

    # Ajuste o espaçamento entre as linhas
    cell_height = 6  # Altura da célula para controlar o espaçamento entre linhas
    pdf.multi_cell(largura_caixa - 10, cell_height, garantia_texto, 0, 'C', fill=False)

    # Espaço após o texto da garantia
    pdf.ln(5)

    # Garantia
    pdf.cell(0, 8, "Assinatura do Cliente:", 0, 1, 'C')
    pdf.multi_cell(0, 8, "___________________________________", 0, 'C')
    pdf.ln(5)

    # Informações do laboratório
    pdf.image('logo.png', x=10, y=pdf.get_y(), w=25)  # Inserir logo
    pdf.set_font('Arial', 'B', 8)

    # Espaço antes de iniciar os elementos
    wpp_text = '81996615539'
    inst_text = '@_oticanewlook'
    tele_text = '8132048009'

    icon_width = 5  # largura das imagens
    icon_text_padding = 2  # espaço entre a imagem e o texto
    icon_y_offset = 7.5  # deslocamento vertical das imagens
    horizontal_offset = 10  # deslocamento horizontal adicional
    total_width = (icon_width + icon_text_padding + pdf.get_string_width(wpp_text) +
                10 +  # espaço entre WhatsApp e Instagram
                icon_width + icon_text_padding + pdf.get_string_width(inst_text) +
                10 +  # espaço entre Instagram e número da OS
                    icon_width + icon_text_padding + pdf.get_string_width(tele_text) +
                10 +
                pdf.get_string_width(f'OS: {pedido["id"]}'))

    # Centralização dos contatos (whatsapp e instagram)
    x_start = (pdf.WIDTH - total_width) / 2 + horizontal_offset
    y_start = pdf.get_y()

    # Posicionar e desenhar logo do WhatsApp e número
    pdf.set_x(x_start)
    pdf.image('wpp.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(wpp_text), 20, wpp_text, 0, 0, 'L')

    # Espaço entre o número do WhatsApp e o Instagram
    pdf.set_x(pdf.get_x() + 7)

    # Posicionar e desenhar logo do Instagram e username
    pdf.image('inst.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(inst_text), 20, inst_text, 0, 0, 'L')

    pdf.set_x(pdf.get_x() + 7)

    pdf.image('tele.png', x=pdf.get_x(), y=y_start + icon_y_offset, w=icon_width)
    pdf.set_x(pdf.get_x() + icon_width + icon_text_padding)
    pdf.cell(pdf.get_string_width(tele_text), 20, tele_text, 0, 0, 'L')

    # Número da OS
    pdf.cell(0, 15, f'OS: {pedido["id"]}', 0, 1, 'R')  # Adiciona nova linha após o número da OS

    pdf.set_y(pdf.get_y() - 10)  # Ajuste o valor (-10) conforme necessário para o posicionamento desejado

    # Nome do vendedor
    pdf.cell(0, 15, f'Vendedor: {pedido["vendedor"]}', 0, 1, 'R')  # Adiciona nova linha após o nome do vendedor

    pdf.ln(5)
    # Criação da tabela para detalhes do laboratório
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(60, 6, 'Detalhes do Laboratório (Longe)', 1, 0, 'C')
    pdf.cell(60, 6, 'Detalhes do Laboratório (Perto)', 1, 0, 'C')
    pdf.cell(60, 6, 'DNP', 1, 1, 'C')

    pdf.set_font('Arial', '', 6)

    campos_longe = ['Longe_ESF_OD', 'Longe_CIL_OD', 'Longe_EIXO_OD', 'Longe_ESF_OE', 'Longe_CIL_OE', 'Longe_EIXO_OE']
    campos_perto = ['Perto_ESF_OD', 'Perto_CIL_OD', 'Perto_EIXO_OD', 'Perto_ESF_OE', 'Perto_CIL_OE', 'Perto_EIXO_OE']
    dnp = ['DNP_OD', 'DNP_OE','DNP_ALT','DNP_COR']

    # Preenche os dados na tabela
    for i in range(len(campos_longe)):
        pdf.cell(30, 6, f"{campos_longe[i].replace('_', ' ')}:", 1, 0, 'L')
        pdf.cell(30, 6, f"{pedido[campos_longe[i]]}", 1, 0, 'L')
        pdf.cell(30, 6, f"{campos_perto[i].replace('_', ' ')}:", 1, 0, 'L')
        pdf.cell(30, 6, f"{pedido[campos_perto[i]]}", 1, 0, 'L')
        if i < len(dnp):  # Certifique-se de que o índice esteja dentro do intervalo para DNP
            pdf.cell(30, 6, f"{dnp[i].replace('_', ' ')}:", 1, 0, 'L')
            pdf.cell(30, 6, f"{pedido[dnp[i]]}", 1, 1, 'L')
        else:
            pdf.cell(30, 6, "", 1, 0, 'L')
            pdf.cell(30, 6, "", 1, 1, 'L')

    # Espaçamento para os novos campos (space-around style)
    pdf.ln(4)  # Espaço extra após a tabela de detalhes do laboratório

    pdf.set_font('Arial', 'B', 8)
    pdf.cell(90, 6, f"Lentes: {pedido['lentes']}", 1, 0, 'L')
    pdf.cell(90, 6, f"Nome: {pedido['nome']}", 1, 1, 'L')

    pdf.cell(90, 6, f"Entregue: {pedido['Entregue']}", 1, 0, 'L')
    pdf.cell(90, 6, f"Observação: {pedido['Obs']}", 1, 1, 'L')

    pdf.cell(90, 6, f"Entregue Data/Hora: {pedido['Entrega_lab']}", 1, 1, 'L')

    # Nome do arquivo PDF será o ID do pedido
    nome_arquivo = f"pedido_{pedido['id']}.pdf"
    pdf.output(nome_arquivo)
    print(f'PDF gerado com sucesso!\nNome do arquivo: {nome_arquivo}')

if __name__ == "__main__":
    iniciar_historico_cliente()

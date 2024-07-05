import PySimpleGUI as sg
import pandas as pd
from datetime import datetime
from fpdf import FPDF

# Função para carregar dados do Excel
def carregar_pedidos():
    try:
        df = pd.read_excel('pedidos.xlsx')
        df['id'] = df['id'].astype(int)
        return df
    except FileNotFoundError:
        colunas = ['id', 'data_saida', 'data_entrega', 'nome', 'telefone', 'cpf', 'lentes', 'referencia', 'valor', 'p_conta', 'resta', 'garantia', 
                   'os', 'vendedor', 'Longe_ESF_OD', 'Longe_CIL_OD', 'Longe_EIXO_OD', 'Longe_ESF_OE', 'Longe_CIL_OE', 'Longe_EIXO_OE',
                   'Perto_ESF_OD', 'Perto_CIL_OD', 'Perto_EIXO_OD', 'Perto_ESF_OE', 'Perto_CIL_OE', 'Perto_EIXO_OE', 'DNP_OD', 'DNP_OE', 
                   'DNP_ALT', 'DNP_COR', 'Lente', 'Nome', 'Entregue', 'Obs', 'Entrega_lab']
        return pd.DataFrame(columns=colunas)

# Função para salvar dados no Excel
def salvar_pedidos(df):
    df.to_excel('pedidos.xlsx', index=False)

# Layout para cadastro de pedidos
def layout_pedido(values=None):
    if values is None:
        values = {}
    col1 = [
        [sg.Text('ID'), sg.Input(key='id', default_text=values.get('id', ''))],
        [sg.Text('Data de Saída'), sg.Input(key='data_saida', default_text=values.get('data_saida', ''))],
        [sg.Text('Data de Entrega'), sg.Input(key='data_entrega', default_text=values.get('data_entrega', ''))],
        [sg.Text('Nome'), sg.Input(key='nome', default_text=values.get('nome', ''))],
        [sg.Text('Telefone'), sg.Input(key='telefone', default_text=values.get('telefone', ''))],
        [sg.Text('CPF'), sg.Input(key='cpf', default_text=values.get('cpf', ''))],
        [sg.Text('Lentes'), sg.Input(key='lentes', default_text=values.get('lentes', ''))],
        [sg.Text('Referência'), sg.Input(key='referencia', default_text=values.get('referencia', ''))],
        [sg.Text('Valor'), sg.Input(key='valor', default_text=values.get('valor', ''))],
        [sg.Text('P/ Conta'), sg.Input(key='p_conta', default_text=values.get('p_conta', ''))],
        [sg.Text('Resta'), sg.Input(key='resta', default_text=values.get('resta', ''))],
        [sg.Text('Garantia'), sg.Input(key='garantia', default_text=values.get('garantia', ''))]
    ]
    
    col2 = [
        [sg.Text('OS'), sg.Input(key='os', default_text=values.get('os', ''))],
        [sg.Text('Vendedor'), sg.Input(key='vendedor', default_text=values.get('vendedor', ''))],
        [sg.Text('Longe ESF OD'), sg.Input(key='Longe_ESF_OD', default_text=values.get('Longe_ESF_OD', ''))],
        [sg.Text('Longe CIL OD'), sg.Input(key='Longe_CIL_OD', default_text=values.get('Longe_CIL_OD', ''))],
        [sg.Text('Longe EIXO OD'), sg.Input(key='Longe_EIXO_OD', default_text=values.get('Longe_EIXO_OD', ''))],
        [sg.Text('Longe ESF OE'), sg.Input(key='Longe_ESF_OE', default_text=values.get('Longe_ESF_OE', ''))],
        [sg.Text('Longe CIL OE'), sg.Input(key='Longe_CIL_OE', default_text=values.get('Longe_CIL_OE', ''))],
        [sg.Text('Longe EIXO OE'), sg.Input(key='Longe_EIXO_OE', default_text=values.get('Longe_EIXO_OE', ''))],
        [sg.Text('Perto ESF OD'), sg.Input(key='Perto_ESF_OD', default_text=values.get('Perto_ESF_OD', ''))],
        [sg.Text('Perto CIL OD'), sg.Input(key='Perto_CIL_OD', default_text=values.get('Perto_CIL_OD', ''))],
        [sg.Text('Perto EIXO OD'), sg.Input(key='Perto_EIXO_OD', default_text=values.get('Perto_EIXO_OD', ''))],
        [sg.Text('Perto ESF OE'), sg.Input(key='Perto_ESF_OE', default_text=values.get('Perto_ESF_OE', ''))],
        [sg.Text('Perto CIL OE'), sg.Input(key='Perto_CIL_OE', default_text=values.get('Perto_CIL_OE', ''))],
        [sg.Text('Perto Eixo OE'), sg.Input(key='Perto_EIXO_OE', default_text=values.get('Perto_EIXO_OE', ''))]
    ]
    
    col3 = [
        [sg.Text('DNP OD'), sg.Input(key='DNP_OD', default_text=values.get('DNP_OD', ''))],
        [sg.Text('DNP OE'), sg.Input(key='DNP_OE', default_text=values.get('DNP_OE', ''))],
        [sg.Text('DNP ALT'), sg.Input(key='DNP_ALT', default_text=values.get('DNP_ALT', ''))],
        [sg.Text('DNP COR'), sg.Input(key='DNP_COR', default_text=values.get('DNP_COR', ''))],
        [sg.Text('Lente'), sg.Input(key='Lente', enable_events=True, default_text=values.get('Lente', ''))],  # Campo para atualizar automaticamente
        [sg.Text('Entregue'), sg.Input(key='Entregue', default_text=values.get('Entregue', ''))],
        [sg.Text('Observações'), sg.Input(key='Obs', default_text=values.get('Obs', ''))],
        [sg.Text('Entrega do laboratório/hora'), sg.Input(key='Entrega_lab', default_text=values.get('Entrega_lab', ''))],
        [sg.Button('Salvar'), sg.Button('Cancelar')]
    ]

    layout = [
        [sg.Column(col1), sg.Column(col2), sg.Column(col3)]
    ]
    
    return layout

# Função para iniciar o sistema com menu
def iniciar_sistema():
    layout = [
        [sg.Button('Criar Pedido')],
        [sg.Button('Editar Pedido')],
        [sg.Button('Deletar Pedido')],
        [sg.Button('Imprimir Pedido')],
        [sg.Button('Sair')]
    ]
    window = sg.Window('Sistema de Pedidos', layout)

    while True:
        event, _ = window.read()
        if event == 'Criar Pedido':
            criar_pedido()
        elif event == 'Editar Pedido':
            editar_pedido()
        elif event == 'Deletar Pedido':
            deletar_pedido()
        elif event == 'Imprimir Pedido':
            imprimir_pedido()
        elif event == sg.WIN_CLOSED or event == 'Sair':
            break

    window.close()

# Função para criar um novo pedido
def criar_pedido():
    df_pedidos = carregar_pedidos()
    if not df_pedidos.empty:
        next_id = df_pedidos['id'].max() + 1
    else:
        next_id = 1
    
    layout = layout_pedido()  # Chama a função layout_pedido sem passar values
    window = sg.Window('Criar Pedido', layout, resizable=True, finalize=True)
    window.maximize()

    # Preencher o campo ID automaticamente
    window['id'].update(next_id)  # Atualiza o campo 'id' com o valor de next_id
    window['id'].update(disabled=True)  # Desabilita a edição do campo 'id'

    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED or event == 'Cancelar':
            break
        if event == 'Salvar':
            values['id'] = next_id
            for key, value in values.items():
                if isinstance(value, str) and value.strip() == '':
                    values[key] = None
            df_pedidos = pd.concat([df_pedidos, pd.DataFrame([values])], ignore_index=True)
            df_pedidos['id'] = df_pedidos['id'].astype(int)  # Garantir que o ID seja sempre int
            salvar_pedidos(df_pedidos)
            sg.popup('Pedido salvo com sucesso!')
            break

    window.close()

# Função para editar um pedido existente
def editar_pedido():
    df_pedidos = carregar_pedidos()
    layout_busca = [
        [sg.Text('ID do pedido a ser editado'), sg.Input(key='id')],
        [sg.Button('Buscar'), sg.Button('Cancelar')]
    ]
    window_busca = sg.Window('Editar Pedido', layout_busca)
    
    while True:
        event, values = window_busca.read()
        if event == sg.WIN_CLOSED or event == 'Cancelar':
            break
        if event == 'Buscar':
            pedido_id = values['id']
            if pd.notna(pedido_id) and str(pedido_id).strip() != '':
                pedido_id = str(pedido_id).strip()
                if pedido_id in df_pedidos['id'].astype(str).values:
                    pedido = df_pedidos[df_pedidos['id'].astype(str) == pedido_id].iloc[0]
                    window_busca.close()
                    
                    layout_edicao = layout_pedido(pedido)  # Passa os valores do pedido encontrado para o layout de edição
                    window_edicao = sg.Window('Editar Pedido', layout_edicao, resizable=True, finalize=True)
                    window_edicao.maximize()
                    
                    while True:
                        event, values = window_edicao.read()
                        if event == sg.WIN_CLOSED or event == 'Cancelar':
                            window_edicao.close()  # Adiciona o fechamento da janela aqui
                            break
                        if event == 'Salvar':
                            # Atualiza os valores do pedido no DataFrame
                            for key in values:
                                if key != 'id':  # Não tentar alterar o id
                                    if pd.isna(values[key]):
                                        df_pedidos.loc[df_pedidos['id'].astype(str) == pedido_id, key] = None
                                    else:
                                        df_pedidos.loc[df_pedidos['id'].astype(str) == pedido_id, key] = values[key]
                            df_pedidos['id'] = df_pedidos['id'].astype(int)  # Garantir que o ID seja sempre int
                            salvar_pedidos(df_pedidos)
                            sg.popup('Pedido editado com sucesso!')
                            window_edicao.close()
                            break
                else:
                    sg.popup('ID do pedido não encontrado. Tente novamente.')
                    continue
    window_busca.close()


# Função para deletar um pedido existente
def deletar_pedido():
    df_pedidos = carregar_pedidos()
    layout = [
        [sg.Text('ID do pedido a ser deletado'), sg.Input(key='id')],
        [sg.Button('Deletar'), sg.Button('Cancelar')]
    ]
    window = sg.Window('Deletar Pedido', layout)
    event, values = window.read()
    if event == 'Deletar':
        pedido_id = values['id']
        df_pedidos = df_pedidos[df_pedidos['id'].astype(str) != pedido_id]
        df_pedidos['id'] = df_pedidos['id'].astype(int)  # Garantir que o ID seja sempre int
        salvar_pedidos(df_pedidos)
        sg.popup('Pedido deletado com sucesso!')
    window.close()

# Função para imprimir um pedido existente em PDF
def imprimir_pedido():
    df_pedidos = carregar_pedidos()
    layout = [
        [sg.Text('ID do pedido a ser impresso'), sg.Input(key='id')],
        [sg.Button('Imprimir'), sg.Button('Cancelar')]
    ]
    window = sg.Window('Imprimir Pedido', layout)
    event, values = window.read()
    if event == 'Imprimir':
        pedido_id = values['id']
        if pd.notna(pedido_id) and str(pedido_id).strip() != '':
            pedido_id = str(pedido_id).strip()
            if pedido_id in df_pedidos['id'].astype(str).values:
                pedido = df_pedidos[df_pedidos['id'].astype(str) == pedido_id].iloc[0]
                gerar_pdf_estilizado(pedido)  # Ajuste: chama gerar_pdf_estilizado apenas com o pedido
                sg.popup('Pedido impresso com sucesso!')
            else:
                sg.popup('ID do pedido não encontrado. Tente novamente.')
    window.close()


class PDF(FPDF):
    def __init__(self, pedido):
        super().__init__()
        self.pedido = pedido
        self.WIDTH = 210  # largura da página em mm (A4)
        self.HEIGHT = 297  # altura da página em mm (A4)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'@Oticanewlook', 0, 0, 'C')

# Função para gerar um PDF estilizado com os dados de um pedido específico
def gerar_pdf_estilizado(pedido):
    pdf = PDF(pedido)
    pdf.add_page()

    # Informações do laboratório
    pdf.image('logo.png', x=10, y=pdf.get_y(), w=15)  # Inserir logo
    pdf.set_font('Arial', 'B', 8)

    # Espaço antes de iniciar os elementos
# Dimensões das imagens e textos
    wpp_text = '81996615539'
    inst_text = '@oticanewlook'
    icon_width = 5  # largura das imagens
    icon_text_padding = 2  # espaço entre a imagem e o texto
    icon_y_offset = 7.5  # deslocamento vertical das imagens
    horizontal_offset = 10  # deslocamento horizontal adicional
    total_width = (icon_width + icon_text_padding + pdf.get_string_width(wpp_text) +
                10 +  # espaço entre WhatsApp e Instagram
                icon_width + icon_text_padding + pdf.get_string_width(inst_text) +
                10 +  # espaço entre Instagram e número da OS
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

    # Espaço entre o username do Instagram e o número da OS
    pdf.set_x(pdf.get_x() + 5)

    # Número da OS
    pdf.cell(0, 15, f'OS: {pedido["id"]}', 0, 1, 'R')  # Adiciona nova linha após o número da OS

    pdf.set_y(pdf.get_y() - 10)  # Ajuste o valor (-10) conforme necessário para o posicionamento desejado

    # Nome do vendedor
    pdf.cell(0, 15, f'Vendedor: {pedido["vendedor"]}', 0, 1, 'R')  # Adiciona nova linha após o nome do vendedor

    pdf.ln(2)  # Espaçamento opcional após o nome do vendedor

    
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

    pdf.set_font('Arial', 'B', 8)
    # Detalhes do pedido
    pdf.cell(0, 6, f"Lentes: {pedido['lentes']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Referência: {pedido['referencia']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Valor: {pedido['valor']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"P/ Conta: {pedido['p_conta']}", 1, 1, 'L', 0)
    pdf.cell(0, 6, f"Resta: {pedido['resta']}", 1, 1, 'L', 0)
    
    #Garantia
    pdf.set_font('Arial', 'B', 8)
    pdf.cell(0, 8, 'GARANTIA:', 0, 1, 'C')

    # Posição para o texto de garantia
    x = pdf.get_x()
    y = pdf.get_y()

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

    """
        # Define a posição inicial e dimensões da caixa para o texto de garantia
    x = 10
    y = 10
    largura_caixa = pdf.w - 20  # Largura disponível na página menos as margens
    altura_caixa = 52

    # Desenha a caixa com bordas simples
    pdf.set_fill_color(230, 230, 230)  # Cor de fundo da caixa
    pdf.set_draw_color(100, 100, 100)  # Cor da borda da caixa
    pdf.rect(x, y, largura_caixa, altura_caixa, 'FD')

    # Configurações para o texto dentro da caixa
    pdf.set_xy(x + 5, y + 5)  # Posiciona o texto dentro da caixa
    pdf.set_font('Arial', 'I', 8)  # Define a fonte e tamanho

    # Ajuste o espaçamento entre as linhas
    cell_height = 6  # Altura da célula para controlar o espaçamento entre linhas
    pdf.multi_cell(largura_caixa - 10, cell_height, garantia_texto, 0, 'L', fill=False)

    # Espaço após o texto da garantia
    pdf.ln(5)

    """ 



    # Garantia
    pdf.cell(0, 8, "Assinatura do Cliente:", 0, 1, 'C')
    pdf.multi_cell(0, 8, "___________________________________", 0, 'C')
    pdf.ln(5)

    # Informações do laboratório
    pdf.image('logo.png', x=10, y=pdf.get_y(), w=15)  # Inserir logo
    pdf.set_font('Arial', 'B', 8)

    # Espaço antes de iniciar os elementos
# Dimensões das imagens e textos
    wpp_text = '81996615539'
    inst_text = '@oticanewlook'
    icon_width = 5  # largura das imagens
    icon_text_padding = 2  # espaço entre a imagem e o texto
    icon_y_offset = 7.5  # deslocamento vertical das imagens
    horizontal_offset = 10  # deslocamento horizontal adicional
    total_width = (icon_width + icon_text_padding + pdf.get_string_width(wpp_text) +
                10 +  # espaço entre WhatsApp e Instagram
                icon_width + icon_text_padding + pdf.get_string_width(inst_text) +
                10 +  # espaço entre Instagram e número da OS
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


# Função para imprimir um pedido existente em PDF estilizado
def imprimir_pedido_estilizado():
    df_pedidos = carregar_pedidos()
    layout = [
        [sg.Text('ID do pedido a ser impresso'), sg.Input(key='id')],
        [sg.Button('Imprimir'), sg.Button('Cancelar')]
    ]
    window = sg.Window('Imprimir Pedido Estilizado', layout)
    event, values = window.read()
    if event == 'Imprimir':
        pedido_id = values['id']
        if pd.notna(pedido_id) and str(pedido_id).strip() != '':
            pedido_id = str(pedido_id).strip()
            if pedido_id in df_pedidos['id'].astype(str).values:
                pedido = df_pedidos[df_pedidos['id'].astype(str) == pedido_id].iloc[0]
                gerar_pdf_estilizado(pedido)
            else:
                sg.popup('ID do pedido não encontrado. Tente novamente.')
    window.close()

if __name__ == '__main__':
    iniciar_sistema()

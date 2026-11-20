import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import date, datetime
from docx import Document
import os
import json
import re
import subprocess

# Define a aparência inicial da janela.
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

# Cria a janela principal do AutoDocs.
app = ctk.CTk()
app.title("AutoDocs")
app.geometry("1000x700")
app.minsize(800, 600)

# Configura o espaço usado pelas áreas da interface.
app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

# Define os caminhos dos arquivos e pastas do projeto.
PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_CONFIG = os.path.join(PASTA_PROJETO, "config.json")
ARQUIVO_DADOS = os.path.join(PASTA_PROJETO, "dados_contrato.json")
PASTA_DOCUMENTOS = os.path.join(PASTA_PROJETO, "Documentos_Gerados")

# Guarda o caminho do modelo Word selecionado nas configurações.
caminho_modelo = ""


# Carrega do JSON o caminho do modelo Word salvo anteriormente.
def carregar_configuracoes():
    global caminho_modelo

    try:
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as arquivo:
            configuracoes = json.load(arquivo)
            caminho_modelo = configuracoes.get("caminho_modelo", "")

    except (FileNotFoundError, json.JSONDecodeError):
        caminho_modelo = ""


# Salva no JSON o caminho do modelo Word selecionado.
def salvar_configuracoes():
    configuracoes = {"caminho_modelo": caminho_modelo}

    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as arquivo:
        json.dump(configuracoes, arquivo, indent=4, ensure_ascii=False)


# Salva os dados preenchidos no formulário em um arquivo JSON.
def salvar_dados_contrato(dados):
    with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, indent=4, ensure_ascii=False)


# Lê e retorna os dados do contrato que estão no arquivo JSON.
def carregar_dados_contrato():
    with open(ARQUIVO_DADOS, "r", encoding="utf-8") as arquivo:
        return json.load(arquivo)


# Converte um valor digitado no formato brasileiro para número decimal.
def converter_valor(valor):
    valor = valor.strip().replace("R$", "").strip()

    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    return float(valor) if valor else 0


# Formata um número como moeda brasileira (por exemplo, R$ 2.000,00).
def formatar_moeda(valor):
    return (
        f"R$ {valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


# Verifica se o CPF possui 11 dígitos e os dígitos verificadores corretos.
def validar_cpf(cpf):
    cpf = "".join(filter(str.isdigit, cpf))

    if len(cpf) != 11 or cpf == cpf[0] * 11:
        return False

    soma = sum(int(cpf[i]) * (10 - i) for i in range(9))
    digito1 = (soma * 10 % 11) % 10

    if digito1 != int(cpf[9]):
        return False

    soma = sum(int(cpf[i]) * (11 - i) for i in range(10))
    digito2 = (soma * 10 % 11) % 10

    return digito2 == int(cpf[10])


# Confere se a data está no formato DD/MM/AAAA.
def validar_data(data_texto):
    return datetime.strptime(data_texto, "%d/%m/%Y")


# Formata o telefone conforme a quantidade de dígitos.
def formatar_telefone(telefone):
    telefone = "".join(filter(str.isdigit, telefone))

    if len(telefone) == 11:
        return f"({telefone[:2]}) {telefone[2:7]}-{telefone[7:]}"

    if len(telefone) == 10:
        return f"({telefone[:2]}) {telefone[2:6]}-{telefone[6:]}"

    return telefone


# Substitui um identificador no parágrafo, mesmo quando o Word divide
# o texto entre diferentes partes (runs).
def substituir_em_paragrafo(paragrafo, identificador, valor):
    substituicoes = 0

    while True:
        texto_completo = "".join(run.text for run in paragrafo.runs)

        posicao = texto_completo.find(identificador)

        if posicao == -1:
            break

        inicio = posicao
        fim = posicao + len(identificador) - 1

        posicao_inicio = None
        posicao_fim = None

        acumulado = 0

        # Descobre em quais partes do parágrafo o identificador está.
        for indice, run in enumerate(paragrafo.runs):
            tamanho = len(run.text)

            if tamanho == 0:
                continue

            if posicao_inicio is None and inicio < acumulado + tamanho:
                posicao_inicio = (
                    indice,
                    inicio - acumulado
                )

            if fim < acumulado + tamanho:
                posicao_fim = (
                    indice,
                    fim - acumulado
                )
                break

            acumulado += tamanho

        if posicao_inicio is None or posicao_fim is None:
            break

        indice_inicio, deslocamento_inicio = posicao_inicio
        indice_fim, deslocamento_fim = posicao_fim

        # Se o identificador estiver em uma única parte, faz a troca nela.
        if indice_inicio == indice_fim:
            run = paragrafo.runs[indice_inicio]

            run.text = (
                run.text[:deslocamento_inicio]
                + str(valor)
                + run.text[deslocamento_fim + 1:]
            )

        # Se estiver dividido, junta a substituição e limpa as partes antigas.
        else:
            primeiro_run = paragrafo.runs[indice_inicio]
            ultimo_run = paragrafo.runs[indice_fim]

            primeiro_run.text = (
                primeiro_run.text[:deslocamento_inicio]
                + str(valor)
            )

            for indice in range(indice_inicio + 1, indice_fim):
                paragrafo.runs[indice].text = ""

            ultimo_run.text = ultimo_run.text[deslocamento_fim + 1:]

        substituicoes += 1

    return substituicoes


# Percorre os parágrafos e as tabelas do documento, inclusive tabelas
# que estejam dentro de outras tabelas.
def percorrer_paragrafos(paragrafos, tabelas):
    for paragrafo in paragrafos:
        yield paragrafo

    for tabela in tabelas:
        for linha in tabela.rows:
            for celula in linha.cells:
                yield from percorrer_paragrafos(
                    celula.paragraphs,
                    celula.tables
                )


# Reúne os parágrafos do corpo, dos cabeçalhos e dos rodapés do Word.
def obter_todos_paragrafos(documento):
    yield from percorrer_paragrafos(
        documento.paragraphs,
        documento.tables
    )

    for secao in documento.sections:
        yield from percorrer_paragrafos(
            secao.header.paragraphs,
            secao.header.tables
        )

        yield from percorrer_paragrafos(
            secao.footer.paragraphs,
            secao.footer.tables
        )


# Abre o modelo Word, preenche os identificadores e salva uma cópia.
def gerar_contrato_word():
    if not caminho_modelo or not os.path.isfile(caminho_modelo):
        raise FileNotFoundError(
            "Nenhum modelo Word válido foi configurado. "
            "Verifique as Configurações do AutoDocs."
        )

    dados = carregar_dados_contrato()

    documento = Document(caminho_modelo)

    # Relaciona cada identificador do modelo Word ao dado correspondente.
    substituicoes = {
        "{{LOCATARIO}}": dados["locatario"],
        "{{CPF}}": dados["cpf"],
        "{{TELEFONE}}": formatar_telefone(dados["telefone"]),
        "{{ENDERECO}}": dados["endereco"],
        "{{CHECKIN}}": dados["checkin"],
        "{{CHECKOUT}}": dados["checkout"],
        "{{HOSPEDES}}": str(dados["hospedes"]),
        "{{CONVIDADOS}}": str(dados["convidados"]),
        "{{MOTIVO_EVENTO}}": dados["motivo_evento"],
        "{{VALOR_TOTAL}}": formatar_moeda(dados["valor_total"]),
        "{{VALOR_RESERVA}}": formatar_moeda(dados["valor_reserva"]),
        "{{SALDO}}": formatar_moeda(dados["saldo"]),
        "{{DATA_EMISSAO}}": dados["data_emissao"]
    }

    # Conta quantas vezes cada identificador foi encontrado no modelo.
    encontrados = {identificador: 0 for identificador in substituicoes}

    for paragrafo in obter_todos_paragrafos(documento):
        for identificador, valor in substituicoes.items():

            quantidade = substituir_em_paragrafo(
                paragrafo,
                identificador,
                valor
            )

            encontrados[identificador] += quantidade

    # Impede a geração caso algum identificador esperado não exista no modelo.
    identificadores_ausentes = [
        identificador
        for identificador, quantidade in encontrados.items()
        if quantidade == 0
    ]

    if identificadores_ausentes:
        raise ValueError(
            "Os seguintes identificadores não foram encontrados "
            "no modelo Word:\n\n"
            + "\n".join(identificadores_ausentes)
        )

    # Cria a pasta de saída caso ela ainda não exista.
    os.makedirs(PASTA_DOCUMENTOS, exist_ok=True)

    # Limpa caracteres que não podem ser usados no nome do arquivo.
    nome_seguro = re.sub(
        r'[<>:"/\\|?*]',
        "",
        dados["locatario"]
    ).strip()

    horario = datetime.now().strftime("%Y%m%d_%H%M%S")

    nome_arquivo = f"Contrato_{nome_seguro}_{horario}.docx"

    caminho_saida = os.path.join(
        PASTA_DOCUMENTOS,
        nome_arquivo
    )

    documento.save(caminho_saida)

    return caminho_saida


# Converte o documento Word gerado para PDF usando o LibreOffice.
def converter_word_para_pdf(caminho_word):
    caminho_libreoffice = r"C:\Program Files\LibreOffice\program\soffice.exe"

    pasta_destino = os.path.dirname(caminho_word)

    # Monta o comando para executar a conversão sem abrir a janela do programa.
    comando = [
        caminho_libreoffice,
        "--headless",
        "--convert-to", "pdf",
        "--outdir", pasta_destino,
        caminho_word
    ]

    # Executa a conversão e aguarda no máximo 60 segundos.
    resultado = subprocess.run(
        comando,
        capture_output=True,
        text=True,
        timeout=60
    )

    caminho_pdf = os.path.splitext(caminho_word)[0] + ".pdf"

    # Confere se o LibreOffice terminou sem erro e criou o arquivo PDF.
    if resultado.returncode != 0 or not os.path.exists(caminho_pdf):
        raise RuntimeError(
            f"Erro ao converter contrato para PDF:\n{resultado.stderr}"
        )

    return caminho_pdf


# Recupera as configurações salvas antes de montar a interface.
carregar_configuracoes()


# Cria o menu lateral com os botões de navegação.
menu_lateral = ctk.CTkFrame(app, width=200, corner_radius=0)
menu_lateral.grid(row=0, column=0, sticky="nsew")
menu_lateral.grid_propagate(False)

titulo_menu = ctk.CTkLabel(
    menu_lateral,
    text="AutoDocs",
    font=("Arial", 26, "bold")
)
titulo_menu.pack(pady=(30, 25))

btn_novo_documento = ctk.CTkButton(
    menu_lateral,
    text="Novo documento",
    command=lambda: mostrar_pagina("Novo documento")
)
btn_novo_documento.pack(fill="x", padx=15, pady=8)

btn_documentos = ctk.CTkButton(
    menu_lateral,
    text="Documentos",
    command=lambda: mostrar_pagina("Documentos")
)
btn_documentos.pack(fill="x", padx=15, pady=8)

btn_configuracoes = ctk.CTkButton(
    menu_lateral,
    text="Configurações",
    command=lambda: mostrar_pagina("Configurações")
)
btn_configuracoes.pack(fill="x", padx=15, pady=8)

# Área onde o conteúdo de cada página será exibido.
area_principal = ctk.CTkFrame(app, corner_radius=0)
area_principal.grid(
    row=0,
    column=1,
    sticky="nsew",
    padx=20,
    pady=20
)


# Monta a página do formulário para gerar um novo contrato.
def criar_pagina_novo_documento():

    formulario = ctk.CTkScrollableFrame(
        area_principal,
        corner_radius=10
    )
    formulario.pack(fill="both", expand=True, padx=10, pady=10)

    titulo = ctk.CTkLabel(
        formulario,
        text="Novo contrato de aluguel",
        font=("Arial", 26, "bold")
    )
    titulo.pack(anchor="w", padx=20, pady=(20, 5))

    descricao = ctk.CTkLabel(
        formulario,
        text="Preencha os dados para gerar o documento.",
        font=("Arial", 14)
    )
    descricao.pack(anchor="w", padx=20, pady=(0, 20))

    # Preenche os campos com informações fictícias para facilitar os testes.
    def preencher_dados_teste():
        dados_teste = {
            "nome": "João da Silva",
            "cpf": "529.982.247-25",
            "endereco": "Rua das Flores, 123, Centro, Pelotas/RS",
            "telefone": "53999999999",
            "hospedes": "2",
            "convidados": "5",
            "motivo_evento": "Aniversário",
            "checkin": "20/12/2026",
            "checkout": "25/12/2026",
            "valor_total": "2000,00",
            "valor_reserva": "500,00"
        }

        campos = [
            campo_nome,
            campo_cpf,
            campo_endereco,
            campo_telefone,
            campo_hospedes,
            campo_convidados,
            campo_motivo_evento,
            campo_checkin,
            campo_checkout,
            campo_valor_total,
            campo_reserva
        ]

        valores = list(dados_teste.values())

        # Insere cada dado de teste no campo correspondente.
        for campo, valor in zip(campos, valores):
            campo.delete(0, "end")
            campo.insert(0, valor)

        atualizar_saldo()

    btn_dados_teste = ctk.CTkButton(
        formulario,
        text="Preencher dados de teste",
        height=40,
        fg_color="#455A64",
        hover_color="#37474F",
        command=preencher_dados_teste
    )

    btn_dados_teste.pack(fill="x", padx=20, pady=(5, 15))

    secao_locatario = ctk.CTkLabel(
        formulario,
        text="Dados da locatária",
        font=("Arial", 18, "bold")
    )
    secao_locatario.pack(anchor="w", padx=20, pady=(10, 10))

    campo_nome = ctk.CTkEntry(
        formulario,
        placeholder_text="Nome completo da locatária",
        height=40
    )
    campo_nome.pack(fill="x", padx=20, pady=6)

    campo_cpf = ctk.CTkEntry(
        formulario,
        placeholder_text="CPF",
        height=40
    )
    campo_cpf.pack(fill="x", padx=20, pady=6)

    campo_endereco = ctk.CTkEntry(
        formulario,
        placeholder_text="Endereço completo da locatária",
        height=40
    )
    campo_endereco.pack(fill="x", padx=20, pady=6)

    linha_contato = ctk.CTkFrame(
        formulario,
        fg_color="transparent"
    )
    linha_contato.pack(fill="x", padx=20, pady=6)
    linha_contato.grid_columnconfigure((0, 1), weight=1)

    campo_telefone = ctk.CTkEntry(
        linha_contato,
        placeholder_text="Telefone",
        height=40
    )
    campo_telefone.grid(row=0, column=0, sticky="ew", padx=(0, 5))

    campo_hospedes = ctk.CTkEntry(
        linha_contato,
        placeholder_text="Hóspedes para pernoite",
        height=40
    )
    campo_hospedes.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    campo_convidados = ctk.CTkEntry(
        formulario,
        placeholder_text="Quantidade de convidados do evento",
        height=40
    )
    campo_convidados.pack(fill="x", padx=20, pady=6)

    campo_motivo_evento = ctk.CTkEntry(
        formulario,
        placeholder_text="Motivo do evento (ex: Comemoração de aniversário)",
        height=40
    )
    campo_motivo_evento.pack(fill="x", padx=20, pady=6)

    secao_periodo = ctk.CTkLabel(
        formulario,
        text="Período do aluguel",
        font=("Arial", 18, "bold")
    )
    secao_periodo.pack(anchor="w", padx=20, pady=(25, 10))

    linha_datas = ctk.CTkFrame(
        formulario,
        fg_color="transparent"
    )
    linha_datas.pack(fill="x", padx=20, pady=6)
    linha_datas.grid_columnconfigure((0, 1), weight=1)

    campo_checkin = ctk.CTkEntry(
        linha_datas,
        placeholder_text="Data de entrada (DD/MM/AAAA)",
        height=40
    )
    campo_checkin.grid(row=0, column=0, sticky="ew", padx=(0, 5))

    campo_checkout = ctk.CTkEntry(
        linha_datas,
        placeholder_text="Data de saída (DD/MM/AAAA)",
        height=40
    )
    campo_checkout.grid(row=0, column=1, sticky="ew", padx=(5, 0))

    secao_valores = ctk.CTkLabel(
        formulario,
        text="Valores do aluguel",
        font=("Arial", 18, "bold")
    )
    secao_valores.pack(anchor="w", padx=20, pady=(25, 10))

    campo_valor_total = ctk.CTkEntry(
        formulario,
        placeholder_text="Valor total do aluguel (ex: 3000,00)",
        height=40
    )
    campo_valor_total.pack(fill="x", padx=20, pady=6)

    campo_reserva = ctk.CTkEntry(
        formulario,
        placeholder_text="Valor pago de entrada/sinal (ex: 500,00)",
        height=40
    )
    campo_reserva.pack(fill="x", padx=20, pady=6)

    # Exibe o saldo que ainda falta pagar.
    saldo_frame = ctk.CTkFrame(
        formulario,
        corner_radius=10
    )
    saldo_frame.pack(fill="x", padx=20, pady=10)

    texto_saldo = ctk.CTkLabel(
        saldo_frame,
        text="Saldo restante a pagar:",
        font=("Arial", 15)
    )
    texto_saldo.pack(side="left", padx=15, pady=15)

    label_saldo = ctk.CTkLabel(
        saldo_frame,
        text="R$ 0,00",
        font=("Arial", 18, "bold")
    )
    label_saldo.pack(side="right", padx=15, pady=15)

    # Calcula o saldo subtraindo o valor da reserva do valor total.
    def atualizar_saldo(event=None):
        try:
            total = converter_valor(campo_valor_total.get())
            reserva = converter_valor(campo_reserva.get())

            saldo = total - reserva

            label_saldo.configure(
                text=formatar_moeda(saldo)
            )

        except ValueError:
            label_saldo.configure(text="Valor inválido")

    # Atualiza o saldo sempre que um dos valores for alterado.
    campo_valor_total.bind("<KeyRelease>", atualizar_saldo)
    campo_reserva.bind("<KeyRelease>", atualizar_saldo)

    secao_modelo = ctk.CTkLabel(
        formulario,
        text="Modelo de contrato",
        font=("Arial", 18, "bold")
    )
    secao_modelo.pack(anchor="w", padx=20, pady=(25, 10))

    # Informa se já existe um modelo Word configurado.
    if caminho_modelo and os.path.isfile(caminho_modelo):
        texto_modelo = (
            f"Modelo configurado: {os.path.basename(caminho_modelo)}"
        )
    else:
        texto_modelo = (
            "Nenhum modelo configurado. Vá até Configurações."
        )

    label_modelo = ctk.CTkLabel(
        formulario,
        text=texto_modelo,
        font=("Arial", 14),
        wraplength=650
    )
    label_modelo.pack(anchor="w", padx=20, pady=5)

    secao_emissao = ctk.CTkLabel(
        formulario,
        text="Data de emissão",
        font=("Arial", 18, "bold")
    )
    secao_emissao.pack(anchor="w", padx=20, pady=(25, 10))

    # Usa a data atual como data de emissão do contrato.
    data_atual = date.today().strftime("%d/%m/%Y")

    label_data = ctk.CTkLabel(
        formulario,
        text=f"Data de emissão do contrato: {data_atual}",
        font=("Arial", 14)
    )
    label_data.pack(anchor="w", padx=20, pady=5)

    # Reúne os campos, verifica os dados e informa se houver algum problema.
    def validar_formulario():

        nome = campo_nome.get().strip()
        cpf = campo_cpf.get().strip()
        endereco = campo_endereco.get().strip()

        telefone = "".join(
            filter(str.isdigit, campo_telefone.get())
        )

        hospedes = campo_hospedes.get().strip()
        convidados = campo_convidados.get().strip()
        motivo_evento = campo_motivo_evento.get().strip()

        checkin = campo_checkin.get().strip()
        checkout = campo_checkout.get().strip()

        valor_total = campo_valor_total.get().strip()
        valor_reserva = campo_reserva.get().strip()

        erros = []

        # Verifica se os dados obrigatórios foram preenchidos corretamente.
        if not nome:
            erros.append("Informe o nome completo da locatária.")

        if not validar_cpf(cpf):
            erros.append("CPF inválido.")

        if not endereco:
            erros.append("Informe o endereço completo da locatária.")

        if len(telefone) not in (10, 11):
            erros.append("Telefone inválido. Informe DDD e número.")

        # Confere se a quantidade de hóspedes é um número maior que zero.
        try:
            quantidade_hospedes = int(hospedes)

            if quantidade_hospedes <= 0:
                erros.append(
                    "A quantidade de hóspedes deve ser maior que zero."
                )

        except ValueError:
            quantidade_hospedes = 0
            erros.append("Informe uma quantidade válida de hóspedes.")

        # Confere se a quantidade de convidados é um número não negativo.
        try:
            quantidade_convidados = int(convidados)

            if quantidade_convidados < 0:
                erros.append(
                    "A quantidade de convidados não pode ser negativa."
                )

        except ValueError:
            quantidade_convidados = 0
            erros.append("Informe uma quantidade válida de convidados.")

        if not motivo_evento:
            erros.append("Informe o motivo do evento.")

        # Verifica se as datas foram digitadas no formato esperado.
        try:
            data_entrada = validar_data(checkin)

        except ValueError:
            data_entrada = None
            erros.append("Data de entrada inválida. Use DD/MM/AAAA.")

        try:
            data_saida = validar_data(checkout)

        except ValueError:
            data_saida = None
            erros.append("Data de saída inválida. Use DD/MM/AAAA.")

        # A data de saída precisa ser posterior à data de entrada.
        if data_entrada and data_saida:
            if data_saida <= data_entrada:
                erros.append("A saída deve ser posterior à entrada.")

        # Converte e valida o valor total do aluguel.
        try:
            total = converter_valor(valor_total)

            if not valor_total:
                erros.append("Informe o valor total do aluguel.")

            elif total <= 0:
                erros.append("O valor total deve ser maior que zero.")

        except ValueError:
            total = 0
            erros.append("Valor total inválido.")

        # Converte e valida o valor pago como entrada ou sinal.
        try:
            reserva = converter_valor(valor_reserva)

            if not valor_reserva:
                erros.append("Informe o valor pago de entrada/sinal.")

            elif reserva < 0:
                erros.append(
                    "O valor da entrada/sinal não pode ser negativo."
                )

        except ValueError:
            reserva = 0
            erros.append("Valor pago de entrada/sinal inválido.")

        # A entrada não pode ser maior que o valor total do aluguel.
        if total > 0 and reserva > total:
            erros.append(
                "O valor da entrada/sinal não pode ser maior que o aluguel."
            )

        # Mostra todos os problemas encontrados e interrompe a geração.
        if erros:
            messagebox.showerror(
                "Verifique os dados",
                "\n\n".join(erros)
            )
            return

        # Organiza os dados validados que serão usados no contrato.
        dados_contrato = {
            "locatario": nome,
            "cpf": cpf,
            "endereco": endereco,
            "telefone": telefone,
            "hospedes": quantidade_hospedes,
            "convidados": quantidade_convidados,
            "motivo_evento": motivo_evento,
            "checkin": checkin,
            "checkout": checkout,
            "valor_total": total,
            "valor_reserva": reserva,
            "saldo": total - reserva,
            "data_emissao": date.today().strftime("%d/%m/%Y")
        }

        try:
            # Salva os dados, gera o Word, converte para PDF e apaga o Word temporário.
            salvar_dados_contrato(dados_contrato)

            caminho_word = gerar_contrato_word()
            caminho_pdf = converter_word_para_pdf(caminho_word)
            os.remove(caminho_word)

            messagebox.showinfo(
                "Contrato gerado!",
                "Os dados foram validados e o contrato foi gerado "
                "com sucesso!\n\n"
                f"Arquivo PDF salvo em:\n{caminho_pdf}"
            )

        # Exibe erros conhecidos, como modelo ausente ou dados incorretos.
        except (FileNotFoundError, ValueError, KeyError) as erro:
            messagebox.showerror(
                "Não foi possível gerar o contrato",
                str(erro)
            )

        # Captura outros erros que possam acontecer durante a geração.
        except Exception as erro:
            messagebox.showerror(
                "Erro inesperado",
                f"Ocorreu um erro durante a geração:\n{erro}"
            )

    btn_gerar = ctk.CTkButton(
        formulario,
        text="Validar dados e gerar contrato",
        height=45,
        font=("Arial", 16, "bold"),
        command=validar_formulario
    )
    btn_gerar.pack(fill="x", padx=20, pady=(20, 10))

    espaco_final = ctk.CTkLabel(formulario, text="")
    espaco_final.pack(pady=15)


# Monta a página usada para escolher o modelo de contrato Word.
def criar_pagina_configuracoes():

    frame_config = ctk.CTkFrame(
        area_principal,
        corner_radius=10
    )
    frame_config.pack(
        fill="both",
        expand=True,
        padx=10,
        pady=10
    )

    titulo = ctk.CTkLabel(
        frame_config,
        text="Configurações",
        font=("Arial", 26, "bold")
    )
    titulo.pack(anchor="w", padx=25, pady=(30, 10))

    descricao = ctk.CTkLabel(
        frame_config,
        text="Configure os arquivos padrão utilizados pelo AutoDocs.",
        font=("Arial", 14)
    )
    descricao.pack(anchor="w", padx=25, pady=(0, 25))

    secao_modelo = ctk.CTkLabel(
        frame_config,
        text="Modelo padrão de contrato",
        font=("Arial", 18, "bold")
    )
    secao_modelo.pack(anchor="w", padx=25, pady=(10, 10))

    label_modelo_config = ctk.CTkLabel(
        frame_config,
        text="",
        font=("Arial", 14),
        wraplength=650,
        justify="left"
    )
    label_modelo_config.pack(anchor="w", padx=25, pady=10)

    # Atualiza o texto da tela conforme a situação do modelo configurado.
    def atualizar_label_modelo():

        if caminho_modelo and os.path.isfile(caminho_modelo):
            label_modelo_config.configure(
                text=f"Modelo atual:\n{caminho_modelo}"
            )

        elif caminho_modelo:
            label_modelo_config.configure(
                text=(
                    "O arquivo configurado não foi encontrado.\n"
                    "Selecione um novo modelo."
                )
            )

        else:
            label_modelo_config.configure(
                text="Nenhum modelo de contrato foi selecionado."
            )

    # Abre a janela de arquivos e salva o modelo escolhido pelo usuário.
    def selecionar_modelo():
        global caminho_modelo

        arquivo = filedialog.askopenfilename(
            title="Selecionar modelo de contrato",
            filetypes=[
                ("Documentos Word", "*.docx")
            ]
        )

        if arquivo:
            caminho_modelo = arquivo
            salvar_configuracoes()
            atualizar_label_modelo()

            messagebox.showinfo(
                "Configuração salva",
                "O modelo padrão de contrato foi salvo com sucesso!"
            )

    btn_selecionar = ctk.CTkButton(
        frame_config,
        text="Selecionar modelo Word (.docx)",
        height=40,
        command=selecionar_modelo
    )
    btn_selecionar.pack(anchor="w", padx=25, pady=10)

    atualizar_label_modelo()


# Troca o conteúdo da área principal conforme a página escolhida no menu.
def mostrar_pagina(nome_pagina):

    # Remove os elementos da página anterior.
    for widget in area_principal.winfo_children():
        widget.destroy()

    if nome_pagina == "Novo documento":
        criar_pagina_novo_documento()
        return

    if nome_pagina == "Configurações":
        criar_pagina_configuracoes()
        return

    titulo = ctk.CTkLabel(
        area_principal,
        text=nome_pagina,
        font=("Arial", 26, "bold")
    )
    titulo.pack(pady=(40, 15))

    if nome_pagina == "Documentos":
        texto = "Aqui serão exibidos os contratos gerados."

    else:
        texto = "Bem-vindo ao AutoDocs!"

    descricao = ctk.CTkLabel(
        area_principal,
        text=texto,
        font=("Arial", 14)
    )
    descricao.pack(pady=10)


# Abre a tela inicial e mantém a aplicação em execução.
mostrar_pagina("Bem-vindo ao AutoDocs")

app.mainloop()
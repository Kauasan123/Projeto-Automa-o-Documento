import customtkinter as ctk
from tkinter import filedialog, messagebox
from datetime import date, datetime
import os
import json

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("AutoDocs")
app.geometry("1000x700")
app.minsize(800, 600)

app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

PASTA_PROJETO = os.path.dirname(os.path.abspath(__file__))
ARQUIVO_CONFIG = os.path.join(PASTA_PROJETO, "config.json")
ARQUIVO_DADOS = os.path.join(PASTA_PROJETO, "dados_contrato.json")

caminho_modelo = ""

def carregar_configuracoes():
    global caminho_modelo

    try:
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as arquivo:
            configuracoes = json.load(arquivo)
            caminho_modelo = configuracoes.get("caminho_modelo", "")

    except (FileNotFoundError, json.JSONDecodeError):
        caminho_modelo = ""


def salvar_configuracoes():
    configuracoes = {"caminho_modelo": caminho_modelo}

    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as arquivo:
        json.dump(configuracoes, arquivo, indent=4, ensure_ascii=False)


def salvar_dados_contrato(dados):
    with open(ARQUIVO_DADOS, "w", encoding="utf-8") as arquivo:
        json.dump(dados, arquivo, indent=4, ensure_ascii=False)

def converter_valor(valor):
    valor = valor.strip().replace("R$", "").strip()

    if "," in valor:
        valor = valor.replace(".", "").replace(",", ".")

    return float(valor) if valor else 0


def formatar_moeda(valor):
    return (
        f"R$ {valor:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


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


def validar_data(data_texto):
    return datetime.strptime(data_texto, "%d/%m/%Y")

carregar_configuracoes()

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

area_principal = ctk.CTkFrame(app, corner_radius=0)
area_principal.grid(
    row=0,
    column=1,
    sticky="nsew",
    padx=20,
    pady=20
)

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

    secao_locatario = ctk.CTkLabel(
        formulario,
        text="Dados do locatário",
        font=("Arial", 18, "bold")
    )
    secao_locatario.pack(anchor="w", padx=20, pady=(10, 10))

    campo_nome = ctk.CTkEntry(
        formulario,
        placeholder_text="Nome completo do locatário",
        height=40
    )
    campo_nome.pack(fill="x", padx=20, pady=6)

    linha_documentos = ctk.CTkFrame(
        formulario,
        fg_color="transparent"
    )
    linha_documentos.pack(fill="x", padx=20, pady=6)
    linha_documentos.grid_columnconfigure((0, 1), weight=1)

    campo_cpf = ctk.CTkEntry(
        linha_documentos,
        placeholder_text="CPF",
        height=40
    )
    campo_cpf.grid(row=0, column=0, sticky="ew", padx=(0, 5))

    campo_cnh = ctk.CTkEntry(
        linha_documentos,
        placeholder_text="Número da CNH",
        height=40
    )
    campo_cnh.grid(row=0, column=1, sticky="ew", padx=(5, 0))

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
        placeholder_text="Quantidade de hóspedes",
        height=40
    )
    campo_hospedes.grid(row=0, column=1, sticky="ew", padx=(5, 0))

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
        placeholder_text="Valor pago pela reserva (ex: 500,00)",
        height=40
    )
    campo_reserva.pack(fill="x", padx=20, pady=6)

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

    campo_valor_total.bind(
        "<KeyRelease>",
        atualizar_saldo
    )
    campo_reserva.bind(
        "<KeyRelease>",
        atualizar_saldo
    )

    secao_modelo = ctk.CTkLabel(
        formulario,
        text="Modelo de contrato",
        font=("Arial", 18, "bold")
    )
    secao_modelo.pack(anchor="w", padx=20, pady=(25, 10))

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

    data_atual = date.today().strftime("%d/%m/%Y")

    label_data = ctk.CTkLabel(
        formulario,
        text=f"Data de emissão do contrato: {data_atual}",
        font=("Arial", 14)
    )
    label_data.pack(anchor="w", padx=20, pady=5)

    def validar_formulario():

        nome = campo_nome.get().strip()
        cpf = campo_cpf.get().strip()
        cnh = campo_cnh.get().strip()
        telefone = "".join(
            filter(str.isdigit, campo_telefone.get())
        )
        hospedes = campo_hospedes.get().strip()
        checkin = campo_checkin.get().strip()
        checkout = campo_checkout.get().strip()
        valor_total = campo_valor_total.get().strip()
        valor_reserva = campo_reserva.get().strip()

        erros = []

        if not nome:
            erros.append(
                "Informe o nome completo do locatário."
            )

        if not validar_cpf(cpf):
            erros.append("CPF inválido.")

        cnh_digitos = "".join(filter(str.isdigit, cnh))

        if len(cnh_digitos) != 11:
            erros.append(
                "A CNH deve possuir 11 dígitos."
            )

        if len(telefone) not in (10, 11):
            erros.append(
                "Telefone inválido. Informe DDD e número."
            )

        try:
            quantidade = int(hospedes)

            if quantidade <= 0:
                erros.append(
                    "A quantidade de hóspedes deve ser maior que zero."
                )

        except ValueError:
            erros.append(
                "Informe uma quantidade válida de hóspedes."
            )

        try:
            data_entrada = validar_data(checkin)

        except ValueError:
            data_entrada = None
            erros.append(
                "Data de entrada inválida. Use DD/MM/AAAA."
            )

        try:
            data_saida = validar_data(checkout)

        except ValueError:
            data_saida = None
            erros.append(
                "Data de saída inválida. Use DD/MM/AAAA."
            )

        if data_entrada and data_saida:
            if data_saida <= data_entrada:
                erros.append(
                    "A saída deve ser posterior à entrada."
                )

        try:
            total = converter_valor(valor_total)

            if not valor_total:
                erros.append(
                    "Informe o valor total do aluguel."
                )

            elif total <= 0:
                erros.append(
                    "O valor total deve ser maior que zero."
                )

        except ValueError:
            total = 0
            erros.append("Valor total inválido.")

        try:
            reserva = converter_valor(valor_reserva)

            if not valor_reserva:
                erros.append(
                    "Informe o valor pago pela reserva."
                )

            elif reserva < 0:
                erros.append(
                    "O valor da reserva não pode ser negativo."
                )

        except ValueError:
            reserva = 0
            erros.append(
                "Valor pago pela reserva inválido."
            )

        if total > 0 and reserva > total:
            erros.append(
                "O valor da reserva não pode ser maior que o aluguel."
            )

        if erros:
            messagebox.showerror(
                "Verifique os dados",
                "\n\n".join(erros)
            )
            return

        dados_contrato = {
            "locatario": nome,
            "cpf": cpf,
            "cnh": cnh_digitos,
            "telefone": telefone,
            "hospedes": quantidade,
            "checkin": checkin,
            "checkout": checkout,
            "valor_total": total,
            "valor_reserva": reserva,
            "saldo": total - reserva,
            "data_emissao": date.today().strftime("%d/%m/%Y")
        }

        try:
            salvar_dados_contrato(dados_contrato)

            messagebox.showinfo(
                "Validação concluída",
                "Todos os dados foram validados e salvos com sucesso!"
            )

        except OSError as erro:
            messagebox.showerror(
                "Erro ao salvar",
                f"Não foi possível salvar os dados:\n{erro}"
            )

    btn_validar = ctk.CTkButton(
        formulario,
        text="Validar dados",
        height=45,
        font=("Arial", 16, "bold"),
        command=validar_formulario
    )
    btn_validar.pack(fill="x", padx=20, pady=(20, 10))

    espaco_final = ctk.CTkLabel(
        formulario,
        text=""
    )
    espaco_final.pack(pady=15)

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
    btn_selecionar.pack(
        anchor="w",
        padx=25,
        pady=10
    )

    atualizar_label_modelo()

def mostrar_pagina(nome_pagina):

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

mostrar_pagina("Bem-vindo ao AutoDocs")

app.mainloop()
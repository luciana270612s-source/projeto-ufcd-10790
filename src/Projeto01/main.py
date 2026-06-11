import tkinter as tk
from tkinter import ttk, messagebox
from openpyxl import load_workbook
from datetime import datetime
import os

# =========================
# CONFIGURAÇÕES DO PROJETO
# =========================
ARQUIVO_EXCEL = "estoque_limpeza.xlsx"
ABA_ESTOQUE = "Estoque"
ABA_MOVIMENTOS = "Movimentos"

# =========================
# CORES DA INTERFACE
# =========================
COR_FUNDO = "#F4F8F7"
COR_TOPO = "#1B5E20"
COR_BOTAO = "#2E7D32"
COR_BOTAO_HOVER = "#1B5E20"
COR_BRANCO = "#FFFFFF"
COR_TEXTO = "#1B1B1B"

# Linhas da tabela
COR_OK_FUNDO = "#DFF5E1"        # Verde claro
COR_OK_TEXTO = "#1B5E20"        # Verde escuro
COR_ALERTA_FUNDO = "#FFD6D6"    # Vermelho claro
COR_ALERTA_TEXTO = "#B00020"    # Vermelho escuro


def numero(valor):
    """Converte valores do Excel em número."""
    if valor is None:
        return 0
    if isinstance(valor, (int, float)):
        return valor

    texto = str(valor).replace("€", "").replace(",", ".").strip()

    try:
        return float(texto)
    except ValueError:
        return 0


def formatar_numero(valor):
    """Mostra números sem .0 quando forem inteiros."""
    valor = numero(valor)
    if valor == int(valor):
        return str(int(valor))
    return str(valor)


def verificar_excel():
    if not os.path.exists(ARQUIVO_EXCEL):
        messagebox.showerror(
            "Erro",
            f"O ficheiro {ARQUIVO_EXCEL} não foi encontrado.\n\n"
            "Coloca o Excel na mesma pasta do programa."
        )
        return False
    return True


def abrir_excel():
    if not verificar_excel():
        return None

    try:
        wb = load_workbook(ARQUIVO_EXCEL)
    except PermissionError:
        messagebox.showerror(
            "Erro",
            "Fecha o ficheiro Excel antes de usar o programa.\n"
            "O Python não consegue guardar alterações com o Excel aberto."
        )
        return None

    if ABA_ESTOQUE not in wb.sheetnames or ABA_MOVIMENTOS not in wb.sheetnames:
        messagebox.showerror(
            "Erro",
            "O Excel precisa ter as abas 'Estoque' e 'Movimentos'."
        )
        wb.close()
        return None

    return wb


def encontrar_linha_produto(ws, codigo):
    for linha in range(2, ws.max_row + 1):
        codigo_linha = ws.cell(row=linha, column=1).value
        if str(codigo_linha).strip() == str(codigo).strip():
            return linha
    return None


def recalcular_produto(wb, codigo):
    """Atualiza entradas, saídas, estoque atual e alerta COMPRAR/OK."""
    ws_estoque = wb[ABA_ESTOQUE]
    ws_movimentos = wb[ABA_MOVIMENTOS]

    linha_produto = encontrar_linha_produto(ws_estoque, codigo)
    if linha_produto is None:
        return

    estoque_inicial = numero(ws_estoque.cell(row=linha_produto, column=5).value)
    estoque_minimo = numero(ws_estoque.cell(row=linha_produto, column=9).value)

    entradas = 0
    saidas = 0
    ultima_compra = ""

    for linha in range(2, ws_movimentos.max_row + 1):
        codigo_movimento = ws_movimentos.cell(row=linha, column=2).value
        tipo = str(ws_movimentos.cell(row=linha, column=4).value).strip().upper()
        quantidade = numero(ws_movimentos.cell(row=linha, column=5).value)
        data_movimento = ws_movimentos.cell(row=linha, column=1).value

        if str(codigo_movimento).strip() != str(codigo).strip():
            continue

        if tipo in ["COMPRA", "AJUSTE +", "ENTRADA"]:
            entradas += quantidade
            if tipo == "COMPRA":
                ultima_compra = data_movimento

        elif tipo in ["RETIRADA", "AJUSTE -", "SAÍDA", "SAIDA"]:
            saidas += quantidade

    estoque_atual = estoque_inicial + entradas - saidas

    if estoque_atual <= estoque_minimo:
        status = "COMPRAR"
    else:
        status = "OK"

    # Colunas da aba Estoque
    ws_estoque.cell(row=linha_produto, column=6).value = entradas
    ws_estoque.cell(row=linha_produto, column=7).value = saidas
    ws_estoque.cell(row=linha_produto, column=8).value = estoque_atual
    ws_estoque.cell(row=linha_produto, column=10).value = status
    ws_estoque.cell(row=linha_produto, column=11).value = ultima_compra


def recalcular_todos_produtos():
    wb = abrir_excel()
    if wb is None:
        return False

    ws_estoque = wb[ABA_ESTOQUE]

    for linha in range(2, ws_estoque.max_row + 1):
        codigo = ws_estoque.cell(row=linha, column=1).value
        produto = ws_estoque.cell(row=linha, column=2).value

        if codigo is not None and produto is not None:
            recalcular_produto(wb, codigo)

    try:
        wb.save(ARQUIVO_EXCEL)
    except PermissionError:
        messagebox.showerror(
            "Erro",
            "Fecha o Excel antes de atualizar ou guardar alterações."
        )
        wb.close()
        return False

    wb.close()
    return True


def carregar_dados():
    if not verificar_excel():
        return []

    wb = abrir_excel()
    if wb is None:
        return []

    ws = wb[ABA_ESTOQUE]
    produtos = []

    for linha in range(2, ws.max_row + 1):
        codigo = ws.cell(row=linha, column=1).value
        produto = ws.cell(row=linha, column=2).value

        if codigo is None or produto is None:
            continue

        categoria = ws.cell(row=linha, column=3).value or ""
        unidade = ws.cell(row=linha, column=4).value or ""
        estoque_inicial = ws.cell(row=linha, column=5).value or 0
        entradas = ws.cell(row=linha, column=6).value or 0
        saidas = ws.cell(row=linha, column=7).value or 0
        estoque_atual = ws.cell(row=linha, column=8).value or 0
        estoque_minimo = ws.cell(row=linha, column=9).value or 0
        comprar = ws.cell(row=linha, column=10).value or ""
        observacoes = ws.cell(row=linha, column=12).value or ""

        produtos.append({
            "codigo": codigo,
            "produto": produto,
            "categoria": categoria,
            "unidade": unidade,
            "estoque_inicial": estoque_inicial,
            "entradas": entradas,
            "saidas": saidas,
            "estoque_atual": estoque_atual,
            "estoque_minimo": estoque_minimo,
            "comprar": comprar,
            "observacoes": observacoes
        })

    wb.close()
    return produtos


def atualizar_tabela(lista=None):
    for item in tabela.get_children():
        tabela.delete(item)

    if lista is None:
        recalcular_todos_produtos()
        produtos = carregar_dados()
    else:
        produtos = lista

    for produto in produtos:
        status = str(produto["comprar"]).upper().strip()

        if status == "COMPRAR":
            tag = "alerta"
        else:
            tag = "ok"

        tabela.insert(
            "",
            tk.END,
            values=(
                produto["codigo"],
                produto["produto"],
                produto["categoria"],
                produto["unidade"],
                formatar_numero(produto["estoque_atual"]),
                formatar_numero(produto["estoque_minimo"]),
                status
            ),
            tags=(tag,)
        )


def pesquisar_produto():
    termo = entrada_pesquisa.get().lower().strip()
    produtos = carregar_dados()

    resultado = []

    for produto in produtos:
        nome = str(produto["produto"]).lower()
        codigo = str(produto["codigo"]).lower()
        categoria = str(produto["categoria"]).lower()

        if termo in nome or termo in codigo or termo in categoria:
            resultado.append(produto)

    atualizar_tabela(resultado)


def mostrar_todos():
    entrada_pesquisa.delete(0, tk.END)
    atualizar_tabela()


def mostrar_produtos_para_comprar():
    produtos = carregar_dados()

    resultado = []
    for produto in produtos:
        if str(produto["comprar"]).upper().strip() == "COMPRAR":
            resultado.append(produto)

    atualizar_tabela(resultado)


def obter_produto_selecionado():
    selecionado = tabela.selection()

    if not selecionado:
        messagebox.showwarning("Aviso", "Seleciona um produto na tabela.")
        return None

    valores = tabela.item(selecionado[0], "values")
    return valores


def registrar_movimento(tipo):
    produto_selecionado = obter_produto_selecionado()

    if produto_selecionado is None:
        return

    codigo = produto_selecionado[0]
    produto = produto_selecionado[1]

    janela = tk.Toplevel(root)
    janela.title(f"Registrar {tipo}")
    janela.geometry("400x330")
    janela.configure(bg=COR_FUNDO)

    tk.Label(janela, text=f"Produto: {produto}", bg=COR_FUNDO, fg=COR_TEXTO, font=("Arial", 11, "bold")).pack(pady=8)
    tk.Label(janela, text=f"Código: {codigo}", bg=COR_FUNDO, fg=COR_TEXTO).pack(pady=3)

    tk.Label(janela, text="Quantidade:", bg=COR_FUNDO, fg=COR_TEXTO).pack()
    entrada_quantidade = tk.Entry(janela, width=30)
    entrada_quantidade.pack(pady=5)

    tk.Label(janela, text="Responsável:", bg=COR_FUNDO, fg=COR_TEXTO).pack()
    entrada_responsavel = tk.Entry(janela, width=30)
    entrada_responsavel.insert(0, "Rafael")
    entrada_responsavel.pack(pady=5)

    tk.Label(janela, text="Preço unitário, apenas para compra:", bg=COR_FUNDO, fg=COR_TEXTO).pack()
    entrada_preco = tk.Entry(janela, width=30)
    entrada_preco.pack(pady=5)

    def salvar_movimento():
        try:
            quantidade = float(entrada_quantidade.get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Erro", "A quantidade deve ser um número.")
            return

        if quantidade <= 0:
            messagebox.showerror("Erro", "A quantidade deve ser maior que zero.")
            return

        responsavel = entrada_responsavel.get().strip()
        if responsavel == "":
            responsavel = "Não informado"

        preco_texto = entrada_preco.get().replace(",", ".").strip()

        if tipo == "Compra" and preco_texto != "":
            try:
                preco_unitario = float(preco_texto)
            except ValueError:
                messagebox.showerror("Erro", "O preço deve ser um número.")
                return
        else:
            preco_unitario = ""

        total_compra = ""
        if tipo == "Compra" and preco_unitario != "":
            total_compra = quantidade * preco_unitario

        wb = abrir_excel()
        if wb is None:
            return

        ws = wb[ABA_MOVIMENTOS]
        nova_linha = ws.max_row + 1
        data_hoje = datetime.now()
        chave_sistema = f"{codigo}|{tipo}"

        ws.cell(row=nova_linha, column=1).value = data_hoje
        ws.cell(row=nova_linha, column=2).value = codigo
        ws.cell(row=nova_linha, column=3).value = produto
        ws.cell(row=nova_linha, column=4).value = tipo
        ws.cell(row=nova_linha, column=5).value = quantidade
        ws.cell(row=nova_linha, column=6).value = responsavel
        ws.cell(row=nova_linha, column=7).value = preco_unitario
        ws.cell(row=nova_linha, column=8).value = total_compra
        ws.cell(row=nova_linha, column=9).value = "Movimento registrado pelo sistema"
        ws.cell(row=nova_linha, column=10).value = chave_sistema

        recalcular_produto(wb, codigo)

        try:
            wb.save(ARQUIVO_EXCEL)
        except PermissionError:
            messagebox.showerror("Erro", "Fecha o Excel antes de guardar alterações.")
            wb.close()
            return

        wb.close()

        messagebox.showinfo("Sucesso", f"{tipo} registrada com sucesso.")
        janela.destroy()
        atualizar_tabela()

    criar_botao(janela, "Salvar", salvar_movimento, largura=20).pack(pady=18)


def adicionar_produto():
    janela = tk.Toplevel(root)
    janela.title("Adicionar produto")
    janela.geometry("430x470")
    janela.configure(bg=COR_FUNDO)

    tk.Label(
        janela,
        text="Adicionar novo produto",
        bg=COR_TOPO,
        fg=COR_BRANCO,
        font=("Arial", 13, "bold"),
        pady=8
    ).pack(fill=tk.X)

    campos = {}

    labels = [
        "Código",
        "Produto",
        "Categoria",
        "Unidade",
        "Estoque Inicial",
        "Estoque Mínimo",
        "Observações"
    ]

    for label in labels:
        tk.Label(janela, text=label, bg=COR_FUNDO, fg=COR_TEXTO).pack(pady=(8, 0))
        entrada = tk.Entry(janela, width=42)
        entrada.pack(pady=3)
        campos[label] = entrada

    def salvar_produto():
        codigo = campos["Código"].get().strip()
        produto = campos["Produto"].get().strip()
        categoria = campos["Categoria"].get().strip()
        unidade = campos["Unidade"].get().strip()
        observacoes = campos["Observações"].get().strip()

        try:
            estoque_inicial = float(campos["Estoque Inicial"].get().replace(",", "."))
            estoque_minimo = float(campos["Estoque Mínimo"].get().replace(",", "."))
        except ValueError:
            messagebox.showerror("Erro", "Estoque inicial e estoque mínimo devem ser números.")
            return

        if codigo == "" or produto == "":
            messagebox.showerror("Erro", "Código e produto são obrigatórios.")
            return

        wb = abrir_excel()
        if wb is None:
            return

        ws = wb[ABA_ESTOQUE]

        if encontrar_linha_produto(ws, codigo) is not None:
            messagebox.showerror("Erro", "Já existe um produto com esse código.")
            wb.close()
            return

        nova_linha = ws.max_row + 1
        status = "COMPRAR" if estoque_inicial <= estoque_minimo else "OK"

        ws.cell(row=nova_linha, column=1).value = codigo
        ws.cell(row=nova_linha, column=2).value = produto
        ws.cell(row=nova_linha, column=3).value = categoria
        ws.cell(row=nova_linha, column=4).value = unidade
        ws.cell(row=nova_linha, column=5).value = estoque_inicial
        ws.cell(row=nova_linha, column=6).value = 0
        ws.cell(row=nova_linha, column=7).value = 0
        ws.cell(row=nova_linha, column=8).value = estoque_inicial
        ws.cell(row=nova_linha, column=9).value = estoque_minimo
        ws.cell(row=nova_linha, column=10).value = status
        ws.cell(row=nova_linha, column=11).value = ""
        ws.cell(row=nova_linha, column=12).value = observacoes

        try:
            wb.save(ARQUIVO_EXCEL)
        except PermissionError:
            messagebox.showerror("Erro", "Fecha o Excel antes de guardar alterações.")
            wb.close()
            return

        wb.close()

        messagebox.showinfo("Sucesso", "Produto adicionado com sucesso.")
        janela.destroy()
        atualizar_tabela()

    criar_botao(janela, "Salvar produto", salvar_produto, largura=22).pack(pady=15)


def criar_botao(local, texto, comando, largura=18):
    return tk.Button(
        local,
        text=texto,
        command=comando,
        width=largura,
        bg=COR_BOTAO,
        fg=COR_BRANCO,
        activebackground=COR_BOTAO_HOVER,
        activeforeground=COR_BRANCO,
        font=("Arial", 10, "bold"),
        relief=tk.FLAT,
        cursor="hand2",
        padx=5,
        pady=5
    )


# =========================
# INTERFACE GRÁFICA
# =========================
root = tk.Tk()
root.title("Sistema de Gestão de Stock de Produtos de Limpeza")
root.geometry("1050x600")
root.configure(bg=COR_FUNDO)

style = ttk.Style()
style.theme_use("clam")

style.configure(
    "Treeview",
    background=COR_BRANCO,
    foreground=COR_TEXTO,
    rowheight=30,
    fieldbackground=COR_BRANCO,
    font=("Arial", 10)
)

style.configure(
    "Treeview.Heading",
    background=COR_TOPO,
    foreground=COR_BRANCO,
    font=("Arial", 10, "bold")
)

style.map(
    "Treeview",
    background=[("selected", "#B2DFDB")],
    foreground=[("selected", COR_TEXTO)]
)

titulo = tk.Label(
    root,
    text="Sistema de Gestão de Stock de Produtos de Limpeza",
    font=("Arial", 18, "bold"),
    bg=COR_TOPO,
    fg=COR_BRANCO,
    pady=14
)
titulo.pack(fill=tk.X)

subtitulo = tk.Label(
    root,
    text="Linhas vermelhas = COMPRAR | Linhas verdes = OK",
    font=("Arial", 10, "bold"),
    bg=COR_FUNDO,
    fg=COR_TEXTO
)
subtitulo.pack(pady=8)

frame_pesquisa = tk.Frame(root, bg=COR_FUNDO)
frame_pesquisa.pack(pady=5)

tk.Label(
    frame_pesquisa,
    text="Pesquisar produto:",
    bg=COR_FUNDO,
    fg=COR_TEXTO,
    font=("Arial", 10, "bold")
).pack(side=tk.LEFT, padx=5)

entrada_pesquisa = tk.Entry(frame_pesquisa, width=42)
entrada_pesquisa.pack(side=tk.LEFT, padx=5)

criar_botao(frame_pesquisa, "Pesquisar", pesquisar_produto).pack(side=tk.LEFT, padx=5)
criar_botao(frame_pesquisa, "Mostrar todos", mostrar_todos).pack(side=tk.LEFT, padx=5)
criar_botao(frame_pesquisa, "Produtos para comprar", mostrar_produtos_para_comprar, largura=22).pack(side=tk.LEFT, padx=5)

frame_tabela = tk.Frame(root, bg=COR_FUNDO)
frame_tabela.pack(pady=10, fill=tk.BOTH, expand=True, padx=15)

colunas = (
    "Código",
    "Produto",
    "Categoria",
    "Unidade",
    "Estoque Atual",
    "Estoque Mínimo",
    "Comprar?"
)

tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=15)

for coluna in colunas:
    tabela.heading(coluna, text=coluna)

# Largura das colunas
tabela.column("Código", width=100, anchor=tk.CENTER)
tabela.column("Produto", width=250)
tabela.column("Categoria", width=160)
tabela.column("Unidade", width=100, anchor=tk.CENTER)
tabela.column("Estoque Atual", width=120, anchor=tk.CENTER)
tabela.column("Estoque Mínimo", width=120, anchor=tk.CENTER)
tabela.column("Comprar?", width=120, anchor=tk.CENTER)

# Cores das linhas da tabela
tabela.tag_configure("alerta", background=COR_ALERTA_FUNDO, foreground=COR_ALERTA_TEXTO)
tabela.tag_configure("ok", background=COR_OK_FUNDO, foreground=COR_OK_TEXTO)

barra_rolagem = ttk.Scrollbar(frame_tabela, orient=tk.VERTICAL, command=tabela.yview)
tabela.configure(yscrollcommand=barra_rolagem.set)

tabela.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
barra_rolagem.pack(side=tk.RIGHT, fill=tk.Y)

frame_botoes = tk.Frame(root, bg=COR_FUNDO)
frame_botoes.pack(pady=12)

criar_botao(frame_botoes, "Adicionar produto", adicionar_produto, largura=20).pack(side=tk.LEFT, padx=6)
criar_botao(frame_botoes, "Registrar compra", lambda: registrar_movimento("Compra"), largura=20).pack(side=tk.LEFT, padx=6)
criar_botao(frame_botoes, "Registrar retirada", lambda: registrar_movimento("Retirada"), largura=20).pack(side=tk.LEFT, padx=6)
criar_botao(frame_botoes, "Atualizar tabela", atualizar_tabela, largura=20).pack(side=tk.LEFT, padx=6)

atualizar_tabela()

root.mainloop()
ARQUIVO_EXCEL = "estoque_limpeza.xlsx"
print("Pasta atual:", os.getcwd())
print("Ficheiros encontrados:", os.listdir())
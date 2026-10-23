from datetime import datetime, timedelta
import os
import threading
import time
import uuid
import customtkinter as ctk
from tkinter import messagebox
from plyer import notification
import pandas as pd

# Configuração inicial do CustomTkinter (Define tema e cores padrão baseados no sistema)
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

import os

PASTA_DADOS = os.path.join(
    os.getenv("APPDATA"),
    "GerenciadorTarefas"
)

os.makedirs(PASTA_DADOS, exist_ok=True)

EXCEL_FILE = os.path.join(
    PASTA_DADOS,
    "tarefas.xlsx"
)

COLUNAS_TAREFAS = ["ID", "Tarefa", "Data", "Hora", "Cor", "Notificado"]
COLUNAS_HISTORICO = ["ID", "Tarefa", "Data", "Hora", "Ação", "DataRegistro"]

class GerenciadorApp(ctk.CTk):

  def __init__(self):
    """Método construtor: Inicializa a janela principal do aplicativo,

    configura o layout (títulos, entradas, listas, botões), carrega as
    tarefas salvas no Excel e inicia o monitor de notificações em segundo plano.
    """
    super().__init__()

    self.title("Gerenciador de Tarefas")
    self.geometry("520x640")
    self.resizable(False, False)

    # --- Layout / Componentes ---
    self.label_titulo = ctk.CTkLabel(
        self, text="Gerenciador de Tarefas", font=("Arial", 20, "bold")
    )
    self.label_titulo.pack(pady=15)

    # Frame de Entrada (Campos de texto para digitar a tarefa e o horário)
    self.frame_entrada = ctk.CTkFrame(self, fg_color="transparent")
    self.frame_entrada.pack(pady=5, padx=20, fill="x")

    self.entrada_tarefa = ctk.CTkEntry(
        self.frame_entrada,
        placeholder_text="Digite a tarefa...",
        font=("Arial", 12),
        height=35,
    )
    self.entrada_tarefa.pack(side="left", expand=True, fill="x", padx=(0, 10))

    #campo data
    self.entrada_data = ctk.CTkEntry(
        self.frame_entrada,
        placeholder_text="DD/MM/AAAA",
        font=("Arial", 12),
        width=100,
        height=35,
    )
    self.entrada_data.pack(side="left", padx=(0, 10))

    #campo hora
    self.entrada_hora = ctk.CTkEntry(
        self.frame_entrada,
        placeholder_text="HH:MM",
        font=("Arial", 12),
        width=80,
        height=35,
    )
    self.entrada_hora.pack(side="right")

    # Botão Adicionar (Aciona a função de cadastrar nova tarefa)
    self.btn_adicionar = ctk.CTkButton(
        self,
        text="Adicionar Tarefa",
        command=self.adicionar_tarefa,
        fg_color="#2ecc71",
        hover_color="#27ae60",
        font=("Arial", 12, "bold"),
        height=40,
    )
    self.btn_adicionar.pack(pady=10, padx=20, fill="x")

    # Lista de Tarefas (Área com rolagem onde os itens cadastrados aparecem)
    self.scroll_tarefas = ctk.CTkScrollableFrame(
        self, label_text="Tarefas e Status", label_font=("Arial", 12, "bold")
    )
    self.scroll_tarefas.pack(pady=10, padx=20, fill="both", expand=True)

    # Botão de Excluir Inferior (Remove os itens marcados com checkbox)
    self.btn_remover = ctk.CTkButton(
        self,
        text="Excluir Tarefa Selecionada",
        command=self.remover_tarefa_selecionada,
        fg_color="#e74c3c",
        hover_color="#c0392b",
        font=("Arial", 11, "bold"),
        height=35,
    )
    self.btn_remover.pack(pady=10, padx=20, fill="x")

    #botão concluir 
    #self.btn_concluir = ctk.CTkButton(
    #self, text="Concluir Selecionada", command=self.concluir_tarefas,
    #fg_color="#3498db", hover_color="#2980b9", font=("Arial", 11, "bold"), height=35
    #)
    #self.btn_concluir.pack(pady=5, padx=20, fill="x") não necesário mais

    #botão de historico
    self.btn_historico = ctk.CTkButton(
        self, text="Visualizar Histórico", command=self.visualizar_historico,
        fg_color="#9b59b6", hover_color="#8e44ad", font=("Arial", 11, "bold"), height=35
    )
    self.btn_historico.pack(pady=5, padx=20, fill="x")
    
    # Dicionário estruturado para gerenciar o estado de cada tarefa em tempo de execução
    self.tarefas_dict = {}

    # Inicialização dos dados e da thread de verificação
    self.carregar_tarefas()
    self.iniciar_verificador_notificacoes()

  def carregar_tarefas(self):
    if os.path.exists(EXCEL_FILE):
        try:
            xls = pd.ExcelFile(EXCEL_FILE)
            df_tarefas = pd.read_excel(xls, "Tarefas")
            self.historico = pd.read_excel(xls, "Historico")
            for _, row in df_tarefas.iterrows():
                self.criar_item_visual(
                    row["ID"], row["Tarefa"], row["Data"], row["Hora"],
                    cor_inicial=row.get("Cor", "Normal"),
                    notificado=bool(row.get("Notificado", False))
                )
        except Exception as e:
            print(f"Erro ao carregar: {e}")

  #salva e carrega duas abas: Tarefas e Historico.
  def salvar_no_excel(self):
    dados = []
    for chave, info in self.tarefas_dict.items():
      dados.append({
          "ID": chave,
          "Tarefa": info["tarefa"],
          "Data": info["data"],
          "Hora": info["hora"],
          "Cor": info["cor_var"].get(),
          "Notificado": info["notificado"],
      })
    df_tarefas = pd.DataFrame(dados, columns=COLUNAS_TAREFAS)
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        df_tarefas.to_excel(writer, sheet_name="Tarefas", index=False)
        self.historico.to_excel(writer, sheet_name="Historico", index=False)
    
  def registrar_historico(self, tarefa, acao):
    registro = {
      "ID": tarefa["ID"],
      "Tarefa": tarefa["tarefa"],
      "Data": tarefa["data"],
      "Hora": tarefa["hora"],
      "Ação": acao,
      "DataRegistro": datetime.now().strftime("%d/%m/%Y %H:%M")
    }
    self.historico = pd.concat([self.historico, pd.DataFrame([registro])], ignore_index=True)

  def visualizar_historico(self):
    hist_win = ctk.CTkToplevel(self)
    hist_win.title("Histórico de Tarefas")
    hist_win.geometry("500x400")

    frame = ctk.CTkScrollableFrame(hist_win, label_text="Histórico", label_font=("Arial", 12, "bold"))
    frame.pack(fill="both", expand=True, padx=10, pady=10)

    for _, row in self.historico.iterrows():
      ctk.CTkLabel(frame, text=f"{row['Ação']} - {row['Tarefa']} ({row['Data']} {row['Hora']}) em {row['DataRegistro']}", font=("Arial", 11)).pack(anchor="w", pady=2)

  def criar_item_visual(self, id_tarefa, tarefa, data, horario, cor_inicial="Normal", notificado=False):
    """Monta graficamente uma linha de tarefa na tela contendo:

    um checkbox de seleção, o texto com o nome/horário e o menu dropdown
    de escolha de cor (Normal, Verde ou Vermelho).
    """
    if id_tarefa in self.tarefas_dict:
      return

    item_frame = ctk.CTkFrame(self.scroll_tarefas, fg_color=("gray85", "gray20"))
    item_frame.pack(fill="x", pady=4, padx=5)

    var = ctk.BooleanVar(value=False)
    checkbox = ctk.CTkCheckBox(
      item_frame,
      text=f"{tarefa} — Prazo: {data} {horario}",
      variable=var,
      font=("Arial", 12),
    )
    checkbox.pack(side="left", padx=10, pady=8)

    # Menu Dropdown para o usuário escolher manualmente a cor/status da tarefa
    cor_var = ctk.StringVar(value=cor_inicial)
    menu_cor = ctk.CTkOptionMenu(
        item_frame,
        values=["Normal", "Verde", "Vermelho"],
        variable=cor_var,
        command=lambda _: self.atualizar_estilo_item(id_tarefa),
        width=110,
        font=("Arial", 11),
    )
    menu_cor.pack(side="right", padx=10)

    # Armazena os componentes e informações no dicionário principal
    self.tarefas_dict[id_tarefa] = {
      "ID": id_tarefa,
      "tarefa": tarefa,
      "data": data,
      "hora": horario,
      "cor_var": cor_var,
      "notificado": notificado,
      "frame": item_frame,
      "checkbox": checkbox,
      "var": var,
    }

    self.atualizar_estilo_item(id_tarefa)

  def atualizar_estilo_item(self, chave):
    """Altera dinamicamente a cor de fundo e do texto do item na tela

    com base na opção escolhida pelo usuário no menu dropdown e salva no Excel.
    """
    info = self.tarefas_dict[chave]
    escolha = info["cor_var"].get()
    frame = info["frame"]
    checkbox = info["checkbox"]

    if escolha == "Verde":
      frame.configure(fg_color=("#d4edda", "#155724"))
      checkbox.configure(text_color=("#155724", "#d4edda"))
    elif escolha == "Vermelho":
      frame.configure(fg_color=("#f8d7da", "#721c24"))
      checkbox.configure(text_color=("#721c24", "#f8d7da"))
    else:  # Normal (Aplica a cor padrão do app)
      frame.configure(fg_color=("gray85", "gray20"))
      checkbox.configure(text_color=("black", "white"))

    self.salvar_no_excel()

  def adicionar_tarefa(self):
    """Valida os campos de entrada digitados pelo usuário,

    cria o item visual na tela e aciona o salvamento no Excel.
    """
    tarefa = self.entrada_tarefa.get().strip()
    data = self.entrada_data.get().strip()
    horario = self.entrada_hora.get().strip()
    
    if not tarefa or not data or not horario:
      messagebox.showwarning("Aviso", "Preencha nome, data e hora!")
      return

    try:
      datetime.strptime(f"{data} {horario}", "%d/%m/%Y %H:%M")
    except ValueError:
      messagebox.showerror("Erro", "Formato inválido! Use DD/MM/AAAA e HH:MM")
      return

    id_tarefa = str(uuid.uuid4())
    self.criar_item_visual(id_tarefa, tarefa, data, horario)
    self.salvar_no_excel()

    self.entrada_tarefa.delete(0, "end")
    self.entrada_data.delete(0,"end")
    self.entrada_hora.delete(0, "end")
    

  def remover_tarefa_selecionada(self):
    """Remove da tela e da planilha do Excel todas as tarefas

    que estiverem com a caixinha de seleção (checkbox) marcada.
    """
    removidas = []
    for chave, info in list(self.tarefas_dict.items()):
      if info["var"].get() == 1:
        info["frame"].destroy()
        del self.tarefas_dict[chave]
        removidas.append(chave)

    if removidas:
      self.salvar_no_excel()
    else:
      messagebox.showwarning(
          "Aviso", "Marque a caixa ao lado da tarefa que deseja excluir."
      )

  def disparar_notificacao(self, tarefa, data, horario):
    """Dispara uma notificação nativa do Windows no notebook

    avisando sobre o prazo da tarefa.
    """
    try:
      notification.notify(
        title="Lembrete de Prazo (15 min)",
        message=f"A tarefa '{tarefa}' vence em {data} {horario}!",
        app_name="Gerenciador de Tarefas",
        timeout=10,
      )
    except Exception as e:
      print(f"Erro na notificação: {e}")

  def iniciar_verificador_notificacoes(self):
    """Inicia uma tarefa secundária (Thread) em segundo plano que roda

    continuamente checando se faltam 15 minutos para o horário de alguma tarefa,
    disparando o aviso caso necessário.
    """

    def verificar():
      while True:
        agora = datetime.now()
        for chave, info in list(self.tarefas_dict.items()):
          if not info["notificado"]:
            try:
              prazo = datetime.strptime(f"{info["data"]} {info["hora"]}", "%d/%m/%Y %H:%M")
              tempo_aviso = prazo - timedelta(minutes=15)
              
              if tempo_aviso <= agora <= prazo:
                self.disparar_notificacao(info["tarefa"], f"{info["data"]} {info["hora"]}")
                info["notificado"] = True
                self.salvar_no_excel()
            except:  
              pass
        time.sleep(15)

    t = threading.Thread(target=verificar, daemon=True)
    t.start()


if __name__ == "__main__":
  # Ponto de entrada: instancia a classe principal e inicia o loop da interface gráfica
  app = GerenciadorApp()
  app.mainloop()
# Projeto-de-um-app-de-agenda
Primeiro projeto de app, pra conhecimento com objetivo de ter evoluções durante o tempo.

# 📋 Gerenciador de Tarefas com Notificações

Aplicativo desktop desenvolvido em Python utilizando **CustomTkinter**, com armazenamento local em Excel e sistema de notificações automáticas para gerenciamento de tarefas e prazos.

---

## 🚀 Funcionalidades

### ✅ Cadastro de Tarefas
- Adicionar tarefas com:
  - Nome da tarefa
  - Data de vencimento
  - Horário de vencimento

### ✅ Identificação Única
- Cada tarefa recebe um identificador único (UUID).
- Permite cadastrar múltiplas tarefas com o mesmo nome, data e horário sem conflitos.

### ✅ Persistência de Dados
- Salvamento automático em arquivo Excel.
- Recuperação automática das tarefas ao abrir o aplicativo.

### ✅ Histórico de Atividades
- Registro permanente de:
  - Tarefas concluídas
  - Tarefas excluídas
- Histórico armazenado em aba separada da planilha.

### ✅ Notificações Automáticas
- Alertas emitidos 15 minutos antes do prazo.
- Utiliza notificações nativas do Windows através da biblioteca Plyer.

### ✅ Controle Visual de Status
Cada tarefa pode receber um status visual:

| Status | Cor |
|----------|----------|
| Normal | Padrão |
| Verde | Concluída ou Em dia |
| Vermelho | Atenção ou Prioridade |

### ✅ Interface Moderna
- Desenvolvida com CustomTkinter.
- Suporte automático aos modos claro e escuro.
- Layout simples e intuitivo.

---

## 🛠 Tecnologias Utilizadas

- Python 3.10+
- CustomTkinter
- Pandas
- OpenPyXL
- Plyer
- Threading
- UUID

---

## 📂 Estrutura dos Dados

O sistema cria automaticamente uma pasta local:

```text
%APPDATA%\GerenciadorTarefas
```

E dentro dela:

```text
tarefas.xlsx
```

A planilha contém duas abas:

### Tarefas

| ID | Tarefa | Data | Hora | Cor | Notificado |
|------|------|------|------|------|------|

### Histórico

| ID | Tarefa | Data | Hora | Ação | DataRegistro |
|------|------|------|------|------|------|


## 🔔 Funcionamento das Notificações

O sistema executa uma verificação contínua em segundo plano.

Quando faltar 15 minutos para o prazo definido, uma notificação será exibida:

```text
Lembrete de Prazo (15 min)

A tarefa "Exemplo" vence em 20/10/2026 14:00.
```

Cada tarefa é notificada apenas uma vez.


## 👨‍💻 Autor

Desenvolvido por Amanda utilizando Python e CustomTkinter.
# 🧾 Gerador de Termos de Responsabilidade - Luma Service

Sistema desktop em Python com interface moderna para geração de termos de responsabilidade personalizados, utilizado pelo Departamento Pessoal e áreas relacionadas.

---

## ✨ Funcionalidades

- ✅ Interface gráfica com `ttkbootstrap` (tema flatly)
- ✅ Geração de documentos Word com campos substituíveis
- ✅ Envio automático via Outlook com anexo
- ✅ Registro de logs com nome do usuário, horário e informações do termo
- ✅ Formatação automática de:
  - CPF → 000.000.000-00
  - Data → 00/00/0000
  - Telefone → (00) 00000-0000
  - Cartão → 0000 0000 0000 0000
- ✅ Detecção dinâmica dos modelos `.docx` disponíveis
- ✅ Suporte a diferentes tipos de equipamentos (Celular, Notebook, Cartão)

---

## 🛠️ Tecnologias utilizadas

- Python 3.11
- [`ttkbootstrap`](https://ttkbootstrap.readthedocs.io/)
- `python-docx`
- `pywin32` (integração com Outlook)
- `os`, `datetime`, `subprocess`, `winreg`

---

## 📁 Estrutura do Projeto

gerador-termo-luma/ 
├── gerador_email.py # Código principal com GUI ├
── modelos/ # Modelos de termos (.docx) ├
── logs/ # Logs de geração ├
── requirements.txt └
── README.md

---

## 🚀 Como executar

1. Clone este repositório:

```bash
git clone https://github.com/seu-usuario/gerador-termo-luma.git
Instale as dependências:

bash
Copiar
Editar
pip install -r requirements.txt
Execute o programa:

bash
Copiar
Editar
python gerador_email.py
📝 Exemplo de log
yaml
Copiar
Editar
[11/04/2025 19:25] Quem gerou: guilherme.marques | Modelo: Termo_Celular.docx | Nome: Guilherme Marques | CPF: 123.456.789-00 | Arquivo: C:\Users\Desktop\Termo_Guilherme.docx
📌 Observações
Os modelos devem conter os placeholders:

[ nome do colaborador ]

[ CPF do colaborador ]

[ data do dia]

[ número do cartão ], [ modelo do notebook ], etc.

Todos os documentos são gerados na Área de Trabalho do usuário.

🤝 Desenvolvido por
Guilherme Marques
Analista de Desenvolvimento de Sistemas


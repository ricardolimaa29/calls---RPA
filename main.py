## Utilizar o Flet com um design simples com um Botão e uma descrição de status
## deixar tudo em exe para ser usado em qualquer maquina
## postar no Github com readme perfeito


import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait

import flet as ft
hoje = datetime.now().strftime("%d/%m")
lista_de_professores = ["Ricardo 🦆","Matheus 👑","Wilck 😎","Johnny 🦹‍♂️","Fernando 😘","Lazaro 🤡"]
lista_de_turmas = ["SALA 04 14H","SALA 04 16H","SALA 04 SEG E QUA","SALA 04 TER E QUI"]
def main(page:ft.Page):
    def pegar_dados_sheet(e):
        turma = turma_dropdown.value
        if not turma:
            mensagens.value = "❗ Por favor, selecione uma turma."
            page.update()
            return
        # Autenticação
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        creds = Credentials.from_service_account_file("credenciais.json", scopes=scopes)
        client = gspread.authorize(creds)

        # Abre a planilha e a aba
        planilha = client.open("PRESENÇA FÁBRICA DE PROGRAMADORES")
        aba = planilha.worksheet(turma)

        dados = aba.get_all_values()
        linha_datas = dados[3]
        if hoje in linha_datas:
            col_index = linha_datas.index(hoje)
            alunos_presentes = [
                linha[1] for linha in dados[3:]
                if col_index < len(linha) and linha[col_index].strip().upper() == "C"
            ]
            if alunos_presentes:
                alunos_presentes_text.value = "\n".join(alunos_presentes)
                mensagens.value = "✅ Chamada localizada! Alunos com presença:"
            else:
                alunos_presentes_text.value = ""
                mensagens.value = "🤷‍♂️ Nenhum aluno com presença 'C' hoje."
        else:
            alunos_presentes_text.value = ""
            mensagens.value = "😢 Data de hoje não encontrada na planilha."
            print(dados[3])
        print("Dados: \n",dados[3])
        page.update()

    page.title = "Calls --- RPA"
    page.theme_mode = "dark"
    page.window.width = 600
    page.window.height = 400
    page.window.max_width = 600
    page.window.max_height = 400
    page.window.min_width = 600
    page.window.min_height = 400
    mensagens = ft.Text(f" ☀, deseja atualizar a chamada de hoje ?  📅  {hoje}", color="White",size=20)
    turma_dropdown = ft.Dropdown(
        label="Selecione a turma",
        width=300,
        options=[ft.dropdown.Option(text=t, key=t) for t in lista_de_turmas]
    )
    professor_dropdown = ft.Dropdown(
        label="Selecione o professor",
        width=300,
        options=[ft.dropdown.Option(text=t, key=t) for t in lista_de_professores]
    )
    comecar = ft.ElevatedButton("Iniciar",on_click=pegar_dados_sheet,width=200,color="Blue",bgcolor="White")
    alunos_presentes_text = ft.Text("",color="Yellow",size=15)

    page.add(ft.Row([mensagens],alignment="center"),
             ft.Row([alunos_presentes_text],alignment="center"),
             ft.Row([professor_dropdown,turma_dropdown],alignment="center"),
             ft.Row([comecar],alignment="center")
             )


ft.app(target=main)


# driver = webdriver.Chrome()


# # Abrir o sistema SENAI
# driver.get("https://diariofic.sp.senai.br/")
# time.sleep(2)
# # Preencher um campo (exemplo)
# campo_usuario = driver.find_element(By.NAME, "aIdentificacao")
# campo_usuario.send_keys("sn1099962")

# campo_senha = driver.find_element(By.NAME, "aSenha")
# campo_senha.send_keys("Rede29@05@@")
# campo_senha.send_keys(Keys.ENTER)
# time.sleep(3)
# driver.get("https://diariofic.sp.senai.br/Turma")


# time.sleep(30)









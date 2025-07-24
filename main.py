import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import flet as ft
import time

hoje = datetime.now().strftime("%d/%m")
hora = datetime.now().time()
hora_agora = hora.hour


def main(page: ft.Page):
    if hora_agora < 12:
        saudacao = "Bom dia 🌄"
    elif hora_agora < 18:
        saudacao = "Boa tarde 🌞"
    else:
        saudacao = "Boa noite 🌃"

    page.title = "Calls --- RPA"
    page.theme_mode = "dark"
    page.window.maximized = True
    page.fonts = {
        "Poppins": "fonts/Poppins-Bold.ttf",
        "Poppins2": "fonts/Poppins-Light.ttf",
        "Poppins3": "fonts/Poppins-Regular.ttf",
    }

    mensagens = ft.Text(f"{saudacao}", color="White", size=30,font_family="Poppins")
    alunos_presentes_text = ft.Text("", color="Yellow", size=15,font_family="Poppins2")

    # Autenticação Google Sheets
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file("credenciais.json", scopes=scopes)
    client = gspread.authorize(creds)

    # Buscar planilhas compartilhadas
    planilhas = client.openall()
    dropdown_planilhas = ft.Dropdown(
        label="Selecione a planilha",
        width=300,
        options=[ft.dropdown.Option(p.title) for p in planilhas],
    )
    dropdown_planilhas_post = ft.Dropdown(
        label="Selecione a chamada",
        width=300,
        options=[ft.dropdown.Option(p.title) for p in planilhas],
        visible=False
    )

    dropdown_abas = ft.Dropdown(
        label="Selecione a aba",
        width=300,
        options=[],
    )

    # Atualiza abas ao selecionar planilha
    def ao_selecionar_planilha(e):
        nome_planilha = dropdown_planilhas.value
        if not nome_planilha:
            mensagens.value = "❗ Selecione uma planilha válida."
            page.update()
            return

        planilha = client.open(nome_planilha)
        abas = planilha.worksheets()
        dropdown_abas.options = [ft.dropdown.Option(aba.title) for aba in abas]
        dropdown_abas.value = None
        mensagens.value = "✅ Planilha carregada, selecione a aba."
        page.update()

    dropdown_planilhas.on_change = ao_selecionar_planilha

    def pegar_dados_sheet(e):
        dropdown_planilhas_post.visible = True
        planilha_nome = dropdown_planilhas.value
        aba_nome = dropdown_abas.value

        if not planilha_nome or not aba_nome:
            mensagens.value = "❗ Selecione a planilha e a aba corretamente."
            alunos_presentes_text.value = ""
            page.update()
            return

        planilha = client.open(planilha_nome)
        aba = planilha.worksheet(aba_nome)

        dados = aba.get_all_values()
        if len(dados) <= 3:
            mensagens.value = "❗ Dados insuficientes na aba."
            alunos_presentes_text.value = ""
            page.update()
            return

        linha_datas = dados[3]
        alunos = dados[4:]  # Pula os 4 primeiros cabeçalhos

        mensagens.value = f"✅ Exibindo dados completos da chamada"
        alunos_presentes_text.value = ""

        colunas = [ft.DataColumn(ft.Text("Aluno", weight="bold"))]
        for data in linha_datas[2:]:
            colunas.append(ft.DataColumn(ft.Text(data)))

        linhas = []
        for aluno in alunos:
            if len(aluno) < 2:
                continue  
            nome = aluno[1]
            linha = [ft.DataCell(ft.Text(nome, weight="bold",font_family="Poppins2"))]

            for i in range(2, len(linha_datas)):
                valor = aluno[i] if i < len(aluno) else ""
                valor = valor.strip().upper()
                cor = "#FF4C4C" if valor == "F" else None
                cor_verde = "#329703" if valor == "C" else None
                linha.append(ft.DataCell(ft.Container(
                    content=ft.Text(valor or "-"),
                    bgcolor=cor if valor == "F" else cor_verde,
                    padding=5,
                    alignment=ft.alignment.center
                )))
            linhas.append(ft.DataRow(linha))

        tabela_chamada = ft.DataTable(
            columns=colunas,
            rows=linhas,
            heading_row_color="#121212",
            border=ft.border.all(1, ft.Colors.GREY_800),
            column_spacing=15,
            data_row_color={"hovered": "#121212"},
            show_checkbox_column=False,
            divider_thickness=0.5
        )

        tabela_scroll_container = ft.Container(
            content=ft.Column(
                [
                    ft.Row(
                        controls=[tabela_chamada],
                        scroll="auto"
                    )
                ],
                scroll="auto", 
                expand=False
            ),
            padding=10,
            bgcolor="#121212",
            border_radius=10,
            # width=1200,
            height=750,
            alignment=ft.alignment.top_left
        )


        area_tabela.controls.clear()
        area_tabela.controls.append(tabela_scroll_container)
        page.update()
        



        tabela_chamada = ft.DataTable(
            columns=colunas,
            rows=linhas,
            heading_row_color="#202020",
            border=ft.border.all(1, ft.Colors.GREY_800),
            column_spacing=15,
            data_row_color={"hovered": "#1A1A1A"},
            show_checkbox_column=False,
            divider_thickness=0.5
        )

        

    area_tabela = ft.Column([], scroll="auto")

    page.controls.clear()
    botao_iniciar = ft.ElevatedButton(
            "Iniciar", on_click=pegar_dados_sheet, width=200, color="Blue", bgcolor="White"
    )

    page.add(
        ft.Row([mensagens], alignment="center"),
        ft.Row([alunos_presentes_text], alignment="center"),
        ft.Row([dropdown_planilhas, dropdown_abas,dropdown_planilhas_post], alignment="center"),
        ft.Row([botao_iniciar], alignment="center"),
        ft.Divider(thickness=1),
        ft.ResponsiveRow([area_tabela],alignment="center")
    )
    page.update()

ft.app(target=main)

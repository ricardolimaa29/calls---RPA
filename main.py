import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import flet as ft
import pandas as pd

hoje = datetime.now().strftime("%d/%m")
hora = datetime.now().time()
hora_agora = hora.hour
data = datetime.now()
data_formatada = data.strftime("%d/%m/%Y %H:%M:%S")


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
    dropdown_planilhas_post.options = [
        ft.dropdown.Option(f"{p.title}") for p in planilhas
    ]



    dropdown_abas = ft.Dropdown(
        label="Selecione a aba",
        width=300,
        options=[],
    )

    selected_file_path = ft.Text(value="", selectable=True)
    

    def on_file_picked(e: ft.FilePickerResultEvent):
        dropdown_planilhas_post.visible = True
        if e.files:
            file = e.files[0]
            selected_file_path.value = file.path
            try:
                df = pd.read_excel(file.path, engine="openpyxl")
                mensagens.value = f"📄 Arquivo carregado com sucesso: {file.name}\nLinhas: {len(df)}"
                mensagens.color = "Green"
                page.update()

                # Salva no client_storage para uso posterior
                page.client_storage.set("dados_excel", df.to_dict(orient="records"))

                # Mostra na tela a tabela da planilha local
                mostrar_planilha_local_em_tabela(df.to_dict(orient="records"))

            except Exception as err:
                mensagens.value = f"❌ Erro ao ler arquivo: {err}"
                mensagens.color = "Red"
                page.update()


            

    file_picker = ft.FilePicker(on_result=on_file_picked)


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

    file_picker = ft.FilePicker(on_result=on_file_picked)
    page.overlay.append(file_picker)

    btn_selecte_planilhas = ft.ElevatedButton("Selecionar Arquivo", style=ft.ButtonStyle(color="White",bgcolor="Green"), on_click=lambda _: file_picker.pick_files(allow_multiple=False))

    page.overlay.append(file_picker)
    def atualizar_planilha_google(e):
        planilha_nome = dropdown_planilhas.value
        aba_nome = dropdown_abas.value
        dados = page.client_storage.get("dados_excel")

        if not planilha_nome or not aba_nome:
            mensagens.value = "❗ Selecione a planilha e aba para atualizar."
            mensagens.color = "Red"
            page.update()
            return

        if not dados:
            mensagens.value = "❗ Nenhum dado carregado do Excel."
            mensagens.color = "Red"
            page.update()
            return

        try:
            planilha = client.open(planilha_nome)
            aba = planilha.worksheet(aba_nome)

            # Converte os dados de dicionário para lista de listas
            colunas = list(dados[0].keys())
            linhas = [colunas] + [[linha[col] for col in colunas] for linha in dados]

            # Limpa a aba e insere os novos dados
            aba.clear()
            aba.update("A1", linhas)

            mensagens.value = f"✅ Planilha '{aba_nome}' atualizada com sucesso!"
            mensagens.color = "Green"
            page.update()
        except Exception as err:
            mensagens.value = f"❌ Erro ao atualizar planilha: {err}"
            mensagens.color = "Red"
            page.update()
    def mostrar_planilha_local_em_tabela(dados_excel):
        if not dados_excel:
            mensagens.value = "❌ Nenhum dado encontrado na planilha local."
            mensagens.color = "Red"
            page.update()
            return

        colunas_tabela = list(dados_excel[0].keys())
        colunas = [ft.DataColumn(ft.Text(col, weight="bold")) for col in colunas_tabela]

        linhas = []
        for linha in dados_excel:
            linha_data = []
            for valor in linha.values():
                valor_texto = str(valor).strip().upper()
                cor = "#FF4C4C" if valor_texto == "F" else "#329703" if valor_texto == "C" else None
                linha_data.append(
                    ft.DataCell(
                        ft.Container(
                            content=ft.Text(valor_texto),
                            bgcolor=cor,
                            padding=5,
                            alignment=ft.alignment.center,
                        )
                    )
                )
            linhas.append(ft.DataRow(linha_data))

        tabela_local = ft.DataTable(
            columns=colunas,
            rows=linhas,
            heading_row_color="#202020",
            border=ft.border.all(1, ft.Colors.GREY_800),
            column_spacing=15,
            data_row_color={"hovered": "#1A1A1A"},
            show_checkbox_column=False,
            divider_thickness=0.5,
        )

        tabela_scroll_container = ft.Container(
            content=ft.Column(
                [ft.Row(controls=[tabela_local], scroll="auto")],
                scroll="auto",
                expand=False,
            ),
            padding=10,
            bgcolor="#121212",
            border_radius=10,
            height=750,
            alignment=ft.alignment.top_left,
        )

        area_tabela.controls.clear()
        area_tabela.controls.append(tabela_scroll_container)
        page.update()

    def pegar_dados_sheet(e):
        dropdown_planilhas_post.visible = True
        planilha_nome = dropdown_planilhas.value
        aba_nome = dropdown_abas.value
        botao_atualizar_sheet.visible = True

        if not planilha_nome or not aba_nome:
            mensagens.value = "❗ Selecione a planilha e a aba corretamente."
            page.update()
            return

        planilha = client.open(planilha_nome)
        aba = planilha.worksheet(aba_nome)

        dados = aba.get_all_values()
        if len(dados) <= 3:
            mensagens.value = "❗ Dados insuficientes na aba."
            page.update()
            return

        linha_datas = dados[3]
        alunos = dados[4:]  # Pula os 4 primeiros cabeçalhos

        mensagens.value = f"✅ Exibindo dados completos da chamada"

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
            # Agora atualiza a planilha do Google conforme a planilha local
        dados_excel = page.client_storage.get("dados_excel")

        if not dados_excel:
            mensagens.value += "\n❗ Nenhuma planilha local carregada para atualizar os dados."
            mensagens.color = "Red"
            page.update()
            return

        dias_excel = list(dados_excel[0].keys())[2:]  # Assume que datas estão da 3ª coluna em diante

        for i, aluno in enumerate(alunos):
            for j, dia in enumerate(linha_datas[2:], start=2):
                if dia in dias_excel:
                    valor = aluno[j] if j < len(aluno) else ""
                    valor = valor.strip().upper()
                    if valor == "C":
                        aluno[j] = "."
                    elif valor == "F":
                        aluno[j] = "|"

        try:
            novos_dados = dados[:4] + alunos  # cabeçalhos + alunos modificados
            aba.update("A1", novos_dados)
            mensagens.value += "\n✅ Planilha Google atualizada com base na planilha local!"
            mensagens.color = "Green"
            page.update()
        except Exception as err:
            mensagens.value = f"❌ Erro ao atualizar Google Sheets: {err}"
            mensagens.color = "Red"
            page.update()

        

    area_tabela = ft.Column([], scroll="auto")

    page.controls.clear()
    botao_iniciar = ft.ElevatedButton(
            "Iniciar", on_click=pegar_dados_sheet, width=200, color="Blue", bgcolor="White"
    )
    botao_atualizar_sheet = ft.ElevatedButton(
        "Atualizar Planilha", 
        width=200, 
        color="White", 
        bgcolor="Orange",
        on_click=lambda e: atualizar_planilha_google(e),
        visible= False
    )


    page.add(
        ft.Row([mensagens], alignment="center"),
        ft.Row([btn_selecte_planilhas], alignment="center"),
        ft.Row([dropdown_planilhas,dropdown_abas], alignment="center"),
        ft.Row([botao_iniciar,botao_atualizar_sheet], alignment="center"),
        ft.Divider(thickness=1),
        ft.ResponsiveRow([area_tabela],alignment="center"),
    )
    page.update()

ft.app(target=main)

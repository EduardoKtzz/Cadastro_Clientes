from flask import Flask, jsonify
from datetime import datetime
from playwright.async_api import async_playwright
import asyncio
import win32com.client
import pythoncom  # Import necessário para inicializar o COM

app = Flask(__name__)

async def main():
    pythoncom.CoInitialize()  # Inicializar COM para uso com win32com em contexto assíncrono

    async with async_playwright() as p:
        navegador = await p.chromium.launch(headless=False)
        page = await navegador.new_page()
        
        # Acessar a página da VTEX e esperar ela carregar
        await page.goto('https://adcosprofessional.myvtex.com/admin/b2b-organizations/organizations/#/requests')
        await page.wait_for_load_state('networkidle')

        # Login no VTEX
        await page.fill('#email', '*********')
        await page.click('[data-testid="email-form-continue"]')
        await page.fill('input[name="password"]', '**********')
        await page.click('#chooseprovider_signinbtn')
        await page.wait_for_load_state('networkidle')

        # Localizar o iframe e clicar no botão de filtros
        iframe_localizar = page.frame_locator('iframe[title="admin iframe"]')
        await iframe_localizar.locator('button.c-action-primary').nth(0).wait_for(timeout=20000)
        await iframe_localizar.locator('button.c-action-primary:has-text("Status: Todos")').click(timeout=10000)

        # Desmarcar opções de aprovados e recusados
        filtro_caixa = iframe_localizar.locator('.fixed.absolute-ns.w-100.w-auto-ns.z-999.ba.bw1.b--muted-4.bg-base.left-0.br2.o-100')
        await filtro_caixa.wait_for(state='visible', timeout=10000)
        await iframe_localizar.locator('input[name="status-checkbox-group"][value="approved"]').uncheck()
        await iframe_localizar.locator('input[name="status-checkbox-group"][value="declined"]').uncheck()

        # Aplicar filtros
        await iframe_localizar.locator('button:has(div.vtex-button__label:has-text("Apply"))').click(timeout=10000)

        # Rolar até a opção de selecionar 100 emails por página
        dropdown = iframe_localizar.locator('select.o-0.absolute.top-0.left-0.h-100.w-100.bottom-0.t-body.pointer')
        await dropdown.click()
        await dropdown.select_option("100")
        await page.wait_for_timeout(10000)

        # Coletar emails e datas
        todos_dados = await iframe_localizar.locator(r'#render-admin\.app\.b2b-organizations\.organizations > div > div.admin-ui-c-ervJfA > div > div > div.flex.flex-column > div:nth-child(1) > div.vh-100.w-100.dt > div > div:nth-child(1) > div > div:nth-child(2) > div:nth-child(2) > div').all_inner_texts()
        dados_unidos = ''.join(todos_dados)
        dados_separados = dados_unidos.split('\n')

        # Separar emails e datas
        emails = []
        datas = []

        for i in range(len(dados_separados)):
            if i % 3 == 0:
                emails.append(dados_separados[i])
            elif (i - 2) % 3 == 0:
                datas.append(dados_separados[i])

        resultado = list(zip(emails, datas))

        # Abrir Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True

        # Acessar a aba ativa da planilha
        planilha = excel.ActiveWorkbook.ActiveSheet

        # Identificar o email mais recente na linha 16
        ultimo_email = planilha.Range('B16').Value

        # Filtrar novos emails
        novos_resultados = []
        encontrou_ultimo_email = False

        for email, data in resultado:
            if email == ultimo_email:
                encontrou_ultimo_email = True
                break  

            if not encontrou_ultimo_email:
                novos_resultados.append((email, data))

        # Inserir novos dados e empurrar antigos para baixo
        linha_inicial = 16
        Id = 0

        planilha.Columns("H:H").Hidden = False
        planilha.Columns("P:P").Hidden = False
        planilha.Columns("Q:Q").Hidden = False

        for email, data in novos_resultados:
            planilha.Rows(linha_inicial).Insert()

            try:
                # Formatar data no padrão 'DD/MM/YYYY'
                data_formatada = datetime.strptime(data, '%m/%d/%Y').strftime('%d/%m/%Y')
            except ValueError:
                data_formatada = data

            # Inserir os dados
            planilha.Range(f'A{linha_inicial}').Value = Id
            planilha.Range(f'B{linha_inicial}').Value = email
            planilha.Range(f'C{linha_inicial}').Value = data_formatada

            # Inserindo as formulas do robo [Linha H, P, Q]
            planilha.Range(f'H{linha_inicial}').FormulaLocal = f'=ESQUERDA(G{linha_inicial}; 3)'
            planilha.Range(f'P{linha_inicial}').FormulaLocal = f'=SE(O{linha_inicial}="Aprovado";"Aprovado";SE(O{linha_inicial}<>"";"Comunicar";""))'
            planilha.Range(f'Q{linha_inicial}').FormulaLocal = f'=H{linha_inicial}&P{linha_inicial}'

            linha_inicial += 1

        planilha.Columns("H:H").Hidden = True
        planilha.Columns("P:P").Hidden = True
        planilha.Columns("Q:Q").Hidden = True

        excel.ActiveWorkbook.Save()

@app.route('/run_scraper', methods=['POST'])
def run_scraper():
    """Endpoint que executa a função de scraping e atualização da planilha."""
    asyncio.run(main())
    return jsonify({"status": "Scraping executado com sucesso"}), 200

if __name__ == '__main__':
    app.run(debug=True)

import re
import time
import win32com.client
from playwright.async_api import async_playwright
import asyncio
import datetime
from tqdm import tqdm
import os



print(f"""
            
            \033[31m███╗   ██╗███████╗████████╗███████╗██╗     ██╗██╗  ██╗
            ████╗  ██║██╔════╝╚══██╔══╝██╔════╝██║     ██║╚██╗██╔╝
            ██╔██╗ ██║█████╗     ██║   █████╗  ██║     ██║ ╚███╔╝ 
            ██║╚██╗██║██╔══╝     ██║   ██╔══╝  ██║     ██║ ██╔██╗ 
            ██║ ╚████║███████╗   ██║   ██║     ███████╗██║██╔╝ ██╗
            ╚═╝  ╚═══╝╚══════╝   ╚═╝   ╚═╝     ╚══════╝╚═╝╚═╝  ╚═╝\033[0m
            
                            PEGUE SUA PIPOCA!!
""")

async def run():
        # '''
        # 1
        # ======================================================================================
        # ESSA PARTE DO CÓDIGO É RESPONSÁVEL POR INICIAR ALL O PROCESSO, IRÁ ABRIR O GOOGLE,
        # POR SUAS INFORMAÇÕES DE LOGIN E SENHA E CLICAR NO CAMPO DO BOTÃO QUE FAZ A SOLICITAÇÃO
        # DO CÓFIGO PARA ACESSAR O SITE.
        # ======================================================================================
        # '''
    async with async_playwright() as p:
        
        data_atual = datetime.datetime.now()
        username = os.getlogin()
        
        mes = data_atual.month
        ano = data_atual.year
        
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context(color_scheme='dark', record_video_dir='/video')
        page = await browser.new_page()



        await page.goto('https://operacao.ccee.org.br/ui/home')  # Navegue até a página desejada
        # Substitua 'seletor' pelo seletor CSS do elemento
        await page.locator('#mat-input-0').fill('') # INSIRA SEU LOGIN 
        await page.locator("#INPUT").fill('') # INSIRA SUA SENHA
        await page.get_by_role("button", name="Entrar").click()
        
        await page.get_by_role("button", name="Email").click()
        
        '''
        2
        ===================================================================================
        ESSA PARTE DO CÓDIGO PEGA O CÓDIGO NO E-MAIL E COLOCA NO CAMPO PARA ACESSAR A CCEE
        ===================================================================================
        '''
        async def obter_codigo_autorizacao():
            outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
            inbox = outlook.GetDefaultFolder(6)  # "6" refere-se ao índice da caixa de entrada
            messages = inbox.Items
            last_message = None

            time.sleep(45)
            while True:
                try:
                    message = messages.GetLast()
                    if message != last_message:  # Verifica se há um novo email
                        last_message = message
                        # Usando expressão regular para buscar o código de autorização
                        padrao_codigo = re.compile(r"CCEE: o seu codigo de acesso e \s*(\d+)", re.IGNORECASE)
                        match = padrao_codigo.search(message.body)
                        if match:
                            codigo_autorizacao = match.group(1)
                            await page.locator("#mfaCode").fill(codigo_autorizacao) # O CÓDIGO É INSERIDO NESSE CAMPO 
                            await page.get_by_role("button", name="Ok").click()
                            break
                    else:
                        print("Código de autorização não encontrado no email.")
                except Exception as e:
                    print("Red")
                    break
                
                
                
        '''
        3
        ===================================================================================
        ESSA PARTE DO CÓDIGO IRÁ PERCORRER A PÁGINA E CLICAR NOS DEMAIS CAMPOS, ATÉ CHEGAR 
        A PARTE DO DOWLOAD QUE FICA DENTRO DE UM LOOP
        ===================================================================================
        '''
        
        await obter_codigo_autorizacao()
        
        await page.locator("#cdk-drop-list-4").get_by_role("button", name="expand_more Menu de acessos").click()
        await page.get_by_role("menuitem", name="DRI", exact=True).locator("a").click()
        
        await page.frame_locator("iframe").get_by_title('Painéis de Controle').click()

        await page.locator("iframe").content_frame.get_by_label("7.Energia de Reserva").click()
        await page.locator("iframe").content_frame.get_by_label("Energia de Reserva", exact=True).click()
        await page.frame_locator("iframe").get_by_text("RES002 - Energia de Reserva").click()
        
        await page.frame_locator("iframe").get_by_role("textbox", name=f"/{str(12)}").click()
        

        
        for ii in range(1, 13):
            
            mes = str((str(0) + str(((ii +1) - 1)) if ii < 10 else (ii + 1) - 1))
            
            await page.frame_locator("iframe").get_by_title(f"{2024}/{mes}").click()

            time.sleep(2)
            
            await page.frame_locator("iframe").get_by_role("textbox", name=f"/{mes}").press("Tab") 
            time.sleep(2)
            
            await page.frame_locator('iframe').locator('.data').nth(1).click()
                
            await page.frame_locator("iframe").get_by_title(f"2024_{mes}_RECEITA DE VENDA CER", exact=True).click()
            
            # await page.press('Tab')

            # time.sleep(1.5)
            
            textos = [
                    
            "CARNAUBA", "REDUTO", "SANTO CRISTO", "SAO JOAO"
        ]
            
            caminho = f'C:/Users/{username}/VOLTALIA/Common Brazil - Commercialization and Regulation/2 - COMERCIAL/01 Back Office/19 - CCEE/33 - DRI CCEE/2024/Anual'
            
            for i, txt in tqdm(enumerate(textos),desc="Exportando para PDF", total=len(textos),):
                
                await page.frame_locator("iframe").get_by_text(txt, exact=True).click()
                # await page.pause()
                await page.frame_locator("iframe").get_by_title("Aplicar todos os valores").click()
                # await page.frame_locator("iframe").get_by_role("button", name="Aplicar").click()
                time.sleep(3)
                await page.frame_locator("iframe").get_by_role("button", name="Opções de Página").click()
                await page.wait_for_timeout(1000)  
                await page.frame_locator("iframe").get_by_label("Imprimir").click()
                await page.wait_for_timeout(1000)  
                
                async with page.expect_popup() as page1_info:
                    await page.frame_locator("iframe").get_by_text("Página Atual como HTML").click()
                    
                page1 = await page1_info.value
                await page.wait_for_timeout(3000)  
                # await page1.pdf(path=f'{caminho}/RES002 {txt}.pdf')
                await page1.pdf(path=f'{caminho}/{mes}/RES002 {txt} {mes}.2024 .pdf', format='A1', width="100")
                await page1.close()  
                
            await page.frame_locator("iframe").get_by_role("textbox", name=f"/{str((str(0) + str(((ii +1) -1)) if ii < 10 else (ii + 1) - 1))}").click()
            
        print(" ATÉ AQUI NOS AJUDOU O SENHOR!!! ")

        await context.close()
        await browser.close()
        
asyncio.run(run())





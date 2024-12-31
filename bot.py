# Import for the Desktop Bot
from botcity.core import DesktopBot, Backend
#importar blibbliotecas do Excel
from botcity.plugins.excel import BotExcelPlugin
#importar bibliotecas datas
from datetime import date 
# Disable errors if we are not connected to Maestro
#BotMaestroSDK.RAISE_NOT_CONNECTED = False

def main():
   
    #Instanciar o bot
    bot = DesktopBot()
    #instanciar Excel
    bot_excel = BotExcelPlugin()

    #read desktop application Banco
    app_path = "E:\\RPA\BotCity\\Drives\\banco\\Banco.exe"
    bot.execute(app_path)
    #conectar banco ao backend
    title_bank = "Alex Diogo - Bank"
    bot.connect_to_app(Backend.WIN_32, path=app_path, title=title_bank)
    #instanciar o Bank
    bank_window = bot.find_app_window(title=title_bank, class_name="WindowsForms10.Window.8.app.0.141b42a_r7_ad1")
    
    # Caminho para o arquivo Excel
    caminho_excel = "E:\\RPA\\RPA_Studio\\DRIVES\\BankSystem\\extract.xlsx"
    
    # Leitura dos dados do Excel
    dados = bot_excel.read(caminho_excel).as_list("extrato")[1:]
            
    # Mapear os botões de Débito e Crédito
    btn_Debito = bot.find_app_element(from_parent_window=bank_window, auto_id="radioButton_Debito")
    btn_Credito = bot.find_app_element(from_parent_window=bank_window, auto_id="radioButton_Credito")
    # Mapear o campos de valor, data e descrição
    campo_descricao = bot.find_app_element(from_parent_window=bank_window, auto_id="textBox_Descricao")
    campo_valor = bot.find_app_element(from_parent_window=bank_window, auto_id="textBox_Valor") 
    campo_data = bot.find_app_element(from_parent_window=bank_window, auto_id="textBox_Data")   
    # Mapear o botão de Gravar
    btn_gravar = bot.find_app_element(from_parent_window=bank_window, auto_id="button_Gravar")
    
    # Filtrar os dados do Excel
    for linha in dados:
        # Desempacotamento dos valores em variáveis
        tipo, descricao, valor, data = linha
        # Conversão do valor para float
        valor = str(int(float(valor)))
        # Conversão da data para o formato dd/mm/aaaa
        data = date.strftime(data, "%d/%m/%Y")

        # verificar condição se Débito ou Crédito
        if tipo == "Débito":
            btn_Debito.click()
        else:
            btn_Credito.click()
        # Preencher os campos de descrição, valor e data
        campo_descricao.set_text(descricao)
        campo_valor.set_text(str(valor))
        campo_data.set_text(data)   
        btn_gravar.click()

    # Fechar a aplicação     
    process = bot.find_app_window(title=title_bank)
    process.close()
    
def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
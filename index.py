import pandas as pd
import phonenumbers
from phonenumbers import carrier
import urllib.error

def verifica_whatsapp(numero):
    try:
        telefone = phonenumbers.parse(numero)
        number_type = telefone.type
        return carrier._is_mobile(number_type)
    except phonenumbers.NumberFormatException:
        return False

def main():
    # Substitua 'URL_DO_SEU_ARQUIVO_EXCEL' pela URL real do seu arquivo Excel online
    url_arquivo_excel = 'https://1drv.ms/x/s!Ah-BrmRQFNiwhUgRqbb2bD8HFR0V?e=7M5L18'

    try:
        # Tente carregar a planilha Excel online sem especificar o mecanismo
        df = pd.read_excel(url_arquivo_excel, engine='openpyxl')

        # Coluna que contém os números de telefone
        coluna_telefones = 'NumerosTelefone'

        # Adiciona uma nova coluna para indicar se tem WhatsApp ou não
        df['TemWhatsApp'] = df[coluna_telefones].apply(verifica_whatsapp)

        # Substitua 'URL_DO_RESULTADO_EXCEL' pela URL real do seu arquivo Excel de saída
        url_resultado_excel = 'https://1drv.ms/x/s!Ah-BrmRQFNiwhUgRqbb2bD8HFR0V?e=7M5L18'
        
        # Salva o resultado em uma nova planilha Excel online
        df.to_excel(url_resultado_excel, index=False)

    except urllib.error.HTTPError as e:
        if e.code == 503:
            print("Erro 503: O serviço está temporariamente indisponível. Tente novamente mais tarde.")
        else:
            print(f"Erro ao acessar o arquivo Excel online: {e}")

if name == "main":
    main()

import pandas as pd
import phonenumbers
from phonenumbers import carrier

def verifica_whatsapp(numero):
    try:
        telefone = phonenumbers.parse(numero)
        return carrier._is_mobile(number_type:=phonenumbers.number_type(telefone))
    except phonenumbers.NumberFormatException:
        return False

def main():
    # Substitua 'caminho/do/arquivo.xlsx' pelo caminho real do seu arquivo Excel
    caminho_arquivo = 'https://1drv.ms/x/s!Ah-BrmRQFNiwhUgRqbb2bD8HFR0V?e=7M5L18'

    # Carrega a planilha Excel
    df = pd.read_excel(caminho_arquivo)

    # Coluna que contém os números de telefone
    coluna_telefones = 'NumerosTelefone'

    # Adiciona uma nova coluna para indicar se tem WhatsApp ou não
    df['TemWhatsApp'] = df[coluna_telefones].apply(verifica_whatsapp)

    # Salva o resultado em uma nova planilha Excel
    df.to_excel('https://1drv.ms/x/s!Ah-BrmRQFNiwhUq_TTuRNB3LVept?e=iv1z1a', index=False)

if __name__ == "__main__":
    main()

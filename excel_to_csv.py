import pandas as pd

while True:
    salvou = input('Você já salvou o excel na pasta arquivos_excel? [s]im: ').lower().startswith('s')
    if salvou is True:

        nome_do_arquivo = input('Qual o nome do arquivo que deseja converter em csv? ') 

        excel_file = (f"/home/inteligencia.bbc/venvs/responsys/excel_to_csv/arquivos_excel/{nome_do_arquivo}.xlsx") 
        csv_file = (f"/home/inteligencia.bbc/venvs/responsys/excel_to_csv/arquivos_csv/{nome_do_arquivo}.csv")      

        df = pd.read_excel(excel_file)

        df.to_csv(csv_file, index=False, encoding='utf-8')

        print(f"Arquivo convertido com sucesso! ")
        print(f"O arquivo CSV salvo em: {csv_file}")
    
        sair = input('Mais algum arquivo? [s]im: ').lower().startswith('s')
    
        if sair is False:
            print('Então até mais! ')
            break

    else:
        print('Pois salve lá na pastinha')

    

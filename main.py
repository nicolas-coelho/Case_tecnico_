import pandas as pd
import requests
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

#Leitura do arquivo
df = pd.read_excel("data/3-colaboradores.xlsx")

#Eficiência Agi: consulta dos feriados
def obter_feriados_ano(data_exemplo):
    ano = str(data_exemplo).split('-')[0]
    url = f"https://brasilapi.com.br/api/feriados/v1/{ano}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return [f['date'] for f in response.json()]
    except:
        pass
    return []

feriados_cache = obter_feriados_ano(df.iloc[0]['Data_Admissao'])

driver = webdriver.Edge()

resultados_prioridade = []

try:
    for index, linha in df.iterrows():
        # tratamento de dados
        nome = str(linha["Nome"]).strip()
        cep = str(linha["CEP"]).strip().replace("-", "")
        cod_banco = str(linha["Codigo_Banco"]).strip()
        data_adm = str(linha["Data_Admissao"]).split(' ')[0]
        
        # Abordagem RPA utilizarei o selenium
        logradouro_rpa = ""
        try:
            driver.get("https://www.consultarcep.com.br/")
            campo = driver.find_element(By.ID, "q")
            campo.clear()
            campo.send_keys(cep)
            time.sleep(3)
            botao_pesquisar = driver.find_element(By.CLASS_NAME, "submit")
            botao_pesquisar.click()
            time.sleep(3)            
            botao_resultado = driver.find_element(By.XPATH, "(//a[@class='gs-title'])[1]")
            logradouro_rpa = botao_resultado.text.lower()
            print(logradouro_rpa)
        except:
            logradouro_rpa = "erro"


        # Abordagem API (ViaCEP) utlizando os requests
        logradouro_api = ""
        try:
            res_via = requests.get(f"https://viacep.com.br/ws/{cep}/json/")
            dados_via = res_via.json()
            logradouro_api = dados_via.get("logradouro", "").lower()
        except:
            logradouro_api = "vazio"

        #Consulta de Dados Bancários (Brasil API)
        banco_valido = False
        try:
            res_banco = requests.get(f"https://brasilapi.com.br/api/banks/v1/{cod_banco}")
            if res_banco.status_code == 200:
                banco_valido = True
        except:
            banco_valido = False

        # Validação Cruzada e Regra de Negócio
        cep_confirmado = (logradouro_api != "" and logradouro_api in logradouro_rpa)

        if data_adm in feriados_cache:
            prioridade = "BLOQUEADO"
        elif cep_confirmado and banco_valido:
            prioridade = "ALTA"
        else:
            prioridade = "BAIXA"
            
        resultados_prioridade.append(prioridade)
        print(f"Processado: {nome} | Prioridade: {prioridade}")

    df['Prioridade_Agi'] = resultados_prioridade

finally:
    driver.quit()

# Salvando o resultado final
df.to_excel("tests/colaboradores_processados.xlsx", index=False)

"""
Considero a abordagem via API significativamente mais resiliente e mais aplicável que o RPA por três motivos:

1. Estabilidade: APIs fornecem dados estruturados que os JSON que não tem layouts visuais.
   Se o site mudar a cor de um botão ou o nome de uma classe CSS, o RPA quebra, 
   enquanto a API continua funcionando.

2. Performance e Custos: A chamada de API é processada muito mais rápido e consome 
   pouquíssima memória RAM o que em casos mais extensos com muitos funcionarios exigiria muito com computador.
   Além que o RPA carrega imagens, propagandas entre outras coisas que deixam o desempenho lento.

3. Manutenção: O código da API é mais simples e fácil de executar o garante melhores entregas e utilização da ferramenta.
"""
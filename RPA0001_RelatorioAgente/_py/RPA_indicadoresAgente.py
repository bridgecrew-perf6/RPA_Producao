from selenium import webdriver;
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
import pyautogui as pyag
from datetime import date
from datetime import datetime
import pandas as pd
import win32com.client

# ---------------------------------------------------| LOG
# Criado por Thiago Luca em 30/01/2022 10:33
# v 1.0 - Compilada em 31/01/22

# ---------------------------------------------------| REQUISITOS
# variável %USERNAME% funcionando
# pip install selenium
# pip install pandas
# pip install pyautogui
# pip install pypiwin32
# pip install lxml
# chromedriver configurado

# ---------------------------------------------------|FUNÇÕES
# Função Login
def realizaLogin():
    # p-nome é o id do campo de nome do usuário
    # p-senha é o id do campo de senha da tela
    gchrome.find_element(By.ID, "p-nome").send_keys(user);
    gchrome.find_element(By.ID, "p-senha").send_keys(passwd);
    gchrome.find_element(By.ID, "p-senha").send_keys(Keys.RETURN);  # Enter
    print('Login realizado: {}'.format(datetime.now()));
    pyag.sleep(3.5);  #


# Função Selecionar agente
def selecionarAgente(agenteOMNI):
    # id="select2-p-agente-container" >>> id do seletor de agentes
    # Lista de agentes
    # "2551 - I9TELL - BELO HORIZONTE-MG"
    # "1630 - CDCL VISÃO - BELO HORIZONTE-MG"
    # "1631 - CDCL SNV SUL - CURITIBA-PR"
    # "2945 - I9TELL - BELO HORIZONTE-MG"
    # botão azul na tele de seleção de agentes: id="bt-validar"
    print('Agente selecionado: {}'.format(agenteOMNI));
    gchrome.find_element(By.ID, "select2-p-agente-container").click();
    gchrome.find_element(By.CLASS_NAME, 'select2-search__field').send_keys(agenteOMNI);
    gchrome.find_element(By.CLASS_NAME, 'select2-search__field').send_keys(Keys.RETURN);  # Enter
    gchrome.find_element(By.ID, 'bt-validar').click();
    pyag.sleep(3.5);  # lentidão do server da omni
    print('Seleção de agentes fim: {}'.format(datetime.now()));
    return agenteOMNI;


# Função acessar o relatório
def acessarRelatorio():
    # Trecho customizado para o usuário 2551THIAGOL
    # Acessar o relatório Previa da inadinplencia
    # Gerencial
    gchrome.find_element(By.XPATH, '//*[@id="navbar-collapse-1"]/ul[1]/li[1]/a').click();
    pyag.sleep(0.5);
    # Cobrança Amigável
    gchrome.find_element(By.XPATH, '//*[@id="navbar-collapse-1"]/ul[1]/li[1]/ul/li[1]/a').click();
    pyag.sleep(0.5);
    # Relatórios Gerenciais
    gchrome.find_element(By.XPATH, '//*[@id="navbar-collapse-1"]/ul[1]/li[1]/ul/li[1]/ul/li[4]/a').click();
    pyag.sleep(0.5);
    # Previa da inadinplencia
    gchrome.find_element(By.XPATH, '//*[@id="navbar-collapse-1"]/ul[1]/li[1]/ul/li[1]/ul/li[4]/ul/li[5]/a').click();
    pyag.sleep(0.5);
    print('Pagina do relatório acessada');


# função de download de todas as bases dos relatórios
def iteracaoBases():
    # KB Previa da inadinplendcia
    # >> value=1 | INADIMPLÊNCIA 123/x - MATRIZ]
    baixarRelatorios(1);
    # >> value=5 | INADIMPLÊNCIA 123/x - 0 A 30 DIAS
    baixarRelatorios(5);
    # >> value=10 | INADIMPLÊNCIA M3OVER30
    baixarRelatorios(10);
    # >> value=11 | INADIMPLÊNCIA M6OVER60
    baixarRelatorios(11);
    print('Fim dos downloads dos indicadores');
    gchrome.find_element(By.ID, 'logo-omni').click();


# Função para baixar o arquivo de cada base
def baixarRelatorios(valueIX):
    gchrome.switch_to.frame("frm-principal");
    gchrome.find_element(By.NAME, 'p_tipo').click();
    a = gchrome.find_elements(By.XPATH, '/html/body/form/table/tbody/tr[3]/td[2]/select/option');
    rel = a[valueIX - 1].text;
    a[valueIX - 1].click();
    gchrome.find_element(By.ID, 'prn').click();
    pyag.sleep(5);
    gchrome.find_element(By.ID, 'excel').click();
    wait.until(EC.number_of_windows_to_be(2));
    for window_handle in gchrome.window_handles:
        if window_handle != janelaOrigem:
            gchrome.switch_to.window(window_handle)
            pyag.sleep(5);
            gchrome.close();

    # retorna para janela original
    gchrome.switch_to.window(janelaOrigem);
    gchrome.back();
    pyag.sleep(2);
    print('Download da base: {} do agente {}'.format(rel, agenteAtual));


# Função rotina
def funcRotina():
    # Acesso ao relatório customizado para o usuário 2551THIAGOL
    acessarRelatorio();
    pyag.sleep(3.5);
    # Download das 4 bases do agente
    iteracaoBases();
    print('Fim da rotina do agente: {}'.format(agenteAtual));


# Pegar a lista de arquivos baixaods
def funcArquivosBaixados():
    gchrome.switch_to.window(janelaOrigem);
    gchrome.get('chrome://downloads');
    textFromDownloads = gchrome.find_element(By.XPATH, '/html/body');
    # Não sei o pq mas a pagina de download do google chrome não retorna nenhum item, seja por xpath ou por id
    txt_downloads = open(endArquivo, 'w', encoding='utf8');
    txt_downloads.write(textFromDownloads.text);
    txt_downloads.close();
    listaDownloads = [];
    txt_downloads = open(endArquivo, 'r', encoding='utf8');
    for linha in txt_downloads:
        if linha.startswith(prefixoOMNI):
            print('Index: {} Arquivo {}'.format(len(listaDownloads), linha));
            # fazendo o insert do nome do arquivo sem o caractere de enter
            # contempla os 2 servidores da OMNI
            listaDownloads.append(linha[:len(linha) - 1]);
    return listaDownloads;


def exporter(minhaOrigem, meuDestino):
    dataXLS = pd.read_html(minhaOrigem);
    bolinha = dataXLS[0][dataXLS[0].columns].replace(['=', '"'], "", regex=True);
    writer = pd.ExcelWriter(meuDestino, engine='xlsxwriter');
    bolinha.to_excel(writer, sheet_name='Sheet1', index=False);
    writer.save();


# ---------------------------------------------------|PARAMETROS
# apenas para imprimir em qual agente estamos
agenteAtual = 'nenhum';
# Parametros dos arquivos
prefixoOMNI = 'pck_matriz';
endArquivo = 'W:\\txtDownloads.txt';
# dados de acesso
user = '2551THIAGOL';  # melhorar isso com um acesso a um DB? API?
passwd = '01SANTOS2007';  # melhorar isso com um acesso a um DB? e quando a senha mudar?

# prefixo do caminho completo do arquivo, parametrizar conforme o ambiente
prefixDownload = 'C:\\Users\\%USERNAME%\\Downloads\\'
prefixDestino = 'W:\\PLANEJAMENTO\\Relatorios Diarios\\Agente\\Bases\\'

# parametros de exportação dos arquivos

# a ordem tem que ser inversa
ordemFiles = ["OVER60 - ", "over30 - ", "123x_0 a 30_Dias - ", "123x_Matriz - "];
ordemFolders = ['Over60_ATUAL\\ATUAL\\', 'Over30_ATUAL\\ATUAL\\', '123-0_ATUAL\\ATUAL\\', '123xMATRIZ_ATUAL\\ATUAL\\']
ordemFoldersAbertura = ['Over60\\', 'Over30\\', '123-0\\', '123xMATRIZ\\']
ordemAgentes = ["1631", "1630"];

# O script vai utilizar a data de hoje do sistema
dMenos1 = '';
if date.today().day == 1:
    dMenos1 = date(date.today().year, date.today().month, 1);
else:
    dMenos1 = date(date.today().year, date.today().month, date.today().day - 1);

dataFinal = dMenos1.strftime('%d/%m/%Y');
# primeira data do mês
primeirodia = date(date.today().year, date.today().month, 1);
dataInicial = primeirodia.strftime('%d/%m/%Y');
print('Carga de dados de acesso: {}'.format(datetime.now()));

# Lista de arquivos baixados
listaDownloads = [];

# ---------------------------------------------------|> MAIN
print('Inicio: {}'.format(datetime.now()));
# criar o objeto do navegador
gchrome = webdriver.Chrome();

# criação do wait
wait = WebDriverWait(gchrome, 10);

# acessar o site do bizfacil
gchrome.get('https://www.bizfacil.com.br');

# registrando a ID da janela do chrome
janelaOrigem = gchrome.current_window_handle;

print('Acesso ao site: {}'.format(datetime.now()));
pyag.sleep(3.5);

# Realizando o login no site
realizaLogin();

# Selecionando o agente 1630
agenteAtual = selecionarAgente('1630 - CDCL VISÃO - BELO HORIZONTE-MG');
funcRotina();

# Selecionando o agente 1631
agenteAtual = selecionarAgente('1631 - CDCL SNV SUL - CURITIBA-PR');
funcRotina();

# pegar o nome dos arquivos baixados
arquivosXLS = funcArquivosBaixados();
gchrome.quit();  # fechar completamente o chrome
print('Downloads OMNI: {}'.format(datetime.now()));

# abrir os arquivos, salvar nas bases correspondentes
for p1 in ordemAgentes:
    baseIXAgente = ordemAgentes.index(p1) * 4;
    for p2 in ordemFiles:
        baseIXArq = ordemFiles.index(p2) + 1;
        ixXLS = (baseIXAgente + ordemFiles.index(p2));
        nomeArquivo = arquivosXLS[ixXLS];
        caminhoArquivoOrigem = prefixDownload + nomeArquivo;
        caminhoArquivoDestino = prefixDestino + ordemFolders[ordemFiles.index(p2)] + p2 + p1 + '.xlsx';
        exporter(caminhoArquivoOrigem, caminhoArquivoDestino);
        print('Exportação do arquivo: {}\nCaminho: {}\n'.format(p2 + p1 + '.xlsx', caminhoArquivoDestino))

if date.today().day == 1:
    print('Hoje é dia de exportar as bases de arbertura.');
    for p1 in ordemAgentes:
        baseIXAgente = ordemAgentes.index(p1) * 4;
        for p2 in ordemFiles:
            ixXLS = (baseIXAgente + ordemFiles.index(p2));
            nomeArquivo = arquivosXLS[ixXLS];
            caminhoArquivoOrigem = prefixDownload + nomeArquivo;
            caminhoArquivoDestino = prefixDestino + ordemFoldersAbertura[ordemFiles.index(p2)] + p2 + p1 + '.xlsx';
            exporter(caminhoArquivoOrigem, caminhoArquivoDestino);
            print('Exportação do arquivo de BASE: {}\nCaminho: {}\n'.format(p2 + p1 + '.xlsx', caminhoArquivoDestino))
else:
    print('Não é dia de exportar as bases.');

print('Bases salvas: {}'.format(datetime.now()));

#---------------------------| PENDENCIAS PROX VERSÕES
#1- Atualizar o excel
#2- Printar o relatório
#3- Enviar por email

#complier command: pysintaller --onefile RPA_indicadoresAgente.py
#-w? qund usar Janela de interação com o usuário

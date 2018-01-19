##############################IMPORTAR BIBLIOTECAS##############################
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import re
import os
import glob
import math
import shutil
import win32com.client as win32
##############################IMPORTAR BIBLIOTECAS##############################


##############################PARAMETROS##############################
#Login e senha
login = "fabiocimerman@pragmapatrimonio.com.br"
senha = "070595Fc"

#Datas de download
dt_inicial = "06/30/2017"
dt_final= "09/29/2017"

#Lista de empresas a serem baixadas
lista_ciq = "C:\\Users\\dsterenfeld\\Desktop\\ORFEU\\Base\\Download\\Lista - Copia.txt"

#URL das trasncrições
html = "https://www.capitaliq.com/CIQDotNet/Transcripts/Summary.aspx"

#Caminho que está o driver do chrome
cam_exe = "P:\\15. PUBLICO\\AI\\Projetos Python\\1.1.1 Base de Dados\\1. Download\\chromedriver.exe"

#Caminho que o chrome baixa
cam_down = "C:\\Users\\dsterenfeld\\Downloads\\"

#Caminho pasta final
cam_pfinal = "C:\\Users\\dsterenfeld\\Desktop\\ORFEU\\Base\\Earning Calls\\"
##############################PARAMETROS##############################

#Abre Instância de word que vai ser usado para converter de .rtf para .docx
word = win32.gencache.EnsureDispatch('Word.Application')
word.Visible = False

#Apaga os arquivos existentes na pasta de download final
for f in glob.glob(cam_pfinal+"*"):
    os.remove(f)

#Ler a lista de empresas que serão buscadas
input_file = open(lista_ciq, 'r')

#Cria o objeto e tenta entrar no HTML
browser = webdriver.Chrome(cam_exe)
try:
    browser.get(html)
except:
    pass

#Preenche login senha e clicka no botao de login
browser.find_element_by_css_selector("input[id='username']").send_keys(login)
browser.find_element_by_css_selector("input[id='password']").send_keys(senha)
browser.find_element_by_name('myLoginButton').click()

#Começar a usar tempos por causa do delay de carregamento do site
time.sleep(10)

#Loop buscar Empresas na lista e selecionar (se colocar mais tempo nos sleeps abaixo maior a chance de não pular uma empresa, se for menos que 100 empresas recomendo botar 3)
count_lines = 0
for line in input_file:
    time.sleep(3) 
    
    #Escrever o codigo da empresa n ocampo de busca
    browser.find_element_by_name("_criteria$_searchSection$_searchToggle$_criteria__searchSection__searchToggle__entitySearch_searchbox").send_keys(line)   
    time.sleep(3)
    
    #Seta para baixo
    browser.find_element_by_name("_criteria$_searchSection$_searchToggle$_criteria__searchSection__searchToggle__entitySearch_searchbox").send_keys(u'\ue015') 
    time.sleep(3)
    
    #Enter
    browser.find_element_by_name("_criteria$_searchSection$_searchToggle$_criteria__searchSection__searchToggle__entitySearch_searchbox").send_keys(u'\ue007') 
    
    print (line)
    count_lines += 1
print ('numero de empresas:', count_lines)

#Seleciona opção de data inicial e final e preenche data inicial e data final e clicka search
time.sleep(2)
browser.find_element_by_id("_criteria__searchSection__searchToggle__dateRange_myAllHistoryButton").click()
time.sleep(1)
browser.find_element_by_name("_criteria$_searchSection$_searchToggle$_btGo$_saveBtn").click()
time.sleep(10)
browser.execute_script("javascript:__doPostBack('_transcriptsGrid$_dataGrid','Sort$Title')")
time.sleep(10)
#Loop na tabela e ver se os calls já estão salvos na base
#Salva um objeto com a tabela de dados do Capital IQ
tabela = browser.find_elements_by_xpath("//table[@id='_transcriptsGrid__dataGrid']/tbody/tr")
linha = 0
dados_calls = {}
dados_final = {}

#Pega o número total de transcrições e calcula o número de paginas
html_source = browser.page_source
num_transc = re.findall(' of [0-9]?[0-9],?[0-9]?[0-9]?[0-9]?', html_source)
num_transc = (re.findall(r'[0-9]?[0-9],?[0-9]?[0-9]?[0-9]', num_transc[0])[0])
num_transc = int(num_transc.replace(",",""))
num_pag = math.ceil(num_transc/25)

#Esse primeiro loop serve se tiver mais de 25 trasncrições e tiver que mudar de pagina
for g in range(1, num_pag+1):
    if g>1:
        #Sempre que mudar de pagina tem que salvar um objeto novo com a tabela
        tabela = browser.find_elements_by_xpath("//table[@id='_transcriptsGrid__dataGrid']/tbody/tr")
    g=g+1
    linha_tabela =0 
    
    #Loop nas linhas da tabela
    for e in tabela: 
        coluna = 1
        #Loop nas colunas da tabela
        for td in e.find_elements_by_xpath(".//td"):
            #Salva os dados cadastrais dos calls em uma tabela
            dados_calls[linha,coluna] = td.text
            coluna = coluna + 1
            print (td.text)
        #Se for um campo com informações uteis (não campo de pagina)
        if linha_tabela >0 and "Viewing" not in dados_calls[linha,1] :
            a =  dados_calls[linha,3]
            #Salva o CIQ da empresa (as vezes tem varios), busca o que tiver entre parenteses no cadastro
            ciq_empresa = re.findall('\((.*?)\)', a)
            data_call = dados_calls[linha,2]
            print (ciq_empresa)
            #Checka se é earnings call
            if dados_calls[linha,4]=='Earnings Call':
                #Salva o trimestre da empresa, buscando algo que seja formato: "QX XXX"
                trimestre = re.findall('Q[0-9] [0-9][0-9][0-9][0-9]', a)
                print(trimestre)
                #Download de words (usando o link "href" dentro do objeto)
                elems = e.find_elements_by_xpath("//a[@class='binderIcoSprite_doctype_word']")
                browser.get(elems[linha_tabela].get_attribute("href"))
                time.sleep(5)
                
                #Salva os dados dos calls
                dados_final[linha,0] = data_call #Data do call
                dados_final[linha,1] = trimestre #Trimestre de referencia
                dados_final[linha,2] = ciq_empresa #Codigo da empresa
                nome = str(dados_final[linha,2])
                dados_final[linha,3] = dados_calls[linha,4] #Tipo de call (Earnings)
                dados_final[linha,4] = dados_calls[linha,3] #Descricao de call
                dados_final[linha,5] = cam_down + str(re.split('\n', dados_calls[linha,3])[0]).replace("/","-") + ".rtf" #Caminho que vai ser salvo
                
                caminho_final1 = str(dados_final[linha,1])
                caminho_final2 = str(dados_final[linha,0]).replace(":","-")
                caminho_final3 = nome.replace(":","-")
                caminho_final3 = caminho_final3.replace("/","-")
                caminho_final = cam_pfinal + caminho_final1.replace("'","") + "_" + caminho_final2.replace("'","") +"_" + caminho_final3.replace("'","") + ".rtf"
                
                #Copia o arquivo para pasta destino (se der erro espera mais 5s e copia e depois continua o resto do código)
                while True:
                    try:                      
                        shutil.copy(str(dados_final[linha,5]), caminho_final)
                        break
                    except:
                        time.sleep(2)
                        continue
#                try:
#                    time.sleep(3)
#                    shutil.copy(str(dados_final[linha,5]), caminho_final)   
#                except:
#                    try:
#                        time.sleep(30)
#                        shutil.copy(str(dados_final[linha,5]), caminho_final) 
#                        pass
#                    except:
#                        time.sleep(100)
#                        shutil.copy(str(dados_final[linha,5]), caminho_final) 
#                        pass
                
                #Salva o arquivo no formato correto(docx)
                word.Documents.Open(caminho_final)
                word.ActiveDocument.SaveAs2(cam_pfinal + caminho_final1.replace("'","") + "_" + caminho_final2.replace("'","") +"_" + caminho_final3.replace("'","") + ".docx",16)
                word.ActiveDocument.Close()
                
        #Se tiver na última linha, mudar de página
        if linha_tabela==26:
            browser.execute_script("javascript:__doPostBack('_transcriptsGrid$_dataGrid','Page$" + str(g) + "')")
            time.sleep(50)
        
        #Tem esses dois contadores, um que zera toda vez que muda de pagina e outro que vai até o final
        linha = linha + 1
        linha_tabela=linha_tabela+1

#Apaga os arquivos rtf (pq agora vamos usar os docx)
for f in glob.glob(cam_pfinal + "*.rtf"):
    os.remove(f)
    
##############################FUNÇÕES##############################
##############################FUNÇÕES##############################


##############################TO-DO##############################
####Falta check se ja foi importado (criar um log? olhar na base ver se ja existe?)
####Ajustar Tempos (sleep)
##############################TO-DO##############################
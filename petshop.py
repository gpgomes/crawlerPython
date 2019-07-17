from bs4 import BeautifulSoup
from selenium import webdriver
import datetime
import time
from urllib.request import urlopen
import xlsxwriter
import requests

class Produto:
    def __init__(self, nome, link):
        self.nome = nome
        self.link = link
        self.codItem = ''
        self.preco = ''
        self.detalhes = ''

    def SetCodItem(self, codItem):
        self.codItem = codItem

    def SetPreco(self, preco):
        self.preco = preco

    def SetDetalhes(self, detalhes):
        self.detalhes = detalhes

def log(mensagem):
    hora_atual = time.time()
    st = datetime.datetime.fromtimestamp(hora_atual).strftime('%Y-%m-%d %H:%M:%S')
    print('[%s] %s' % (st, mensagem))

def buscaPaginasProdutos(pesquisa):
    # Usando o selenium para facilitar a pesquisa e obter as páginas da listagem dos produtos
    url = 'https://www.extra.com.br/'

    options = webdriver.ChromeOptions()
    #options.add_argument('--headless')
    options.add_argument('--disable-extensions')
    options.add_argument("--window-size=1920x1080")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-gpu")

    try:
        drive = webdriver.Chrome(options = options)
        drive.implicitly_wait(10)
        drive.set_page_load_timeout(30)

        try:
            drive.get(url)

            busca = drive.find_element_by_id('ctl00_TopBar_PaginaSistemaArea1_ctl05_ctl00_txtBusca')
            busca.send_keys(pesquisa)

            botao_buscar = drive.find_element_by_id('ctl00_TopBar_PaginaSistemaArea1_ctl05_ctl00_btnOK')
            botao_buscar.click()

            novaUrl = drive.current_url
            log(novaUrl)

            linkCategoria = drive.find_element_by_partial_link_text(pesquisa)
            linkCategoria.click()

            novaUrl = drive.current_url

            drive.quit()

            return novaUrl
        except:
            drive.quit()
            log("Erro ao buscar categoria pelo selenium")
            return ''

    except Exception as E:
        print(E)
        log("Erro ao iniciar drive")


def buscaURLs(url):
    html = urlopen(url)
    soup = BeautifulSoup(html, 'html.parser')
    div_lista_paginas = soup.findAll("li", {"class" : "neemu-pagination-inner"})
    if len(div_lista_paginas) > 1:
        link = div_lista_paginas[1].findAll("a", href=True)
        return ("https:"+link[0]['href'][0:-1])
    else:
        return ''

def busca_produtos(urlPaginasProdutos, num_pagina):
    url = urlPaginasProdutos + str(num_pagina)
    html = urlopen(url)
    soup = BeautifulSoup(html, 'html.parser')
    lista_produtos = []
    div_produtos = soup.findAll("div", {"class": "nm-product-info"})

    for div_produto in div_produtos:
        div_nome = div_produto.findAll("div", {"class": "nm-product-name"})
        for nome in div_nome:
            links = nome.find_all('a', href=True)
            for link in links:
                lista_produtos.append(Produto(link.text, 'https:' + link['href']))

    return lista_produtos

def LimpaCodItem(texto):
    texto = texto.replace('(', '').replace(')', '').replace('Cód. Item', '').replace(' ', '')
    return texto

def buscaDetalhes(produto):
    url = produto.link
    user_agent = {'User-agent': 'Mozilla/5.0'}
    html = requests.get(url, headers=user_agent)
    soup = BeautifulSoup(html.text, 'html.parser')

    #Procurando Cód. Item
    divProdutoNome = soup.findAll("div", {"class": "produtoNome"})
    for produtoNome in divProdutoNome:
        codItem = produtoNome.findAll("span", {"itemprop" : "productID"})
        if (len(codItem) == 1):
            produto.SetCodItem(LimpaCodItem(codItem[0].text))

    #Procurando Preço
    divPreco = soup.findAll("strong", {"id" : "ctl00_Conteudo_ctl00_precoPorValue"})
    for detalhesPreco in divPreco:
        for detalhePreco in detalhesPreco:
            if (detalhePreco != "R$"):
                produto.SetPreco("R$" + detalhePreco.text)

    #Procurando pela descrição do produto
    divDescricao = soup.findAll("div", {"id" : "descricao"})
    for descricao in divDescricao:
        produto.SetDetalhes(descricao.text.strip())


def criaPlanilhas(listaProdutos):
    workbook = xlsxwriter.Workbook('produtos.xlsx')
    worksheet = workbook.add_worksheet()

    worksheet.write(0, 0, "Nome")
    worksheet.write(0, 1, "Cód. Item")
    worksheet.write(0, 2, "Preço")
    worksheet.write(0, 3, "Detalhes")

    linha = 1

    for produto in listaProdutos:
        worksheet.write(linha, 0, produto.nome)
        worksheet.write(linha, 1, produto.codItem)
        worksheet.write(linha, 2, produto.preco)
        worksheet.write(linha, 3, produto.detalhes)
        linha += 1

    workbook.close()

if __name__ == "__main__":
    # Fazendo a pesquisa pela categoria "Pet Shop"
    pesquisa = 'Pet Shop'

    log("Realizando filtro da categoria")
    url = buscaPaginasProdutos(pesquisa)

    # Buscando como ficam a urls quando mudamos de pagina
    log("Buscando paginas de produtos")

    if url == '':
        # caso o selenium não consiga buscar a página de filtro, como o servidor bloqueando testes automatizados
        url = 'https://buscando2.extra.com.br/busca?q=' + pesquisa.replace(' ', '+')

    # aplicando filtro de categoria
    urlPaginasProdutos = buscaURLs(url)

    if urlPaginasProdutos != '':
        log("Buscando produtos")
        listaProdutos = []
        paginaFinal = 20

        for i in range(1, paginaFinal + 1):
            log("Buscando produtos da página %d" % i)
            listaProdutos = listaProdutos + (busca_produtos(urlPaginasProdutos, i))

        if len(listaProdutos) > 0:
            log("Total de %d produtos encontrados" % len(listaProdutos))
            contador = 0
            for produto in listaProdutos:
                log("Busca detalhes do produto n. %d" % contador)
                buscaDetalhes(produto)
                contador += 1

            log("Gerando planilha de produtos")
            criaPlanilhas(listaProdutos)
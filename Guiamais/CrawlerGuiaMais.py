import urllib
import urllib2
import sys
import re
from bs4 import BeautifulSoup
from tempfile import TemporaryFile
from xlwt import Workbook, easyxf
from xlrd import open_workbook,XL_CELL_TEXT
from unicodedata import normalize
from os.path import join, dirname, abspath, isfile

# a ideia principal desse bot eh fazer pesquisa no site do guiamais.com.br com relacao as empresas contendo local e telefone para contato
# facilita a formacao de de listas para trabalhar com telemarketing
# basta abrir o programa e abrir o programa e dizer o que se deseja pesquisar no guiamais
# necessario LISTA-DE-MUNICIPIOS-SIGMA.xlsx para ser realizada a filtragem de quais cidades devem ser pesquisadas

def main():
    #Leitura dos municipios
    Freq = 2500 # Set Frequency To 2500 Hertz --- Dur = 500 ou 3000
    listaMunicipios = [] #inicializa lista
    listaUF=[]
    listaSeg = []
    bookMunicipios = open_workbook('Filtrar.xlsx') # acessa a planilha com os municipios importantes
    sheet = bookMunicipios.sheet_by_index(0)
    for i in range(0,sheet.nrows):
        cell = sheet.cell(i,0)
        listaMunicipios = listaMunicipios + [cell.value.strip()] # le e adiciona na lista
        listaUF = listaUF + [sheet.cell(i,1).value.strip()]
    ############################################################
    sheet = bookMunicipios.sheet_by_index(1)
    for i in range(0,sheet.nrows):
        cell = sheet.cell(i,0)
        listaSeg = listaSeg + [cell.value.strip()] # le e adiciona na lista
    # interface com o pesquisador (guiamais)
    # pesquisa = raw_input('O que deseja pesquisar?') 
    
    telList = []
    #criacao da planilha   
    titleStyle = easyxf('alignment: horizontal center;' 'font:bold True;')
    colStyle = easyxf('alignment: horizontal center;')
    for j in range(0,len(listaSeg)):
        pesquisa = normalize('NFKD',listaSeg[j]).encode('ASCII','ignore').decode('ASCII')
        urlpesquisa = pesquisa.lower().replace(" ","+")
    ##############################################
        for i in range(0,len(listaMunicipios)):

            cidadeSemAcento= normalize('NFKD',unicode(listaMunicipios[i])).encode('ASCII','ignore').decode('ASCII')
            url = 'http://www.guiamais.com.br/encontre?searchbox=true&what=' + urlpesquisa +"&where=" +cidadeSemAcento.strip().lower().replace(" ","+")+"%"+"2C+"+listaUF[i].lower()
            print url
            print listaMunicipios[i]
            if(not isfile(pesquisa+ " "+ listaMunicipios[i] +'.xls')):
                count = 1
                book = Workbook()
                lista = book.add_sheet('lista')
                lista.write(0,0,'Indice',titleStyle)
                lista.write(0,1,'Empresa',titleStyle)
                lista.write(0,2,'Telefone',titleStyle)
                lista.write(0,3,'Cidade',titleStyle)
                lista.write(0,4,'UF',titleStyle)
                lista.col(0).width = 2000
                lista.col(1).width = 15000
                lista.col(2).width = 10000
                lista.col(3).width = 5000
                lista.col(4).width = 1000
                ###################################################################
                # interface com o site
                try:
                    req = urllib2.Request(url)
                    response = urllib2.urlopen(req)
                    the_page = response.read()
                    soup = BeautifulSoup(the_page)
                    ###################################
                    # indice inicial para a leitura das empresas
                    for listshop in soup.findAll('div', {"itemtype":"http://schema.org/LocalBusiness"}):   
                        link = listshop.h2.a['href']
                        pattern = re.compile(u'<\/?\w+\s*[^>]*?\/?>', re.DOTALL | re.MULTILINE | re.IGNORECASE | re.UNICODE)
                        tel = listshop.find('ul',{"class":"advPhone"}).text.replace("ver telefone","").strip().replace("\n","")
                        telefone = normalize('NFKD',unicode(tel)).encode('ASCII','ignore').decode('ASCII').replace(" ","").replace(")",") ").replace("LigueGratis","").replace("("," (")
                        if(listshop.find('span',{"itemprop":"addressLocality"}).text in listaMunicipios and telefone not in telList): #funciona porque ambas a variaveis sao unicode             
                            lista.write(count,0,count,colStyle)
                            lista.write(count,1,listshop.h2.text.strip(),colStyle)
                            lista.write(count,3,listshop.find('span',{"itemprop":"addressLocality"}).text,colStyle)
                            lista.write(count,4,listshop.find('span',{"itemprop":"addressRegion"}).text,colStyle)
                            lista.write(count,2,telefone,colStyle)
                            telList = telList + [telefone]
                            print "  " +listshop.h2.text.strip()
                            print "  " +listshop.find('span',{"itemprop":"addressLocality"}).text
                            print "  " +listshop.find('span',{"itemprop":"addressRegion"}).text
                            print "  " +telefone
                            count= 1+count
                    nextsite = soup.find('link', {"rel":"next"})
                    if(nextsite != None):
                    	url = nextsite['href']
                        req = urllib2.Request(url)
                        response = urllib2.urlopen(req)
                        the_page = response.read()
                        soup = BeautifulSoup(the_page)
                    while (nextsite != None):
                        print nextsite['href']
                        #if(countpage != 1): # nao conta a primeira leitura uma vez que ja foi vista a pagina antes do while
                        for listshop in soup.findAll('div', {"itemtype":"http://schema.org/LocalBusiness"}):   
                            link = listshop.h2.a['href']
                            pattern = re.compile(u'<\/?\w+\s*[^>]*?\/?>', re.DOTALL | re.MULTILINE | re.IGNORECASE | re.UNICODE)
                            tel = listshop.find('ul',{"class":"advPhone"}).text.replace("ver telefone","").strip().replace("\n","")
                            telefone = normalize('NFKD',unicode(tel)).encode('ASCII','ignore').decode('ASCII').replace(" ","").replace(")",") ").replace("LigueGratis","").replace("("," (")
                            #telefone = pattern.sub(u" ", tel).replace(" ","").replace(")",") ")
                            if(listshop.find('span',{"itemprop":"addressLocality"}).text in listaMunicipios and telefone not in telList): #funciona porque ambas a variaveis sao unicode             
                                  lista.write(count,0,count,colStyle)
                                  lista.write(count,1,listshop.h2.text.strip(),colStyle)
                                  lista.write(count,3,listshop.find('span',{"itemprop":"addressLocality"}).text,colStyle)
                                  lista.write(count,4,listshop.find('span',{"itemprop":"addressRegion"}).text,colStyle)
                                  lista.write(count,2,telefone,colStyle)
                                  telList = telList + [telefone]
                                  #print "  " +listshop.h2.text.strip()
                                  print "  " +listshop.find('span',{"itemprop":"addressLocality"}).text
                                  print "  " +listshop.find('span',{"itemprop":"addressRegion"}).text
                                  print "  " +telefone
                                  count= 1+count
                        nextsite = soup.find('link', {"rel":"next"})
                        if nextsite != None:
                            newurl =nextsite['href']
                            #countpage = countpage + 1
                            try:
                                req = urllib2.Request(newurl)
                                response = urllib2.urlopen(req)
                                the_page = response.read()
                                soup = BeautifulSoup(the_page)
                          # print " " + soup.find('div', {"class":"resultPage"}).text
                            except urllib2.HTTPError as err:
                                print 'erro: voce pode ficar com menos informacao' #avisa o erro
                                nextsite = None
                                #winsound.Beep(Freq,3000)
                except urllib2.HTTPError as err:
                            print 'erro: ao conectar ao link da proxima cidade sera passada para outra cidade' #avisa o erro
                            nextsite = None
                            #winsound.Beep(Freq,3000)
                #winsound.Beep(Freq,500)
                book.save(pesquisa+ " "+ listaMunicipios[i] +'.xls') #salva o arquivo
        
        book = Workbook()
        lista = book.add_sheet('lista')
        lista.write(0,0,'Indice',titleStyle)
        lista.write(0,1,'Empresa',titleStyle)
        lista.write(0,2,'Telefone',titleStyle)
        lista.write(0,3,'Cidade',titleStyle)
        lista.write(0,4,'UF',titleStyle)
        lista.col(0).width = 2000
        lista.col(1).width = 15000
        lista.col(2).width = 10000
        lista.col(3).width = 5000
        lista.col(4).width = 1000
        count = 1
        telList = []
        for i in range(0,len(listaMunicipios)):
            bookMunicipios = open_workbook(pesquisa+ " "+ listaMunicipios[i] +'.xls') # acessa a planilha com os municipios importantes
            sheet = bookMunicipios.sheet_by_index(0)
            for i in range(1,sheet.nrows):
                if (sheet.cell(i,2).value not in telList):
                    lista.write(count,0,count,colStyle)
                    lista.write(count,1, sheet.cell(i,1).value)
                    lista.write(count,3, sheet.cell(i,3).value,colStyle)
                    lista.write(count,4,sheet.cell(i,4).value,colStyle)
                    lista.write(count,2,sheet.cell(i,2).value,colStyle)
                    telList = telList + [sheet.cell(i,2).value]
                    count = count + 1
        book.save(pesquisa+ 'Geral.xls')
        #winsound.Beep(Freq,500)
    sys.exit()    

if __name__ == "__main__":
    main()
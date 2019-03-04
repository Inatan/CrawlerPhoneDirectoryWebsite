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


def findNumber(link):
    decCod= ['f','e','d','c','b','a','9','8','7','6']
    uniCod= ['d','c','f','e','9','8','b','a','5','4']
    for i in range(0,len(decCod)):
        if decCod[i] == link[45].lower():
            dec = str(i)
    for i in range(0,len(uniCod)):
        if uniCod[i] == link[47].lower():
            un = str(i)
    val = dec+un 
    return val

def main():
    #Leitura dos municipios
    listaUF=['rs','sc','pr','sp']
    ############################################################
    # interface com o pesquisador (teleListas)
    #pesquisa = raw_input('O que deseja pesquisar?') 
    #pesquisa = pesquisa.strip()
    #urlpesquisa = pesquisa.lower().replace(" ","+")
    telList = []
    reload(sys)
    sys.setdefaultencoding('latin-1')
    #criacao da planilha   
    titleStyle = easyxf('alignment: horizontal center;' 'font:bold True;')
    colStyle = easyxf('alignment: horizontal center;')
    listaSeg=[]
    with open('Segmentos tele listas.txt') as lines:
        listaSeg = lines.readlines()
    
    for j in range(0,len(listaSeg)):
        pesquisa = normalize('NFKD',unicode(listaSeg[j]).strip()).encode('ASCII','ignore').decode('ASCII').strip()
        urlpesquisa = pesquisa.lower().replace(" ","+")
        ##############################################
        for i in range(0,len(listaUF)):
            url = 'http://www.telelistas.net/'+ listaUF[i] +'/cidade/' + urlpesquisa
            print url
            if(not isfile(pesquisa+ " "+ listaUF[i] +'.xls')):
                count = 1
                book = Workbook()
                lista = book.add_sheet('lista')
                lista.write(0,0,'Indice',titleStyle)
                lista.write(0,1,'Empresa',titleStyle)
                lista.write(0,2,'Cidade',titleStyle)
                lista.write(0,3,'Endereco',titleStyle)
                lista.write(0,4,'Numero',titleStyle)
                lista.write(0,5,'Bairro',titleStyle)
                lista.write(0,6,'CEP',titleStyle)
                lista.write(0,7,'DDD',titleStyle)
                lista.write(0,8,'Telefone',titleStyle)
                lista.col(0).width = 2000
                lista.col(1).width = 15000
                lista.col(2).width = 5000
                lista.col(3).width = 15000
                lista.col(4).width = 2500
                lista.col(5).width = 10000
                lista.col(6).width = 10000
                lista.col(7).width = 1000
                lista.col(8).width = 10000
                ###################################################################
                # interface com o site
                req = urllib2.Request(url)
                response = urllib2.urlopen(req)
                the_page = response.read()
                soup = BeautifulSoup(the_page)
                link = None
                ###################################
                # indice inicial para a leitura das empresas
                linkpage=soup
                soup= soup.find('div',{"id":"Content_Regs"})
                #soup= soup.parent

                #print len(soup.findAll('table',{"width":"468"}))
                cnt=1
                if soup is not None:
                    for listshop in soup.findAll('table',{"width":"468"}):
                        print cnt
                        cnt=cnt+1                      
                        if listshop.find('td',{"width":"324"}) is not None: 
                            name= listshop.find('td',{"width":"324"}).text
                            link=listshop.find('td',{"width":"324"})
                            link='http:' + link.a['href']
                            end= listshop.findAll('td',{"width":"294"})[len(listshop.findAll('td',{"width":"294"}))-1].text.strip().replace("\r\n",",")
                        if listshop.find('td',{"width":"414"}) is not None: 
                            name= listshop.find('td',{"width":"414"}).text
                            print name
                            link=listshop.find('td',{"width":"414"})
                            link='http:'+link.a['href']
                            end= listshop.findAll('td',{"width":"345"})[len(listshop.findAll('td',{"width":"345"}))-2].text.strip().replace("\r\n",",")
                            print end
                        print len(listshop.findAll('td',{"width":"294"}))
                        if(link is not None):
                        #print end
                            lstEnd=re.split(' - +|\,+',end,)
                            print lstEnd
                            if  len(lstEnd) == 6:
                                #print 'oipoio'
                                print link
                                req2 = urllib2.Request(link)
                                response2 = urllib2.urlopen(req2)
                                page2 = response2.read()
                                soup2 = BeautifulSoup(page2)
                                tel = soup2.find('div',{"id":"telInfo"})
                                if tel != None and ('Tel:' in tel.div.text or 'Cel:' in tel.div.text) :
                                    tel = tel.div
                                    telefone= tel.text.partition("|")[0].replace("Tel: ","").replace("Cel:","").replace(" ","").strip() + findNumber(tel.img['src']).replace(" ","").strip()
                                    ddd=telefone.split(')')[0].replace("(","")
                                    telnum=telefone.split(')')[1]
                                    if(telefone not in telList): #funciona porque ambas a variaveis sao unicode             
                                        lista.write(count,0,count,colStyle)
                                        lista.write(count,1,name.strip(),colStyle)
                                        lista.write(count,2,lstEnd[3].strip(),colStyle)
                                        lista.write(count,3,lstEnd[0].strip(),colStyle)
                                        lista.write(count,4,lstEnd[1].strip(),colStyle)
                                        lista.write(count,5,lstEnd[2].strip(),colStyle)
                                        lista.write(count,6,lstEnd[5].replace("CEP:","").strip(),colStyle)
                                        lista.write(count,7,ddd,colStyle)
                                        lista.write(count,8,telnum,colStyle)
                                        telList = telList + [telnum]
                                        print "  " +name
                                        print "  " +lstEnd[3].strip()
                                        print "  " +lstEnd[0].strip()
                                        print "  " +lstEnd[1].strip()
                                        print "  " +lstEnd[2].strip()
                                        print "  " +lstEnd[5].replace("CEP:","")
                                        print "  " +ddd
                                        print "  " +telnum
                                        count= 1+count
                                #print 'erro de resposta do servidor'
                    nextsite = linkpage.find('img', {"src":"//imgs.telelistas.net/img/por_rodape_prox.gif"})
                    if(nextsite != None):
                        nextsite = nextsite.parent
                        print 'http://www.telelistas.net' + nextsite['href']
                        url = 'http://www.telelistas.net'+ nextsite['href']
                        req = urllib2.Request(url)
                        response = urllib2.urlopen(req)
                        the_page = response.read()
                        soup = BeautifulSoup(the_page)
                    while (nextsite != None):
                        linkpage=soup
                        soup= soup.find('div',{"id":"Content_Regs"})
                        #print len(soup.findAll('table',{"width":"468"}))
                        for listshop in soup.findAll('table',{"width":"468"}):
                            print cnt
                            cnt=cnt+1
                            if listshop.find('td',{"width":"324"}) is not None: 
                                name= listshop.find('td',{"width":"324"}).text
                                link=listshop.find('td',{"width":"324"})
                                link='http:' +link.a['href']
                                end= listshop.findAll('td',{"width":"294"})[len(listshop.findAll('td',{"width":"294"}))-1].text.strip().replace("\r\n",",")
                            if listshop.find('td',{"width":"414"}) is not None: 
                                name= listshop.find('td',{"width":"414"}).text
                                print name
                                link=listshop.find('td',{"width":"414"})
                                link='http:' + link.a['href']
                                end= listshop.findAll('td',{"width":"345"})[len(listshop.findAll('td',{"width":"345"}))-2].text.strip().replace("\r\n",",")
                                print end
                            print len(listshop.findAll('td',{"width":"294"}))
                            if(link is not None):
                            #print end
                                lstEnd=re.split(' - +|\,+',end,)
                                print lstEnd
                                if  len(lstEnd) == 6:
                                    try:
                                        req2 = urllib2.Request(link)
                                        response2 = urllib2.urlopen(req2)
                                        page2 = response2.read()
                                        soup2 = BeautifulSoup(page2)
                                        tel = soup2.find('div',{"id":"telInfo"})
                                        if tel != None:
                                            tel = tel.div
                                            telefone= tel.text.partition("|")[0].replace("Tel: ","").replace(" ","").strip() + findNumber(tel.img['src']).replace(" ","").strip()
                                            print telefone
                                            ddd=telefone.split(')')[0].replace("(","")
                                            telnum=telefone.split(')')[1]
                                            if(telefone not in telList): #funciona porque ambas a variaveis sao unicode             
                                                lista.write(count,0,count,colStyle)
                                                lista.write(count,1,name.strip(),colStyle)
                                                lista.write(count,2,lstEnd[3].strip(),colStyle)
                                                lista.write(count,3,lstEnd[0].strip(),colStyle)
                                                lista.write(count,4,lstEnd[1].strip(),colStyle)
                                                lista.write(count,5,lstEnd[2].strip(),colStyle)
                                                lista.write(count,6,lstEnd[5].replace("CEP:","").strip(),colStyle)
                                                lista.write(count,7,ddd,colStyle)
                                                lista.write(count,8,telnum,colStyle)
                                                telList = telList + [telnum]
                                                print "  " +name.strip()
                                                print "  " +lstEnd[3].strip()
                                                print "  " +lstEnd[0].strip()
                                                print "  " +lstEnd[1].strip()
                                                print "  " +lstEnd[2].strip()
                                                print "  " +lstEnd[5].replace("CEP:","")
                                                print "  " +ddd
                                                print "  " +telnum
                                                count= 1+count
                                    except urllib2.HTTPError as err:
                                        print 'erro consulta de ' + pesquisa #avisa o erro
                                    #except: 
                                        #print 'erro de resposta do servidor'
                        nextsite = linkpage.find('img', {"src":"//imgs.telelistas.net/img/por_rodape_prox.gif"})
                        if nextsite != None:    
                            nextsite = nextsite.parent
                            newurl = 'http://www.telelistas.net'+ nextsite['href']
                            try:
                                req = urllib2.Request(newurl)
                                response = urllib2.urlopen(req)
                                the_page = response.read()
                                soup = BeautifulSoup(the_page)
                                print newurl
                            except urllib2.HTTPError as err:
                                print 'erro: voce pode ficar com menos informacao' #avisa o erro
                                nextsite = None
                book.save(pesquisa+ " "+ listaUF[i] +'.xls') #salva o arquivo
        
        book = Workbook()
        lista = book.add_sheet('lista')
        lista.write(0,0,'Indice',titleStyle)
        lista.write(0,1,'Empresa',titleStyle)
        lista.write(0,2,'Cidade',titleStyle)
        lista.write(0,3,'Endereco',titleStyle)
        lista.write(0,4,'Numero',titleStyle)
        lista.write(0,5,'Bairro',titleStyle)
        lista.write(0,6,'CEP',titleStyle)
        lista.write(0,7,'DDD',titleStyle)
        lista.write(0,8,'Telefone',titleStyle)
        lista.col(0).width = 2000
        lista.col(1).width = 15000
        lista.col(2).width = 5000
        lista.col(3).width = 15000
        lista.col(4).width = 2500
        lista.col(5).width = 10000
        lista.col(6).width = 10000
        lista.col(7).width = 1000
        lista.col(8).width = 10000
        telList = []
        count = 1
        for i in range(0,len(listaUF)):
            bookMunicipios = open_workbook(pesquisa+ " "+ listaUF[i] +'.xls') # acessa a planilha com os municipios importantes
            sheet = bookMunicipios.sheet_by_index(0)
            for i in range(1,sheet.nrows):
                if (sheet.cell(i,7).value not in telList):
                    lista.write(count,0,count,colStyle)
                    lista.write(count,1, sheet.cell(i,1).value)
                    lista.write(count,3, sheet.cell(i,3).value,colStyle)
                    lista.write(count,4,sheet.cell(i,4).value,colStyle)
                    lista.write(count,2,sheet.cell(i,2).value,colStyle)
                    lista.write(count,5,sheet.cell(i,5).value,colStyle)
                    lista.write(count,6,sheet.cell(i,6).value,colStyle)
                    lista.write(count,7,sheet.cell(i,7).value,colStyle)
                    telList = telList + [sheet.cell(i,7).value]
                    count = count + 1
        book.save(pesquisa+ 'Geral.xls')
    sys.exit()       

if __name__ == "__main__":
    main()

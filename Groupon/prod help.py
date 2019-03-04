for j in range(0,len(listaProdSegmentos)):

            if(not isfile(listaProdSegmentos[j] +'.xls')):

                count = 1
                book = Workbook()
                lista = book.add_sheet('lista')
                lista.write(0,0,'Indice',titleStyle)
                lista.write(0,1,'Empresa',titleStyle)
                lista.write(0,2,'Telefone',titleStyle)
                lista.write(0,3,'Cidade',titleStyle)
                lista.write(0,4,'Promocao',titleStyle)
                lista.write(0,5,'Preco',titleStyle)
                lista.write(0,7,'OFF em %',titleStyle)
                lista.write(0,6,'Preco original',titleStyle)
                lista.col(0).width = 2000
                lista.col(1).width = 15000
                lista.col(2).width = 10000
                lista.col(3).width = 5000
                lista.col(4).width = 30000
                lista.col(5).width = 5000
                lista.col(6).width = 5000
                lista.col(7).width = 5000
                print listaProdSegmentos[j] + '\n' + listaProdLinks[j]
                ###################################################################
                # interface com o site
                try:
                    req2 = urllib2.Request(listaProdLinks[j])
                    #req2 = urllib.request.urlretrieve(req2) 
                    response2 = urllib2.urlopen(req2)
                    the_page2 = response2.read()
                    soup2 = BeautifulSoup(the_page2)    
                    print soup2.find('div',{"id":"deal-space"}).text 
                    print soup2.find('figure',{"class":"deal-card deal-list-tile deal-tile deal-tile-standard"})
                    for proposta in soup2.findAll('figure',{"class":"deal-card deal-list-tile deal-tile deal-tile-standard"}):
                        #print proposta.find('p',{"class":"merchant-name should-truncate "}).text
                        print proposta.a['href'].replace("//","https://")
                        linkProposta= proposta.a['href'].replace("//","http://")
                        ###################################
                        # indice inicial para a leitura das empresas
                        #try:
                        req3 = urllib2.Request(linkProposta)
                        response3 = urllib2.urlopen(req3)
                        the_page3 = response3.read()
                        soup3 = BeautifulSoup(the_page3)
                        pattern = re.compile(u'<\/?\w+\s*[^>]*?\/?>', re.DOTALL | re.MULTILINE | re.IGNORECASE | re.UNICODE)
                        tel =  re.findall(r'\(\d{2}\) *\d{4}\-?\d{4}|\d{2} \d{4}\-?\d{4}' , soup3.text)
                        print tel
                        if(tel == []):
                            telefone = '-'
                        else:
                            telefone = pattern.sub(u" ", tel[0]).replace(" ","").replace(")",") ")
                        
                        tituloProposta=soup3.h5.text
                        valorPromocional= proposta.find('s',{"class":"discount-price"}).text
                        val1= re.search('\d*\.?\d+(\,\d{2})?',valorPromocional ).group(0).replace('.','').replace(',','.')
                        if(proposta.find('s',{"class":"original-price"}).text is not None and proposta.find('s',{"class":"original-price"}).text != ''):
                            valorReal= proposta.find('s',{"class":"original-price"}).text
                            val2= re.search('\d*\.?\d+(\,\d{2})?',valorReal ).group(0).replace('.','').replace(',','.')
                            percentagem = str(float((float(val1)/float(val2))*100))
                        else:
                            percentagem= '-'
                            valorReal='-'
                        if(telefone not in telList or telefone == '-'): #funciona porque ambas a variaveis sao unicode             
                            print "  " +tituloProposta
                            print "  " +telefone
                            #print "  " +soup3.h1.text.replace("\n","")
                            #print "  " +valorPromocional
                            #print "  " +valorReal
                            #print "  " +percentagem
                            lista.write(count,0,count,colStyle)
                            lista.write(count,1,tituloProposta,colStyle)
                            lista.write(count,3,'-',colStyle)
                            lista.write(count,2,telefone,colStyle)
                            lista.write(count,4,soup3.h1.text.replace("\n",""),colStyle)
                            lista.write(count,5,valorPromocional,colStyle)
                            lista.write(count,6,valorReal,colStyle)
                            lista.write(count,7,percentagem,colStyle)
                            telList = telList + [telefone]
                            count= 1+count
                    nextsite = soup2.find('a', {"rel":"next"})
                    if(nextsite != None):
                        url = 'https://www.groupon.com.br'+nextsite['href']
                        req2 = urllib2.Request(url)
                        response2 = urllib2.urlopen(req2)
                        the_page2 = response2.read()
                        soup2 = BeautifulSoup(the_page2)
                    while (nextsite != None):
                        print 'https://www.groupon.com.br'+nextsite['href']
                        #if(countpage != 1): # nao conta a primeira leitura uma vez que ja foi vista a pagina antes do while
                        for proposta in soup2.findAll('figure',{"class":"deal-card deal-list-tile deal-tile deal-tile-standard"}):   
                            print proposta.a['href'].replace("//","https://")
                            linkProposta= proposta.a['href'].replace("//","https://")
                            #try:
                            req3 = urllib2.Request(linkProposta)
                            response3 = urllib2.urlopen(req3)
                            the_page3 = response3.read()
                            soup3 = BeautifulSoup(the_page3)
                            pattern = re.compile(u'<\/?\w+\s*[^>]*?\/?>', re.DOTALL | re.MULTILINE | re.IGNORECASE | re.UNICODE)
                            tel =  re.findall(r'\(\d{2}\) *\d{4}\-?\d{4}|\d{2} \d{4}\-?\d{4}' , soup3.text)
                            print tel
                            telefone = pattern.sub(u" ", tel[0]).replace(" ","").replace(")",") ")
                            valorPromocional= proposta.find('s',{"class":"discount-price"}).text
                            val1= re.search('\d*\.?\d+(\,\d{2})?',valorPromocional ).group(0).replace('.','').replace(',','.')
                            tituloProposta=soup3.h5.text
                            if(proposta.find('s',{"class":"original-price"}).text is not None and proposta.find('s',{"class":"original-price"}).text != ''):
                                valorReal= proposta.find('s',{"class":"original-price"}).text
                                val2= re.search('\d*\.?\d+(\,\d{2})?',valorReal ).group(0).replace('.','').replace(',','.')
                                percentagem = str(float((float(val1)/float(val2))*100))
                            else:
                                percentagem= '-'
                                valorReal='-'
                            if(telefone not in telList): #funciona porque ambas a variaveis sao unicode             
                                print "  " +tituloProposta
                                print "  " +telefone
                                #print "  " +soup3.h1.text.replace("\n","")
                                #print "  " +valorPromocional
                                #print "  " +valorReal
                                #print "  " +percentagem
                                lista.write(count,0,count,colStyle)
                                lista.write(count,1,tituloProposta,colStyle)
                                lista.write(count,3,'-',colStyle)
                                lista.write(count,2,telefone,colStyle)
                                lista.write(count,4,soup3.h1.text.replace("\n",""),colStyle)
                                lista.write(count,5,valorPromocional,colStyle)
                                lista.write(count,6,valorReal,colStyle)
                                lista.write(count,7,percentagem,colStyle)
                                telList = telList + [telefone]
                               
                                count= 1+count
                            #except urllib2.HTTPError as err:
                            #    print 'erro consulta de ' + listaSegmentos[j] #avisa o erro
                                #winsound.Beep(Freq,3000)
                            #except: 
                            #    print 'erro de resposta do servidor'
                                #winsound.Beep(Freq,3000)
                        nextsite = soup2.find('a', {"rel":"next"})
                        if nextsite != None:
                            newurl ='https://www.groupon.com.br'+nextsite['href']
                            #countpage = countpage + 1
                            try:
                                req = urllib2.Request(newurl)
                                response = urllib2.urlopen(req)
                                the_page = response.read()
                                soup = BeautifulSoup(the_page)
                            except urllib2.HTTPError as err:
                                print 'erro: voce pode ficar com menos informacao' #avisa o erro
                                nextsite = None
                except urllib2.HTTPError as err:
                            print 'erro: ao conectar ao link da proxima cidade sera passada para outra cidade' #avisa o erro
                            nextsite = None
                book.save(listaProdSegmentos[j]+'.xls') #salva o arquivo
        book = Workbook()
        lista = book.add_sheet('lista')
        lista.write(0,0,'Indice',titleStyle)
        lista.write(0,1,'Empresa',titleStyle)
        lista.write(0,2,'Telefone',titleStyle)
        lista.write(0,3,'Cidade',titleStyle)
        lista.write(0,4,'Promocao',titleStyle)
        lista.write(0,5,'Preco',titleStyle)
        lista.write(0,7,'OFF em %',titleStyle)
        lista.write(0,6,'Preco original',titleStyle)
        lista.col(0).width = 2000
        lista.col(1).width = 15000
        lista.col(2).width = 10000
        lista.col(3).width = 5000
        lista.col(4).width = 30000
        lista.col(5).width = 5000
        lista.col(6).width = 5000
        lista.col(7).width = 5000
        count = 1
        telList = []
        for j in range(0,len(listaProdSegmentos)):
            bookMunicipios = open_workbook(listaProdSegmentos[j] +'.xls') # acessa a planilha com os municipios importantes
            sheet = bookMunicipios.sheet_by_index(0)
            for k in range(1,sheet.nrows):
                if (sheet.cell(i,2).value not in telList):
                    lista.write(count,0,count,colStyle)
                    lista.write(count,1, sheet.cell(k,1).value)
                    lista.write(count,3, sheet.cell(k,3).value,colStyle)
                    lista.write(count,2,sheet.cell(k,2).value,colStyle)
                    lista.write(count,5, sheet.cell(k,5).value)
                    lista.write(count,7, sheet.cell(k,6).value,colStyle)
                    lista.write(count,4,sheet.cell(k,4).value,colStyle)
                    telList = telList + [sheet.cell(k,2).value]
                    count = count + 1
    book.save('Produtos ' + 'Geral.xls')
from pymongo import MongoClient
from pymongo.server_api import ServerApi
import tkinter as tk
from tkinter import *
from tkinter import filedialog,ttk, messagebox
import pandas as pds
from geopy.geocoders import Bing,ArcGIS
import time
from datetime import datetime
import viacep
from pycep_correios import get_address_from_cep, WebService
from math import sin, cos, sqrt, radians, asin
import pybingmaps
import sys
import os
import arcgis as AG
import simplekml
import scipy
import time
import pyautogui
import win32clipboard as clip
import win32con
from io import BytesIO

#TESTE R. Antônio Drumond, 350 - Monte Castelo, Fortaleza - CE, 60325-700
# -3.728608, -38.548797

class Application:
    def __init__(self, master):
        self.widget1 = Frame(master,bg='black')
        self.widget1.pack()
        self.widget1["pady"] = 10
        vcmd = (master.register(self.validate))      

        self.variable = StringVar(master)
        self.variable.set("Selecione") 

        self.variable_1 = StringVar(master)
        self.variable_1.set("ÚNICA")

        self.variable_2 = StringVar(master)
        self.variable_2.set("Selecione")

        self.variable_3 = StringVar(master)
        self.variable_3.set("Selecione")
        
        self.msg = Label(self.widget1, text="Viability B2B", bg = 'black', fg = 'white')
        self.msg["font"] = ("Verdana", "10", "italic", "bold")
        self.msg.pack ()

        self.form_1 = Frame(master,bg = 'black')
        self.form_1.pack()
        self.form_1.columnconfigure(0,weight=1)
        self.form_1.columnconfigure(1,weight=3)
        
        self.segundoContainer = Frame(master,bg='black')
        self.segundoContainer["pady"] = 10
        self.segundoContainer.pack()
        self.segundoContainer.columnconfigure(0,weight=1)
        self.segundoContainer.columnconfigure(1,weight=1)
        
        self.quintoContainer = Frame(master,bg='black')
        self.quintoContainer["pady"] = 30
        self.quintoContainer.pack()

        self.campo_1 = Label(self.form_1,text='Razão Social:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=0,sticky=tk.W,pady=2)
        self.campo_1_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        self.campo_1_entry.grid(column=1,row=0,sticky=tk.W,columnspan=2)
        self.campo_1_entry.focus()

        self.campo_2 = Label(self.form_1,text='Nome Fantasia*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=1,sticky=tk.W,pady=2)
        self.campo_2_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        self.campo_2_entry.grid(column=1, row=1,sticky=tk.W,columnspan=2)
        self.campo_2_entry.focus()
        
        self.campo_3 = Label(self.form_1,text='CNPJ:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=2,sticky=tk.W,pady=2)
        self.campo_3_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'), validate = 'key', validatecommand = (vcmd,'%S'))
        self.campo_3_entry.grid(column=1, row=2,sticky=tk.W,columnspan=2)
        self.campo_3_entry.focus()

        self.campo_4 = Label(self.form_1,text='Logradouro*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=3,sticky=tk.W,pady=2)
        self.campo_4_entry = Entry(self.form_1, width=23,font=('helvetica', 10, 'bold'))
        self.campo_4_entry.grid(column=1, row=3,sticky=tk.W)
        self.campo_4_entry.focus()
        
        self.campo_4_num = Label(self.form_1,text='Número*:             ',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=2, row=3,sticky=tk.W,pady=2)
        self.campo_4_entry_num = Entry(self.form_1, width=7,font=('helvetica', 10, 'bold'))
        self.campo_4_entry_num.grid(column=2, row=3,sticky=tk.E)
        self.campo_4_entry_num.focus()

        self.campo_4_1 = Label(self.form_1,text='CEP*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=4,sticky=tk.W,pady=2)
        self.campo_4_1_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'), validate = 'key', validatecommand = (vcmd,'%S'))
        self.campo_4_1_entry.grid(column=1, row=4,sticky=tk.W,columnspan=2)
        self.campo_4_1_entry.focus()
        
        self.campo_5 = Label(self.form_1,text='Cidade*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=5,sticky=tk.W,pady=2)
        #self.campo_5_entry = OptionMenu(self.form_1, self.variable_3, 'AAU - Acarau','ACR - Acopiara','AEZ - Arneiroz','AIB - Aiuaba','AJU - Aracaju','ANE - Antonina Do Norte','AQZ - Aquiraz','ARP - Araripe','ASE - Assare','ATS - Altos','BBB - Beberibe','BBH - Barbalha','BJS - Brejo Santo','BQS - Barra Dos Coqueiros','CAA - Carpina','CCM - Camocim','CED - Cedro','CEJ - Cedro de Sao Joao','CIW - Carire','CMS - Campos Sales','CNB - Carnaiba','CTO - Crato','CTS - Crateus','CUC - Caucaia','CVL - Cascavel','CXS - Caxias','CYN - Catarina','CYS - Carius','CZW - Cruz','DNP - Divina Pastora','ESB - Eusebio','FBO - Farias Brito','FHH - Frecheirinha','FLA - Fortaleza','FLS - Flores','FORT - Fortim','GCA - Graca','GNJ - Granja','GOI - Goiana','HZT - Horizonte','IAGA - Itaitinga','IAU - Iguatu','ICO - Ico','IDA - Independencia','IJD - Itaporanga DAjuda','IMC - Ilha De Itamaraca','IMW - Itarema','INX - Ibiapina','IOP - Itapipoca','IPJ - Ipojuca','ISS - Igarassu','ITS - Itapissuma','ITZ - Imperatriz','JAS - Jucas','JAT - Jati','JIJO - Jijoca De Jericoacoara','JNE - Juazeiro Do Norte','JPO - Japoata','LAT - Lagarto','LIO - Limoeiro','LNJ - Laranjeiras','LNT - Limoeiro Do Norte','LVM - Lavras Da Mangabeira','MAH - Missao Velha','MBC - Mombaca','MCW - Maracanau','MDB - Mirandiba','MPA - Amapá','MRT - Mauriti','MUE - Maranguape','MUM - Mucambo','MVA - Morada Nova','NRO - Nossa Senhora Do Socorro','NZM - Nazare Da Mata','OLD - Olinda','OOS - Oros','PBU - Parambu','PDJ - Pindoretama','PIM - Parnamirim','PJS - Pacajus','PKT - Pacatuba','PLH - Paudalho','PNA - Parnaiba','POE - Pentecoste','PPI - Propria','PQT - Piquet Carneiro','PTX - Porteiras','PUB - Pacatuba','PUI - Paulista','PUJ - Pacuja','PUP - Parauapebas','PUU - Paracuru','PWB - Paraipaba','QXA - Quixada','QXO - Quixelo','REE - Rosario Do Catete','RUS - Russas','SBN - Sao Benedito','SCV - Sao Cristovao','SGA - Sao Goncalo Do Amarante','SGI - Salgueiro','SHD - Serra Talhada','SIQ - Salitre','SLS - Sao Luis','SOL - Sobral','SRU - Sao Luis Do Curu','SUU - Surubim','SYR - Siriri','SZC - Santa Cruz Do Capibaribe','TEV - Terra Nova','TIG - Tiangua','TIU - Timbauba','TLH - Telha','TMN - Timon','TNT - Tabuleiro Do Norte','TRR - Tarrafas','TRY - Trairi','TSA - Teresina','TTA - Toritama','TUA - Taua','UBJ - Ubajara','VJE - Verdejante','VZG - Varzea Alegre')
        self.campo_5_entry = Entry(self.form_1, width=23,font=('helvetica', 10, 'bold'))
        self.campo_5_entry.grid(column=1, row=5,sticky=tk.W,columnspan=2)
        self.campo_5_entry.focus()
        
        self.campo_6 = Label(self.form_1,text='UF*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=2, row=5,sticky=tk.W,pady=2)
        self.campo_6_entry = Entry(self.form_1, width=10,font=('helvetica', 10, 'bold'))
        self.campo_6_entry.grid(column=2, row=5,sticky=tk.E)
        self.campo_6_entry.focus()
        
        self.campo_7 = Label(self.form_1,text='Coordenadas:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=7,sticky=tk.W,pady=2)
        self.campo_7_entry = Entry(self.form_1, width=19,font=('helvetica', 10, 'bold'))
        self.campo_7_entry.grid(column=1, row=7,sticky=tk.W)
        self.campo_7_entry.focus()
        self.init_placeholder(self.campo_7_entry,"Latitude")

        self.campo_7_entry_1 = Entry(self.form_1, width=19,font=('helvetica', 10, 'bold'))
        self.campo_7_entry_1.grid(column=1, row=7,sticky=tk.E,columnspan=2)
        self.campo_7_entry_1.focus()
        self.init_placeholder(self.campo_7_entry_1,"Longitude")
        
        self.campo_8 = Label(self.form_1,text='Tipo de Serviço*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=8,sticky=tk.W,pady=2)
        self.campo_8_entry = OptionMenu(self.form_1, self.variable, "SEMI-DEDICADO", "IP DEDICADO")
        self.campo_8_entry.grid(column=1, row=8,sticky=tk.W,pady=2,columnspan=2)
        self.campo_8_entry.focus()
        
        self.campo_9 = Label(self.form_1,text='IP Público:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=9,sticky=tk.W,pady=2)
        #self.campo_9_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        self.campo_9_entry = OptionMenu(self.form_1, self.variable_2, "SIM", "NÃO")
        self.campo_9_entry.grid(column=1, row=9,sticky=tk.W,pady=2,columnspan=2)
        self.campo_9_entry.focus()
        
        self.campo_10 = Label(self.form_1,text='Tipo de Acesso:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=10,sticky=tk.W,pady=2)
        self.campo_10_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        self.campo_10_entry.grid(column=1, row=10,sticky=tk.W,columnspan=2)
        self.campo_10_entry.focus()
        
        self.campo_11 = Label(self.form_1,text='Banda*:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=11,sticky=tk.W,pady=2)
        self.campo_11_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'), validate = 'key', validatecommand = (vcmd,'%S'))
        self.campo_11_entry.grid(column=1, row=11,sticky=tk.W,columnspan=2)
        self.campo_11_entry.focus()
        
        #self.campo_12 = Label(self.form_1,text='PTT:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=12,sticky=tk.W,pady=2)
        #self.campo_12_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        #self.campo_12_entry.grid(column=1, row=12,sticky=tk.W,columnspan=2)
        #self.campo_12_entry.focus()
        
        self.campo_13 = Label(self.form_1,text='VLans (qtd.):',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=13,sticky=tk.W,pady=2)
        self.campo_13_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'), validate = 'key', validatecommand = (vcmd,'%S'))
        self.campo_13_entry.grid(column=1, row=13,sticky=tk.W,columnspan=2)
        self.campo_13_entry.focus()
        
        #self.campo_14 = Label(self.form_1,text='Tem ASN?:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=14,sticky=tk.W,pady=2)
        #self.campo_14_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        #self.campo_14_entry.grid(column=1, row=14,sticky=tk.W,columnspan=2)
        #self.campo_14_entry.focus()
        
        #self.campo_15 = Label(self.form_1,text='Contato:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=15,sticky=tk.W,pady=2)
        #self.campo_15_entry = Entry(self.form_1, width=40,font=('helvetica', 10, 'bold'))
        #self.campo_15_entry.grid(column=1, row=15,sticky=tk.W,columnspan=2)
        #self.campo_15_entry.focus()

        #self.drop = Label(self.form_1, text='Tipo de plano*: ', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=16,sticky=tk.W,pady=2)
        #self.w = OptionMenu(self.form_1, self.variable, "BANDA LARGA", "PONTO A PONTO - S/GPON", "PONTO A PONTO - C/GPON")
        #self.w.grid(column=1, row=16,sticky=tk.W,pady=2,columnspan=2)
       
        self.drop_1 = Label(self.form_1, text='Viabilidade*: ', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=17,sticky=tk.W,pady=2)
        self.w_1 = OptionMenu(self.form_1, self.variable_1, "ÚNICA", "MASSIVA")
        self.w_1.grid(column=1, row=17,sticky=tk.W,pady=2,columnspan=2)
        
        self.consultar = Button(self.segundoContainer,text='Executar', width=10, relief=GROOVE,font=('helvetica', 10, 'bold'))
        self.consultar["command"] = self.getPRE
        self.consultar.grid(column=0, row=0,sticky=tk.W,pady=2, padx=3)

        self.copy = Button(self.segundoContainer, command= lambda : self.cliente() ,text='Copiar', width=10, relief=GROOVE,font=('helvetica', 10, 'bold'))
        self.copy.grid(column=1, row=0,sticky=tk.W,pady=2, padx=3)

        self.template = Button(self.segundoContainer, command= lambda : self.baixar_templ() ,text='Template', width=10, relief=GROOVE,font=('helvetica', 10, 'bold'))
        self.template.grid(column=2, row=0,sticky=tk.W,pady=2, padx=3)
        
        self.mensagem = Label(self.quintoContainer, text="", bg = 'black', fg = 'white', font=('helvetica', 7, 'bold'))
        self.mensagem.pack()
        self.mensagem_1 = Label(self.quintoContainer, text="", bg = 'black', fg = 'white', font=('helvetica', 10, 'bold'))
        self.mensagem_1.pack()

        self.version = version
        
        self.assinatura = Label(master, text="Desenvolvido por Sergio Tavora                                                                            " + self.version,bg='black', fg='white', font=('helvetica', 7, 'bold')).place(x=5,y=530)
        
    def getCoor (self,df1):
        
        api_1 = "AgVOdIegqF8C4XE0d4nxcSoGrHXvguPwdPky3AzwCnGECAAK6_c4M2H83lJtgHip"
        tempo = datetime.now()
        nome = 'MOBWIRE_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S'
        
        #geolocator = Bing(api_key = api_1)
        geolocator = AG.ArcGIS(user_agent = nome)
        df3 = pds.DataFrame()

        df2 = pds.DataFrame(df1)
        linhas = len(df2.index)
        lat = []
        lon = []
        j = 0
        for i in range(0,linhas):

            logradouro = str(df2.loc[i,'LOGRADOURO']) + " " + str(df2.loc[i,'NÚMERO'])
            cidade = str(df2.loc[i,'MUNICÍPIO'])
            estado = str(df2.loc[i,'UF'])
            cep = str(df2.loc[i,'CEP'])
            
            location = geolocator.geocode(log = logradouro, city = cidade,uf = estado, postal = cep, timeout=5)
            #print(location)
            print('Consultando Coordenadas: ',i*100/linhas,'%. ')
            self.mensagem['text'] = 'Consultando Coordenadas: ' + str(i*100/linhas) + '%. '
            root.update()
            try:
                lat.append(location.latitude)
                lon.append(location.longitude)
            except:
                lat.append('')
                lon.append('')
            if j == 1000: #A cada X consultas eh feita o backup
                df3['LAT'] = lat
                df3['LON'] = lon
                nome = 'BKP_COORDENADAS.xlsx'
                try:
                    df3.to_excel(nome)
                except:
                    df3.to_csv(nome)
                df3.drop(df3.index, inplace=True)
                j = 0
            j += 1
            time.sleep(0.5)
    
    
        df2['LAT'] = lat
        df2['LON'] = lon
        
        tempo = datetime.now()
        nome = 'resultado_coordenada_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S.xlsx'

        #df2.to_excel(nome)

        #self.mensagem['text'] = 'Concluido!'
        #concluido(self)
        root.update()

        return df2
    
    def getEnde (self):
        global df1
    
        import_file_path = filedialog.askopenfilename()
        df1 = pds.read_excel (import_file_path)
        
        api_1 = str(self.api.get())
        
        #geo = Bing(api_key = api_1)
        geo = ArcGIS(user_agent = 'mob')
        df3 = pds.DataFrame()
        
        linhas = len(df1.index)
        lista = []
        lista_1 = []
        lista_2 = []
        lista_3 = []
        lista_4 = []
        lista_5 = []
        lista_6 = []
        lista_7 = []
        lista_8 = []
        j = 0
        for i in range(0,linhas):
            la1 = str(df1.loc[i,'LATITUDE']).replace(',','.')
            lo1 = str(df1.loc[i,'LONGITUDE']).replace(',','.')
            coord = [la1,lo1]
            try:
                consulta = geo.reverse(coord, timeout=2)
                consulta_1 = consulta.raw['LongLabel']
            #consulta = pds.DataFrame(consulta.raw['address'])
            except:
                consulta = ""
                
            try:
                sep = str(consulta.raw['LongLabel']).split(" ")
                tipo_log = sep[0]
            except:
                tipo_log = ""    
                
            try:
                sep_1 = str(consulta.raw['LongLabel']).split(",")
                log = sep_1[0]
                log_comp = sep_1[0]
                sep_1 = log.split(' ')
                log = ' '.join(sep_1[1:])
                log_comp = log_comp.replace(' ', '+')
                #print(log_comp)
            except:
                log = ""
                
            try:
                cidade = str(consulta.raw['City'])
                #print(cidade)
            except:
                cidade = ""
                    
            try:
                uf = str(consulta.raw['Region'])
            except:
                uf = ""
                    
            try:
                cep = str(consulta.raw['Postal'])
            except:
                cep = ""
                
            try:
                pesquisa = str(cep).replace('-','-')
                #pesquisa = 'CE/Acopiara/Pedro+Alves'
                #print(pesquisa)
                #d = viacep.ViaCEP(pesquisa)
                #correios = d.getDadosCEP()
                correios = get_address_from_cep(pesquisa, webservice=WebService.CORREIOS)
                bairro = correios['bairro']
                #print(bairro)
                try:
                    ibge = correios['ibge']
                    #print(ibge)
                except:
                    ibge = ""    
                try:
                    cep = correios['cep']
                    #print(ibge)
                except:
                    cep = ""
            except:
                #print('Entrou')
                try:
                    pesquisa = str(df1.loc[i,'ESTADO']) + '/' + str(cidade) + '/' + str(log_comp)
                    #pesquisa = 'CE/Acopiara/Pedro+Alves'
                    #print(pesquisa)
                    d = viacep.ViaCEP(pesquisa)
                    correios = pds.DataFrame(d.getDadosCEP())
                    bairro = correios.loc[0,'bairro']
                    #print(bairro)
                except:
                    bairro = ""
                try:
                    ibge = correios.loc[0,'ibge']
                    #print(ibge)
                except:
                    ibge = ""
                try:
                    cep = correios.loc[0,'cep']
                    #print(ibge)
                except:
                    cep = ""
                
            print('Consultando Enderecos: ',i*100/linhas,'%. ')
            self.mensagem['text'] = 'Consultando Enderecos: ' + str(i*100/linhas) + '%. '
            root.update()
            
            lista.append(consulta_1)
            lista_1.append(tipo_log)
            lista_2.append(log)
            lista_3.append(cidade)
            lista_4.append(uf)
            lista_5.append(cep)
            lista_6.append(bairro)
            lista_7.append(ibge)
            lista_8.append(cep)
            
            if j == 1000: #A cada X consultas eh feita o backup
                df3['ENDE'] = lista
                df3['TipoLogradouro']=lista_1
                df3['LogradouroEndereco']=lista_2
                df3['NomeCidade']=lista_3
                df3['UFCidade']=lista_4
                df3['CepEndereco']=lista_8
                df3['NomeBairro']=lista_6
                df3['CodigoCidadeIBGE']=lista_7
                nome = 'BKP.xlsx'
                try:
                    df3.to_excel(nome)
                except:
                    df3.to_csv(nome)
                df3.drop(df3.index, inplace=True)
                j = 0
            j += 1
            time.sleep(0.5)
        

        df1['ENDE']=lista
        df1['TipoLogradouro']=lista_1
        df1['LogradouroEndereco']=lista_2
        df1['NomeCidade']=lista_3
        df1['UFCidade']=lista_4
        df1['CepEndereco']=lista_8
        df1['NomeBairro']=lista_6
        df1['CodigoCidadeIBGE']=lista_7
        print(df1['ENDE'])
        #print(df1.head)
        tempo = datetime.now()
        nome = 'resultado_endereco_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S.xlsx'
        
        df1.to_excel(nome)
    
        self.mensagem['text'] = 'Concluido!'
        self.concluido
        root.update()
    
    def getPRE(self):

        self.mensagem['text'] = 'Estabelecendo conxão com o servidor...'
        root.update()

        try:

            self.client = MongoClient("mongodb+srv://user_test:novasenha@cluster0.ihxlfex.mongodb.net/?retryWrites=true&w=majority", server_api=ServerApi('1'))

            print('Conectado!')
            self.mensagem['text'] = 'Conectado!'
            root.update()

        except:

            print('Falha na conexão!')
            self.mensagem['text'] = 'Falha na conexão! Favor  verifique sua conexão com a internet e tente novamente!'
            root.update()
            return

        star = time.time()

        if str(self.campo_6_entry.get()) == 'DATABASE':

            self.database(self.client)

        '''if str(self.campo_6_entry.get()) != 'BYPASS':
            check = True

            if self.variable_1.get() == "ÚNICA":
                check = self.coor_check()

            if check == False:
                return print("Clique de novo!")'''

        #print(str(self.campo_7_entry.get()))
        
        while ((str(self.variable.get()) == "Selecione" or str(self.campo_6_entry.get()) == "" or str(self.campo_5_entry.get()) == "" or str(self.campo_4_1_entry.get()) == "" or str(self.campo_4_entry_num.get()) == "" or str(self.campo_4_entry.get()) == "" or str(self.campo_11_entry.get()) == "" or str(self.campo_2_entry.get()) == "") and str(self.variable_1.get()) == "ÚNICA") or (str(self.variable_1.get()) == "MASSIVA" and (str(self.variable.get()) == "Selecione" or str(self.campo_11_entry.get()) == "")):

            self.mensagem['text'] = 'Preencha todos os campos obrigatórios!'
            root.update()
            return print("Clique de novo!")

        if self.variable_1.get() == "ÚNICA":

            if str(self.campo_7_entry.get()) != "" and str(self.campo_7_entry_1.get()) != "" and str(self.campo_7_entry.get()) != "Latitude" and str(self.campo_7_entry_1.get()) != "Longitude":
                
                data = pds.DataFrame()
                data["LOCAL"] = [self.campo_2_entry.get()]
                data["LAT"] = [self.campo_7_entry.get()]
                data["LON"] = [self.campo_7_entry_1.get()]
                print(data)
            
            elif self.campo_4_entry.get() != "" and self.campo_4_1_entry.get() != "" and self.campo_4_entry_num.get() != "" and self.campo_2_entry.get() != "Selecione" and self.campo_6_entry.get() != "" and self.campo_5_entry.get() != "":

                data_1 = pds.DataFrame()
                data_1["LOCAL"] = [self.campo_2_entry.get()]
                data_1["LOGRADOURO"] = [self.campo_4_entry.get()]
                data_1["NÚMERO"] = [self.campo_4_entry_num.get()]
                data_1["CEP"] = [self.campo_4_1_entry.get()]
                data_1["MUNICÍPIO"] = [self.campo_5_entry.get()]
                data_1["UF"] = [self.campo_6_entry.get()]

                data = self.getCoor(data_1)
                print(data)

            else:

                self.mensagem['text'] = 'Favor fornecer informações de localização do cliente válidos!'
                root.update()

        else:
            
            self.mensagem['text'] = 'Selecionar Planilha de Clientes'
            root.update()
            pl1 = filedialog.askopenfilename()
            pl1_1 = pds.read_excel(pl1)
            data = pds.DataFrame(pl1_1)

            #AQUI VERIFICA A EXISTENCIA DOS CAMPOS DE COORDENDAS E ENDEREÇAMENTO DA PLANILHA CARREGADA

            if str(data.loc[0,'LAT']) != "nan" and str(data.loc[0,'LON']) != "nan":

                print(str(data.loc[0,'LAT']))
                print(str(data.loc[0,'LON']))
                #print("ENTROU ERRADO!")
                data = data
                
            elif str(data.loc[0,'LOGRADOURO']) != "" and str(data.loc[0,'NÚMERO']) != "" and str(data.loc[0,'CEP']) != "" and str(data.loc[0,'MUNICÍPIO']) != "" and str(data.loc[0,'UF']) != "":

                #print("ENTROU!")
                data = self.getCoor(data)

            else:
                self.mensagem['text'] = 'Favor fornecer uma planilha com informações de localização do cliente válidos!'
                root.update()
        
        if self.variable.get() == "SEMI-DEDICADO":

            #caixas_1 = pds.read_excel('CTO.xlsx')
            caixas_1 = pds.read_excel(self.resource_path('CTO.xlsx'))
            caixas = pds.DataFrame(caixas_1)

            '''if self.variable_3.get() != "Selecione":
                
                city = str(self.variable_3.get()).split()[0]
                caixas = caixas[caixas['CIDADE'] == city]
                caixas = caixas.reset_index(drop=True)'''
            
            caixas_CT = caixas
            caixas_CT = caixas_CT.rename(columns={'LATITUDE':'LAT'})
            caixas_CT = caixas_CT.rename(columns={'LONGITUDE':'LON'})

            self.coor_CT = pds.DataFrame(caixas_CT, columns=['LAT','LON'])
            list_coor_CT = self.coor_CT.to_numpy()

            tree_CT = scipy.spatial.cKDTree(list_coor_CT)

        elif self.variable.get() == "IP DEDICADO" and int(self.campo_11_entry.get()) > 400:

            #caixas_1 = pds.read_excel('EMENDA.xlsx')
            caixas_1 = pds.read_excel(self.resource_path('EMENDA.xlsx'),usecols=['NOME_LOCAL','LATITUDE','LONGITUDE'])
            pop_1 = pds.read_excel(self.resource_path('POP.xlsx'),usecols=['NOME_LOCAL','LATITUDE','LONGITUDE'])

            caixas = pds.concat([caixas_1,pop_1],axis='index').reset_index(drop=True)

            caixas_RD = pds.concat([caixas_1,pop_1],axis='index').reset_index(drop=True)
            caixas_RD = caixas_RD.rename(columns={'LATITUDE':'LAT'})
            caixas_RD = caixas_RD.rename(columns={'LONGITUDE':'LON'})

            self.coor_RD = pds.DataFrame(caixas_RD, columns=['LAT','LON'])
            list_coor_RD = self.coor_RD.to_numpy()

            tree_RD = scipy.spatial.cKDTree(list_coor_RD)

        elif self.variable.get() == "IP DEDICADO" and int(self.campo_11_entry.get()) <= 400:

            #caixas_1 = pds.read_excel('POP.xlsx')
            caixas_1 = pds.read_excel(self.resource_path('CTO.xlsx'),usecols=['NOME_LOCAL','LATITUDE','LONGITUDE'])

            '''if self.variable_3.get() != "Selecione":
                
                city = str(self.variable_3.get()).split()[0]
                caixas_1 = caixas_1[caixas_1['CIDADE'] == city]
                caixas_1 = caixas_1.reset_index(drop=True)'''

            emenda_1 = pds.read_excel(self.resource_path('EMENDA.xlsx'),usecols=['NOME_LOCAL','LATITUDE','LONGITUDE'])
            pop_1 = pds.read_excel(self.resource_path('POP.xlsx'),usecols=['NOME_LOCAL','LATITUDE','LONGITUDE'])

            caixas = pds.concat([caixas_1,emenda_1,pop_1],axis='index').reset_index(drop=True)
            caixas = pds.DataFrame(caixas)

            caixas_RD = pds.concat([emenda_1,pop_1],axis='index').reset_index(drop=True)
            caixas_RD = caixas_RD.rename(columns={'LATITUDE':'LAT'})
            caixas_RD = caixas_RD.rename(columns={'LONGITUDE':'LON'})

            caixas_CT = caixas_1
            caixas_CT = caixas_CT.rename(columns={'LATITUDE':'LAT'})
            caixas_CT = caixas_CT.rename(columns={'LONGITUDE':'LON'})

            self.coor_CT = pds.DataFrame(caixas_CT, columns=['LAT','LON'])
            self.coor_RD = pds.DataFrame(caixas_RD, columns=['LAT','LON'])

            list_coor_CT = self.coor_CT.to_numpy()
            list_coor_RD = self.coor_RD.to_numpy()

            tree_CT = scipy.spatial.cKDTree(list_coor_CT)
            tree_RD = scipy.spatial.cKDTree(list_coor_RD)
            #print(caixas)

        d = []
        cx = []
        norm_cx = []

        self.consolidado = caixas
        for i in range(0,len(self.consolidado.index)):
            if str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[0] == 'CTOE':
                if len(str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[1]) == 4:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:16]
                else:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:15]
            elif str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[0] == 'CDOE':
                if len(str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[1]) == 4:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:15]
                else:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:14]
            elif str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[0] == 'CEO':
                if len(str(self.consolidado.loc[i,'NOME_LOCAL']).split('-')[1]) == 4:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:14]
                else:
                    norm = str(self.consolidado.loc[i,'NOME_LOCAL'])[:13]
            else:
                norm = str(self.consolidado.loc[i,'NOME_LOCAL'])

            norm_cx.append(norm)

        self.consolidado['NOME_LOCAL'] = norm_cx
        #tree_all = scipy.spatial.cKDTree(list_coor)

        linhas = len(data.index)
        for i in range(0,linhas):
        
            la1 = str(data.loc[i,'LAT']).replace(',','.')
            lo1 = str(data.loc[i,'LON']).replace(',','.')

            if self.variable.get() == "IP DEDICADO" and int(self.campo_11_entry.get()) <= 400:
            
                dist_CT, index_CT = tree_CT.query([(la1),(lo1)])
                dist_RD, index_RD = tree_RD.query([(la1),(lo1)])

                preco_RD = dist_RD*4.95

                if dist_CT <=800:
                    preco_CT = dist_CT*2
                else:
                    preco_CT = dist_CT*4.95

                if preco_RD <= preco_CT:
                    dist = dist_RD
                    ct = caixas_RD.loc[index_RD,'NOME_LOCAL']
                else:
                    dist = dist_CT
                    ct = caixas_CT.loc[index_CT,'NOME_LOCAL']

            elif self.variable.get() == "IP DEDICADO" and int(self.campo_11_entry.get()) > 400:

                dist_RD, index_RD = tree_RD.query([(la1),(lo1)])

                dist = dist_RD
                ct = caixas_RD.loc[index_RD,'NOME_LOCAL']

            else:

                dist_CT, index_CT = tree_CT.query([(la1),(lo1)])
                
                dist = dist_CT
                ct = caixas_CT.loc[index_CT,'NOME_LOCAL']

            self.mensagem['text'] = 'Etapa 1: ' + str(i*100/linhas) + '%. '
            root.update()

            d.append(dist)
            cx.append(ct)

            #print(i+1)
    
        data['LINEAR'] = d
        data['NXT'] = cx

        #tempo = datetime.now()
        #nome = 'PRE_VIABILIDADE_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S.xlsx'
        #data.to_excel(nome)

        self.mensagem['text'] = 'Etapa 1: Concluido!'
        #self.concluido
        root.update()

        end = time.time()

        #print(end-star)

        self.getVIA(data,caixas)

    def getVIA(self,clientes,caixas):
        
        bing = pybingmaps.Bing("Ah27XoHM8VBbpxG-RaA2ljEPu1SDLVORp6CkkYY-VgXx2OURa96lJNvpSYPOfwpc")
    
        dist = []
        cx = []
        d_min_1 = ""
        dist_1 = []
        m = 0
    
        l = 0
        nxt = 100000000000000
        
        linhas = len(clientes.index)
        for i in range(0,linhas):
            try:
                if clientes.loc[i,"LINEAR"] < 10000:
                    filtro = caixas[caixas['NOME_LOCAL'] == clientes.loc[i,'NXT']]
                    filtro = filtro.reset_index(drop=True)
                
                    la1 = str(clientes.loc[i,'LAT']).replace(',','.')
                    lo1 = str(clientes.loc[i,'LON']).replace(',','.')
                    la2 = str(filtro.loc[0,'LATITUDE']).replace(',','.')
                    lo2 = str(filtro.loc[0,'LONGITUDE']).replace(',','.')
                    #print(filtro)
                
                    start = (float(la1), float(lo1))
                    end = (float(la2), float(lo2))
                    
                    #start = (-3.731074, -38.539697)
                    #end = (-3.732724, -38.539658)
                    #print(start,end)
                    #bing.route(start, end)
                    print(bing.routePathOutput(start, end))
                    l = bing.travelDistance(start, end)
                    d_min_1 = d_min_1 + str(l) + ","
                
                    #print(l)
                
                    if l < nxt:
                        nxt = l
                        caixa = clientes.loc[i,'NXT']
                
                else:
                    nxt = ""
                    caixa = ""
                    d_min_1 = ""


            except:
                nxt = ""
                caixa = ""
                d_min_1 = ""

            self.mensagem['text'] = 'Etapa 2: ' + str((i+1)*100/linhas) + '%. '
            root.update()
            print('Caixa: ' + str(caixa) + '. Distancia ' + str(nxt) + ' metros. ' + 'Viabilidade ' + str(i+1) + ' feita.')

            '''try:
                coef = int(nxt/400)
                cabo = nxt + coef*50 + 50
            except:
                nxt = nxt'''

            dist.append(nxt)
            cx.append(caixa)
            #cpf.append(clientes.loc[i,'ID'])
            dist_1.append(d_min_1[:-1])
            nxt = 100000000000000
            d_min_1 = ""
            
            m += 1
        
        clientes['ROTA'] = dist
        #clientes['NXT_POP_1'] = cx
        #clientes['DISTANCIAS_RAIO_1'] = dist_1
    
        #tempo = datetime.now()
        #nome = 'VIABILIDADE_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S.xlsx'
        #clientes.to_excel(nome)
        #self.porta(clientes,caixas)

        self.mensagem['text'] = 'Etapa 2: Concluido!'
        root.update()

        self.getMateriais(clientes)

    def getMateriais(self,data):

        
        tempo = datetime.now()
        self.nome = 'VIABILIDADE_' + str(tempo.year)+ str(f"{int(tempo.month):02}")+ str(f"{int(tempo.day):02}")+ str(f"{int(tempo.hour):02}")+ str(f"{int(tempo.minute):02}")+ str(f"{int(tempo.second):02}")
        id = str(tempo.year)+ str(f"{int(tempo.month):02}")+ str(f"{int(tempo.day):02}")+ str(f"{int(tempo.hour):02}")+ str(f"{int(tempo.minute):02}")+ str(f"{int(tempo.second):02}")

        lista = []
        lista_hard = []
        lista_as80 = []
        lista_drop = []
        lista_bap = []
        lista_alca = []
        lista_laco = []
        lista_suporte = []
        lista_plaqueta = []
        lista_term = []
        lista_esticad = []
        lista_conec = []
        lista_mikro = []
        lista_conversor = []
        lista_ont_bridge = []
        lista_ont_huawei = []
        lista_rb_750 = []
        lista_rb_760 = []
        lista_rb_2011 = []
        lista_rb_3011 = []
        lista_ccr1009 = []
        lista_caixa = []
        lista_id = []

        linhas = len(data.index)
        for i in range(0,linhas):
            
            try:

                rota = float(data.loc[i,'ROTA'])
                coeficiente = 1

            except:

                rota = 0
                coeficiente = 0

            if str(data.loc[i,'NXT']).split('-')[0] == 'CTOE' and rota <= 800:
                valor = rota*2
                as80 = 0
                try:
                    coef = int(rota/400)
                    drop = rota + coef*50 + 50
                except:
                    drop = rota
                bap = 0
                alca = 0
                laco = 0
                suporte = 0
                plaqueta = 0
                terminador = 1*coeficiente
                esticador = round((rota/35),0)*2
                conector = 2*coeficiente
                mikrotik = 1*coeficiente
            else:
                valor = rota*4.95
                try:
                    coef = int(rota/400)
                    as80 = rota + coef*50 + 50
                except:
                    as80 = rota
                drop = 0
                bap = round(rota/35,0)
                alca = round(bap*0.65*2,0)
                laco = round(bap*0.35,0)
                suporte = alca + laco
                plaqueta = int(((rota/35) + 4)*coeficiente)
                terminador = 1*coeficiente
                esticador = 0
                conector = 0
                mikrotik = 0
            
            lista.append(valor)
            lista_as80.append(as80)
            lista_drop.append(drop)
            lista_bap.append(bap)
            lista_alca.append(alca)
            lista_laco.append(laco)
            lista_suporte.append(suporte)
            lista_plaqueta.append(plaqueta)
            lista_term.append(terminador)
            lista_esticad.append(esticador)
            lista_conec.append(conector)
            lista_mikro.append(mikrotik)

            self.mensagem['text'] = 'Etapa 3: ' + str((i+1)*100/linhas) + '%. '
            root.update()

        data["ORÇAMENTO SERV."] = lista
        data['ORÇAMENTO HARDW.'] = lista
        data["AS80 12F"] = lista_as80
        data["DROP"] = lista_drop
        data["BAP"] = lista_bap
        data["ALÇA"] = lista_alca
        data["LAÇO"] = lista_laco
        data["SUPORTE"] = lista_suporte
        data["PLAQUETA"] = lista_plaqueta
        #data["TERMINADOR"] = lista_term
        data["ESTICADOR"] = lista_esticad
        data["CONECTORES"] = lista_conec
        #data["MIKROTIK"] = lista_mikro

        data.drop("LINEAR",axis=1,inplace=True)

        id_cliente = []
        linhas = len(data.index)
        formulario = []
        result = []
        new_pricing = []
        for i in range(0,linhas):

            try:

                rota = float(data.loc[i,'ROTA'])
                coeficiente = 1

            except:

                rota = 0
                coeficiente = 0

            #id = "Ponto " + str(i+1)
            #id_cliente.append(id)

            custo_prev = 0
            custo_real = 0
            ont_wifi = 0
            ont_bridge = 0
            conversor = 0
            rb_750 = 0
            rb_760 = 0
            rb_2011 = 0
            rb_3011 = 0
            ccr1009 = 0
            caixa = ''

            #NORMALIZA OS NOMES DAS CAIXAS

            if str(data.loc[i,'NXT']).split('-')[0] == 'CTOE':
                if len(str(data.loc[i,'NXT']).split('-')[1]) == 4:
                    caixa = str(data.loc[i,'NXT'])[:16]
                else:
                    caixa = str(data.loc[i,'NXT'])[:15]
            elif str(data.loc[i,'NXT']).split('-')[0] == 'CDOE':
                if len(str(data.loc[i,'NXT']).split('-')[1]) == 4:
                    caixa = str(data.loc[i,'NXT'])[:15]
                else:
                    caixa = str(data.loc[i,'NXT'])[:14]
            elif str(data.loc[i,'NXT']).split('-')[0] == 'CEO':
                if len(str(data.loc[i,'NXT']).split('-')[1]) == 4:
                    caixa = str(data.loc[i,'NXT'])[:14]
                else:
                    caixa = str(data.loc[i,'NXT'])[:13]
            else:
                caixa = str(data.loc[i,'NXT'])

            #print(caixa)

            if (str(data.loc[i,'NXT']).split('-')[0] == 'CTOE' and rota > 1500) or (str(data.loc[i,'NXT']).split('-')[0] != 'CTOE' and rota > 10000) or str(data.loc[i,'ROTA']) == '':
                analise = "INVIÁVEL"
            else:
                analise = "VIÁVEL"

            if self.variable.get() == "SEMI-DEDICADO":

                custo_real = custo_prev + 220
                ont_wifi = 1

                if rota <=800:
                    com = 'ID - ' + id + ' - Via rede Própria\nGPON (SEMI-DEDICADO)\nAtendimento ' + analise + ', via rede própria\n\nNecessário '+ str(rota) +'m de cabo DROP saindo de '+ caixa +'\n\nCusto Serviços: R$'+ str(data.loc[i,'ORÇAMENTO SERV.']) +'\nCusto Hardware: R$' + str(custo_real) + '\n\nCusto total: R$' + str(data.loc[i,'ORÇAMENTO SERV.'] + 220)
                else:
                    com = 'ID - ' + id + ' - Via rede Própria\nGPON (SEMI-DEDICADO)\nAtendimento ' + analise + ', via rede própria\n\nNecessário '+ str(rota) +'m de cabo AS-80 saindo de '+ caixa +'\n\nCusto Serviços: R$'+ str(data.loc[i,'ORÇAMENTO SERV.']) +'\nCusto Hardware: R$' + str(custo_real) + '\n\nCusto total: R$' + str(data.loc[i,'ORÇAMENTO SERV.'] + 220)
                    
            elif self.variable.get() == "IP DEDICADO" and str(data.loc[i,'NXT']).split('-')[0] == 'CTOE':

                if int(self.campo_11_entry.get()) <= 150:
                    
                    if int(self.campo_11_entry.get()) <= 100:
                        modelo = "RB 750"
                        rb = 500
                        rb_750 = 1
                    else:
                        modelo = "RB 760"
                        rb = 500
                        rb_760 = 1

                elif 150 < int(self.campo_11_entry.get()) <= 300:
                    
                    modelo = "RB 2011"
                    rb = 1000
                    rb_2011 = 1

                else:

                    modelo = "RB 3011"
                    rb = 1400
                    rb_3011 = 1
                
                custo_real = custo_prev + 80 + rb
                ont_bridge = 1

                if rota <=800:
                    com = 'ID - ' + id + ' - Via rede Própria\nREDE DEDICADA (IP DEDICADO)\nAtendimento ' + analise + ', via rede própria\n\nNecessário '+ str(rota) +'m de cabo DROP saindo de '+ caixa +'\n\nCusto Serviços: R$'+ str(data.loc[i,'ORÇAMENTO SERV.']) +' \nCusto Hardware: R$' + str(custo_real) + '\n\nCusto total: R$' + str(data.loc[i,'ORÇAMENTO SERV.'] + 80 + rb)
                else:
                    com = 'ID - ' + id + ' - Via rede Própria\nREDE DEDICADA (IP DEDICADO)\nAtendimento ' + analise + ', via rede própria\n\nNecessário '+ str(rota) +'m de cabo AS-80 saindo de '+ caixa +'\n\nCusto Serviços: R$'+ str(data.loc[i,'ORÇAMENTO SERV.']) +' \nCusto Hardware: R$' + str(custo_real) + '\n\nCusto total: R$' + str(data.loc[i,'ORÇAMENTO SERV.'] + 80 + rb)

            else:
                if int(self.campo_11_entry.get()) <= 150:
                    
                    if int(self.campo_11_entry.get()) <= 100:
                        modelo = "RB 750"
                        rb = 500
                        rb_750 = 1
                    else:
                        modelo = "RB 760"
                        rb = 500
                        rb_760 = 1
                elif 150 < int(self.campo_11_entry.get()) <= 300:
                    
                    modelo = "RB 2011"
                    rb = 1000
                    rb_2011 = 1

                else:

                    if int(self.campo_11_entry.get()) <= 500:

                        modelo = "RB 3011"
                        rb = 1400
                        rb_3011 = 1

                    else:

                        modelo = "CCR1009"
                        rb = 4000
                        ccr1009 = 1

                custo_real = custo_prev + 400 + rb
                conversor = 1

                com = 'ID - ' + id + ' - Via rede Própria\nREDE DEDICADA (IP DEDICADO)\nAtendimento ' + analise + ', via rede própria\n\nNecessário '+ str(rota) +'m de cabo AS-80 saindo de '+ caixa +'\n\nCusto Serviços: R$'+ str(data.loc[i,'ORÇAMENTO SERV.']) +' \nCusto Hardware: R$' + str(custo_real) + '\n\nCusto total: R$' + str(data.loc[i,'ORÇAMENTO SERV.'] + 400 + rb)
                
            if analise == "VIÁVEL":
                formulario.append(com)
                result.append(analise)
                lista_ont_huawei.append(ont_wifi)
                lista_ont_bridge.append(ont_bridge)
                lista_conversor. append(conversor)
                new_pricing.append(custo_real)
                lista_rb_750.append(rb_750)
                lista_rb_760.append(rb_760)
                lista_rb_2011.append(rb_2011)
                lista_rb_3011.append(rb_3011)
                lista_ccr1009.append(ccr1009)
                lista_caixa.append(caixa)
                lista_id.append(id)
            else:
                formulario.append(com)
                result.append(analise)
                lista_ont_huawei.append(0)
                lista_ont_bridge.append(0)
                lista_conversor. append(0)
                new_pricing.append("")
                lista_rb_750.append(rb_750)
                lista_rb_760.append(rb_760)
                lista_rb_2011.append(rb_2011)
                lista_rb_3011.append(rb_3011)
                lista_ccr1009.append(ccr1009)
                lista_caixa.append(caixa)
                lista_id.append(id)


        #data['ID_PONTO'] = id_cliente
        data['NXT'] = lista_caixa
        data['ORÇAMENTO HARDW.'] = new_pricing
        data['ONT_WIFI'] = lista_ont_huawei
        data['ONT_BRIDGE'] = lista_ont_bridge
        data['CONVERSOR DE MIDIA'] = lista_conversor
        data['FORMULARIO'] = formulario
        data['RB_750'] = lista_rb_750
        data['RB_760'] = lista_rb_760
        data['RB_2011'] = lista_rb_2011
        data['RB_3011'] = lista_rb_3011
        data['CCR1009'] = lista_ccr1009
        data['ANÁLISE'] = result
        data['ID_VIABILIDADE'] = lista_id

        first_column = data.pop('ID_VIABILIDADE')
        data.insert(0, 'ID_VIABILIDADE', first_column)

        #data.to_excel(nome)
        #f"{int(tempo.second):02}"

        self.mensagem['text'] = 'Etapa 3: Concluido!'
        self.concluido
        root.update()

        try:
            my_collection = self.client['viability']
            my_database = my_collection['viability-b2b']

            entry = {
                    'id_viability':str(id),
                    'cliente':str(self.campo_2_entry.get()),
                    'date':str(tempo),
                    'user':str(os.getlogin()),
                    'tipo':str(self.variable_1.get()),
                    'tx_viabilidade':str(round(((data["ANÁLISE"]=="VIÁVEL").sum())/len(data.index),2)),
                    'app_version':str(self.version)
                    }

            my_database.insert_one(entry)
            self.client.close()

        except:

            self.mensagem['text'] = 'Falha na comunicação do servidor!'
            root.update()

            return

        if self.variable_1.get() == "ÚNICA":
            self.popup_KML(data,id)
        else:
            self.popup_RESULTADO(data,id)

    def concluido(self):
        return messagebox.showinfo('VIABILITY','Concluído!')

    def resource_path(self,relative_path):
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)

    def dist(self,lat1,lon1,lat2,lon2):
        try:
            lat1,lon1,lat2,lon2 = map(radians,[float(lat1),float(lon1),float(lat2),float(lon2)])
            R = 6371

            dlon = lon2 - lon1
            dlat = lat2 - lat1

            a = sin(dlat / 2)**2 + cos(lat1) * cos(lat2) * sin(dlon / 2)**2
            c = 2 * asin(sqrt(a))

            km = R * c
        except:
            km = 1000
        
        return km * 1000

    def popup_KML(self,viabilidade, prim):

        self.popup = Toplevel()
        self.popup.geometry('475x250')
        self.popup.configure(background = 'black')
        self.popup.title("VIABILITY B2B")

        titulo = Label(self.popup, text="RESULTADO",justify='center', bg = 'black', fg = 'white')
        titulo["font"] = ("Verdana", "10", "bold")
        titulo.pack (pady=10)

        conteiner1 = Frame(self.popup,bg='black')
        conteiner1.pack()

        conteiner2 = Frame(self.popup,bg='black')
        conteiner2.pack(pady=2)

        conteinerButton = Frame(self.popup,bg='black')
        conteinerButton.pack(pady=10)
        
        msg = Label(conteiner1, text=str(viabilidade.loc[0,'FORMULARIO']),justify='left', bg = 'black', fg = 'white')
        msg["font"] = ("Verdana", "8", "bold")
        msg.pack()

        data = datetime.now()
        #detail = str(self.campo_2_entry.get()) + '\n' + str(data.day) + '/' + str(data.month) + '/' + str(data.year) + " " + str(data.hour) + 'H' + str(data.minute) + 'MIN' + str(data.second) + 'S'
        detail = prim

        watermark = Label(conteiner2, text=detail,justify='left', bg = 'black', fg = 'white')
        watermark["font"] = ("Verdana", "5", "bold")
        watermark.pack(side = 'left')

        #filtro = viabilidade[viabilidade["ANÁLISE"] == 'VIÁVEL']

        gerar = Button(conteinerButton,text='Gerar KML',command = lambda : [self.KML(viabilidade,prim)], width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar["padx"] = 5
        gerar.pack(side='left',padx=5)

        gerar_1 = Button(conteinerButton,text='Salvar planilha', command = lambda : [viabilidade.to_excel(str(self.nome + ".xlsx"))], width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar_1["padx"] = 5
        gerar_1.pack(side='right',padx=5)

        #gerar_3 = Button(conteinerButton,text='Copiar', command = lambda : self.clipboard(str(viabilidade.loc[0,'FORMULARIO'])), width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar_3 = Button(conteinerButton,text='Print', command = lambda : self.clipboard_IMG(self.popup), width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar_3["padx"] = 5
        gerar_3.pack(side='right',padx=5)

    def popup_RESULTADO(self,viabilidade, prim):

        self.popup = Toplevel()
        self.popup.geometry('275x200')
        self.popup.configure(background = 'black')
        self.popup.title("VIABILITY B2B")
        
        conteiner1 = Frame(self.popup,bg='black')
        conteiner1.pack(pady=2)
        
        conteiner2 = Frame(self.popup,bg='black')
        conteiner2.pack(pady=2)

        titulo = Label(conteiner1, text="RESULTADO",justify='center', bg = 'black', fg = 'white')
        titulo["font"] = ("Verdana", "10", "bold")
        titulo.pack (pady=10)

        texto = "Foi realizada a viabilidade de " + str(len(viabilidade.index)) + " pontos.\n\nPontos viáveis: " + str((viabilidade["ANÁLISE"]=="VIÁVEL").sum()) + " pontos.\nPontos não viáveis: " + str((viabilidade["ANÁLISE"]=="INVIÁVEL").sum()) + " pontos.\n\nTaxa de viabilidade: " + str(round(((viabilidade["ANÁLISE"]=="VIÁVEL").sum())/len(viabilidade.index)*100,2)) + "%"

        msg = Label(conteiner1, text=texto,justify='left', bg = 'black', fg = 'white')
        msg["font"] = ("Verdana", "8", "bold")
        msg.pack (pady=0)

        detail = prim

        watermark = Label(conteiner2, text=detail,justify='left', bg = 'black', fg = 'white')
        watermark["font"] = ("Verdana", "5", "bold")
        watermark.pack()

        conteinerButton = Frame(self.popup,bg='black')
        conteinerButton.pack(pady=10)

        filtro = viabilidade[viabilidade["ANÁLISE"] == 'VIÁVEL']

        gerar = Button(conteinerButton,text='Gerar KML',command = lambda : [self.KML(filtro,prim)], width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar["padx"] = 5
        gerar.pack(side='left',padx=5)

        gerar_1 = Button(conteinerButton,text='Salvar planilha', command = lambda : [viabilidade.to_excel(str(self.nome + ".xlsx"))], width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        #gerar_1 = Button(conteinerButton,text='Salvar planilha', command = lambda : self.clipboard(str(viabilidade.loc[0,'FORMULARIO'])), width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar_1["padx"] = 5
        gerar_1.pack(side='right',padx=5)

    def popup_CHECK(self,text):

        self.popup1 = Toplevel()
        self.popup1.geometry('250x100')
        self.popup1.configure(background = 'black')

        titulo = Label(self.popup1, text="CHECK",justify='center', bg = 'black', fg = 'white')
        titulo["font"] = ("Verdana", "10", "bold")
        titulo.pack (pady=10)

        msg = Label(self.popup1, text=text,justify='left', bg = 'black', fg = 'white')
        msg["font"] = ("Verdana", "8", "bold")
        msg.pack (pady=0)

        conteinerButton = Frame(self.popup1,bg='black')
        conteinerButton.pack(pady=10)

        gerar = Button(conteinerButton,text='OK',command = lambda : [self.popup1.destroy()], width=12, relief=GROOVE,font=('helvetica', 10, 'bold'))
        gerar["padx"] = 5
        gerar.pack(side='left',padx=5)

    def KML(self, data, titulo):

        self.popup.destroy
        print(2)
        
        data = data.merge(self.consolidado,how="left",left_on="NXT",right_on="NOME_LOCAL",suffixes=("","2"))
        print(data)

        nome = 'KML - ' + titulo
        kml = simplekml.Kml(name = nome, open = 1)
        sub = kml.newfolder(name='VIABILIDADES')
        green = sub.newfolder(name='VIABILIDADE 1')
    
        for i in range(0, len(data['LAT'])):

            nome = str(data.loc[i,'LOCAL'])
            KML = green.newpoint(name=nome, coords = [(data.loc[i,'LON'],data.loc[i,'LAT'])] )
            KML.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/red-circle.png'
            KML.style.labelstyle.scale = 0.6
            KML.visibility = 0

            print(str((data.loc[i,'LONGITUDE'],data.loc[i,'LATITUDE'])))
            nome = str(data.loc[i,'NXT'])
            KML = green.newpoint(name=nome, coords = [(data.loc[i,'LONGITUDE'],data.loc[i,'LATITUDE'])] )
            KML.style.iconstyle.icon.href = 'http://maps.google.com/mapfiles/kml/paddle/blu-diamond.png'
            KML.style.labelstyle.scale = 0.6
            KML.visibility = 0

            pasta = 'VIABILIDADE ' + str(i+2)

            if i != len(data['LAT'])-1:
                green = sub.newfolder(name=pasta)
        
            print('CARREGANDO: ',i*100/len(data['LAT']),'% . ',i+1,'pontos gerados')

        #nome_1 = titulo + '.xlsx'
        nome = titulo + '.kml'
        #data.to_excel(nome_1)
        kml.save(nome)

    def clipboard(self,text):

        root.clipboard_clear()
        root.clipboard_append(text)
        root.update()

    def cliente(self):

        form = "Razão Social: " + str(self.campo_1_entry.get()) + "\nNome Fantasia: " + str(self.campo_2_entry.get()) + "\nCNPJ: " + str(self.campo_3_entry.get()) + "\n\nEndereço: " + str(self.campo_4_entry.get()) + " " + str(self.campo_4_entry_num.get()) + "\nCEP: " + str(self.campo_4_1_entry.get()) + "\nCidade: " + str(self.campo_5_entry.get()) + "\nUF: " + str(self.campo_6_entry.get()) + "\nCoordenadas: " + str(self.campo_7_entry.get()) + ", " + str(self.campo_7_entry_1.get()) + "\n\nTipo de Serviço: " + str(self.variable.get()) + "\nNecessário IP Público: " + str(self.variable_2.get()) + "\nTipo de Acesso: " + str(self.campo_10_entry.get()) + "\nBanda: " + str(self.campo_11_entry.get()) + "\nVlans:" + str(self.campo_13_entry.get()) + ""
        self.clipboard(form)

    def validate(self,value_if_allowed):
            if value_if_allowed:
                try:
                    int(value_if_allowed) or str(value_if_allowed) == ""
                    return True
                except ValueError:
                    return False
            else:
                return False

    def coor_check(self):

        api_1 = "AgVOdIegqF8C4XE0d4nxcSoGrHXvguPwdPky3AzwCnGECAAK6_c4M2H83lJtgHip"
        tempo = datetime.now()
        nome = 'MOBWIRE_' + str(tempo.hour) + 'H' + str(tempo.minute) + 'M' + str(tempo.second) + 'S'
        
        #geolocator = Bing(api_key = api_1)
        geolocator = AG.ArcGIS(user_agent = nome)

        logradouro = str(self.campo_4_entry.get()) + " " + str(self.campo_4_entry_num.get())
        cidade = str(self.variable_3.get())
        estado = str(self.campo_6_entry.get())
        cep = str(self.campo_4_1_entry.get())
            
        location = geolocator.geocode(log = logradouro, city = cidade, uf = estado, postal = cep, timeout=5)

        lat1 = self.campo_7_entry.get()
        lon1 = self.campo_7_entry_1.get()
        lat2 = location.latitude
        lon2 = location.longitude

        print(lat2)
        print(lon2)

        distancia = self.dist(lat1,lon1,lat2,lon2)
        print(distancia)

        if distancia > 50:

            self.mensagem['text'] = 'Locais divergentes!'
            root.update()
            self.popup_CHECK('Locais divergentes!')
            return False

        else:

            self.mensagem['text'] = 'Locais checados!'
            self.popup_CHECK('Locais checados!')
            root.update()
            return True

    def clipboard_IMG(self,canvas1):

        x, y = canvas1.winfo_rootx(), canvas1.winfo_rooty()
        w, h = canvas1.winfo_width(), canvas1.winfo_height()

        img = pyautogui.screenshot('screenshot.png', region=(x, y, w, h-45))
        output = BytesIO()
        img.convert('RGB').save(output, 'BMP')
        data = output.getvalue()[14:]
        output.close()
        clip.OpenClipboard()
        clip.EmptyClipboard()
        clip.SetClipboardData(win32con.CF_DIB, data)
        clip.CloseClipboard()

    def remove_placeholder(self,event):
        """Remove placeholder text, if present"""
        placeholder_text = getattr(event.widget, "placeholder", "")
        if placeholder_text and event.widget.get() == placeholder_text:
            event.widget.delete(0, "end")

    def add_placeholder(self,event):
        """Add placeholder text if the widget is empty"""
        placeholder_text = getattr(event.widget, "placeholder", "")
        if placeholder_text and event.widget.get() == "":
            event.widget.insert(0, placeholder_text)

    def init_placeholder(self,widget, placeholder_text):
        widget.placeholder = placeholder_text
        if widget.get() == "":
            widget.insert("end", placeholder_text)
        # set up a binding to remove placeholder text
        widget.bind("<FocusIn>", self.remove_placeholder)
        widget.bind("<FocusOut>", self.add_placeholder)

    def baixar_templ(self):

        template = pds.DataFrame(pds.read_excel(self.resource_path('TEMPLATE.xlsx')))
        
        return template.to_excel("TEMPLATE_MASSIVA.xlsx")

    def database(self,client):

        my_collection = client['viability']
        my_database = my_collection['viability-b2b']

        database = pds.DataFrame(my_database.find({}))
        database.to_excel("DATABASE_VIABILIDADES.xlsx")

class Login:
    def __init__(self, master):
        self.widget1 = Frame(master,bg='black')
        self.widget1.pack()
        self.widget1["pady"] = 10
        
        self.msg = Label(self.widget1, text="LOGIN", bg = 'black', fg = 'white')
        self.msg["font"] = ("Verdana", "10", "italic", "bold")
        self.msg.pack ()

        self.form_1 = Frame(master,bg = 'black')
        self.form_1.pack()
        self.form_1.columnconfigure(0,weight=1)
        self.form_1.columnconfigure(1,weight=3)
        
        self.segundoContainer = Frame(master,bg='black')
        self.segundoContainer["pady"] = 10
        self.segundoContainer.pack()
        self.segundoContainer.columnconfigure(0,weight=1)
        self.segundoContainer.columnconfigure(1,weight=1)
        
        self.quintoContainer = Frame(master,bg='black')
        self.quintoContainer["pady"] = 2
        self.quintoContainer.pack()

        self.campo_1 = Label(self.form_1,text='Login:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=0,sticky=tk.W,pady=2)
        self.campo_1_entry = Entry(self.form_1, width=20,font=('helvetica', 10, 'bold'))
        self.campo_1_entry.grid(column=1,row=0,sticky=tk.W,columnspan=2)
        self.campo_1_entry.focus()

        self.campo_2 = Label(self.form_1,text='Senha:',justify='left', bg = 'black', fg = 'white',font=('helvetica', 10, 'bold')).grid(column=0, row=1,sticky=tk.W,pady=2)
        self.campo_2_entry = Entry(self.form_1, width=20,font=('helvetica', 10, 'bold'),show='*')
        self.campo_2_entry.grid(column=1, row=1,sticky=tk.W,columnspan=2)
        self.campo_2_entry.focus()
        
        self.consultar = Button(self.segundoContainer,text='Acessar',command = lambda : (self.verify(master)), width=10, relief=GROOVE,font=('helvetica', 10, 'bold'))
        #self.consultar["command"] = self.access()
        self.consultar.grid(column=0, row=0,sticky=tk.W,pady=2, padx=3)

        self.mensagem = Label(self.quintoContainer, text="", bg = 'black', fg = 'white', font=('helvetica', 7, 'bold'))
        self.mensagem.pack()

        self.version = version
        
        self.assinatura = Label(master, text="Desenvolvido por Sergio Tavora         " + self.version,bg='black', fg='white', font=('helvetica', 7, 'bold')).place(x=5,y=180)
    
    def access(self, master):

        for child in master.winfo_children():
            child.destroy()

        root.geometry('450x550')
        #root_1.mainloop()
        
        return Application(master)

    def verify(self,master):

        user = str(self.campo_1_entry.get())
        pwd = str(self.campo_2_entry.get())

        while user == '' or pwd == '':
            self.mensagem['text'] = 'Preencha todos os campos de login!'
            root.update()

        self.mensagem['text'] = 'Conectando...'
        root.update()

        client = MongoClient("mongodb+srv://user_test:novasenha@cluster0.ihxlfex.mongodb.net/?retryWrites=true&w=majority", server_api=ServerApi('1'))

        print('CONEXÃO OK')

        my_collection = client['viability']
        my_database = my_collection['users-login']

        dataframe = pds.DataFrame(my_database.find({'user': user, 'password':pwd}))
        print(user)

        if dataframe.loc[0,'status'] != 'ativo':

            self.mensagem['text'] = 'Usuário desativado, favor contatar admin!'
            root.update()

            client.close()

            return print("Tururu")

        if dataframe.empty:

            self.mensagem['text'] = 'Login e/ou Senha incorretos, tente novamente'
            root.update()

            client.close()

            return print("Tururu")

        my_database = my_collection['users-accesslog']

        entry = {
            'user':user,
            'timestamp':str(datetime.now()),
            'app_version':version,
            'localstamp':str(os.getlogin())
            }

        my_database.insert_one(entry)

        client.close()

        return self.access(master)

if __name__ == "__main__":

    global version
    version = 'V1.2 build 001'

    root = tk.Tk()
    root.geometry('250x200')
    #root.attributes("-toolwindow", 1)
    #root.attributes("-topmost", True)
    #root.attributes('-alpha', 0.6)
    root.configure(background = 'black')
    root.title("VIABILITY B2B")
    #Application(root)
    Login(root)
    
    root.mainloop()
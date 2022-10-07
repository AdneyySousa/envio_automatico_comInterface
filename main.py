#importando bibliotecas para uso de janela
import tkinter
import tkinter as tk
from tkinter import ttk, NO, messagebox, END

#importações para banco de dados
from sqlalchemy import create_engine
from sqlalchemy.engine import URL
import pyodbc

#importação pacote envio de emails
import win32com.client as win32

#integração banco dados com python


dados_conexao = ("Driver={BD utilizado};"
                 "Server=endereço servidor;"
                 "Database=nome da tabela pra acesso;")


#fazendo conexao com banco de dados
connection_url = URL.create("mssql+pyodbc", query={"odbc_connect": dados_conexao})
engine = create_engine(connection_url)
conexao = pyodbc.connect(dados_conexao)

cursor = conexao.cursor()


#função para envio dos emails
def enviarEmail():

###################################ENVIO DE EMAILS BRUMADINHO ##############################################33

    #acessando banco de dados para buscar os emails a serem enviados
    consulta_sqlBR = """SELECT * FROM Brumadinho"""
    cursor.execute(consulta_sqlBR)
    linhaBR = cursor.fetchall()

    #repetição criada para que sejam validados o envio dos emails
    for linhasBR in linhaBR:
        textBR= str(linhasBR)
        textFormat = str(textBR.replace(",", "").replace("(", ";").replace(")", "").replace("'",""))
        textoEmailBR = texto_email.get(1.0, "end-1c")
        outlook = win32.Dispatch('outlook.application')
        emailJF = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailJF.To = f"{textFormat}"
        emailJF.Subject = "titulo email"
        # formatado com linguagem html
        emailJF.HTMLBody = f"""                  
                    {textoEmailBR}
        
                      """
        emailJF.Send()


#########################################################FIM ENVIO EMAILS JF############################################

##################################### ENVIO EMAILS UBERLANDIA###############################3
 # acessando banco de dados para buscar os emails a serem enviados
    consulta_sqlUBL = """SELECT * FROM Ubl"""
    cursor.execute(consulta_sqlUBL)
    linhaUBL = cursor.fetchall()

        # convertendo texto do input para inserir no email
    textoEmailUBL = texto_email.get(1.0, "end-1c")
        # repetição criada para que sejam validados o envio dos emails
    for linhasUBL in linhaUBL:
        textUBL = str(linhasUBL)
        textFormatUBL = str(textUBL.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailUBL = outlook.CreateItem(0)

            # configurar as informaçoes do email
        emailUBL.To = f"{textFormatUBL}"
        emailUBL.Subject = "Disponibilidade de horário Novembro"
            # formatado com linguagem html
        emailUBL.HTMLBody = f""" 
                                  <h5>  {textoEmailUBL} </h5>  
        
                                 """

        emailUBL.Send()


############################ FIM ENVIO DE EMAILS UBERLANDIA ##########################################

########################### ENVIO EMAILS DIAMANTINA##################################################
    consulta_sqlDia = """SELECT * FROM Diamantina"""
    cursor.execute(consulta_sqlDia)
    linhaDia = cursor.fetchall()

        # convertendo texto do input para inserir no email
    textoEmailDia = texto_email.get(1.0, "end-1c")
        # repetição criada para que sejam validados o envio dos emails
    for linhasDia in linhaDia:
        textDia = str(linhasDia)
        textFormatDia = str(textDia.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailUBL = outlook.CreateItem(0)

            # configurar as informaçoes do email
        emailUBL.To = f"{textFormatDia}"


        emailUBL.Subject = "titulo email"
            # formatado com linguagem html
        emailUBL.HTMLBody = f""" 
                                  <h5>  {textoEmailDia} </h5>  <br><br>
                                  
                                  
                                
       
                                     """



        emailUBL.Send()


############################ FIM ENVIO EMAILS UBERLANDIA ###################################

########################## ENVIO EMAILS MONTES CLAROS #####################################
    consulta_sqlMoc = """SELECT * FROM Moc"""
    cursor.execute(consulta_sqlMoc)
    linhaMoc = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmailMoc = texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasMoc in linhaMoc:
        textMoc = str(linhasMoc)
        textFormatMoc = str(textMoc.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailMoc = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailMoc.To = f"{textFormatMoc}"

        emailMoc.Subject = "titulo email"
        # formatado com linguagem html
        emailMoc.HTMLBody = f""" 
                                     <h5>  {textoEmailMoc} </h5>  <br><br>
    
    
    
                                        """

        emailMoc.Send()


############################FIM ENVIO EMAIL MONTES CLAROS###########################################3

############################ ENVIO DE EMAIL GOVERNADOR VALADARES#####################################
    consulta_sqlGv = """SELECT * FROM Gv"""
    cursor.execute(consulta_sqlGv)
    linhaGv = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmaiGv = texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasGv in linhaGv:
        textGv = str(linhasGv)
        textFormatGv = str(textGv.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailGv = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailGv.To = f"{textFormatGv}"

        emailGv.Subject = "titulo email"
        # formatado com linguagem html
        emailGv.HTMLBody = f""" 
                                     <h5>  {textoEmaiGv} </h5>  <br><br>
    
    
    
          
                                        """

        emailGv.Send()


####################### FIM ENVIO EMAILS GOVERNADOR VALADARES#######################################

###################### ENVIO EMAILS CONTAGEM#############################################
    consulta_sqlCont = """SELECT * FROM Contagem"""
    cursor.execute(consulta_sqlCont)
    linhaCont = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmaiCont = texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasCont in linhaCont:
        textCont = str(linhasCont)
        textFormatCont = str(textCont.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailCont = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailCont.To = f"{textFormatCont}"

        emailCont.Subject = "titulo email"
        # formatado com linguagem html
        emailCont.HTMLBody = f""" 
                                      <h5>  {textoEmaiCont} </h5>  <br><br>
    
    
    
           
                                         """

        emailCont.Send()


#####################################FIM ENVIO EMAILS CONTAGEM##############################################

#####################################ENVIO EMAILS BRUMADINHO###############################################
    consulta_sqlBru = """SELECT * FROM Brumadinho"""
    cursor.execute(consulta_sqlBru)
    linhaBru = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmaiBru = texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasBru in linhaBru:
        textBru = str(linhasBru)
        textFormatBru = str(textBru.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailBru= outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailBru.To = f"{textFormatBru}"

        emailBru.Subject = "titulo email"
        # formatado com linguagem html
        emailBru.HTMLBody = f""" 
                                       <h5>  {textoEmaiBru} </h5>  <br><br>
    
    
    
             
                                          """

        emailBru.Send()


############################### FIM ENVIO EMAILS BRUMADINHO##################################

############################## ENVIO EMAILS LIBERTAS##########################################
    consulta_sqlLib = """SELECT * FROM Libertas"""
    cursor.execute(consulta_sqlLib)
    linhaLib = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmaiLib= texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasLib in linhaLib:
        textLib= str(linhasLib)
        textFormatLib = str(textLib.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailLib = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailLib.To = f"{textFormatLib}"

        emailLib.Subject = "titulo email"
        # formatado com linguagem html
        emailLib.HTMLBody = f""" 
                                       <h5>  {textoEmaiLib} </h5>  <br><br>
    
    
    
             
                                          """

        emailLib.Send()


################################# FIM EMAIL LIBERTAS#####################################

################################# ENVIO EMAILS DIVINOPOLIS################################
    consulta_sqlDiv = """SELECT * FROM Divinopolis"""
    cursor.execute(consulta_sqlDiv)
    linhaDiv = cursor.fetchall()

    # convertendo texto do input para inserir no email
    textoEmaiDiv = texto_email.get(1.0, "end-1c")
    # repetição criada para que sejam validados o envio dos emails
    for linhasDiv in linhaDiv:
        textDiv = str(linhasDiv)
        textFormatDiv = str(textDiv.replace(",", "").replace("(", ";").replace(")", "").replace("'", ""))

        outlook = win32.Dispatch('outlook.application')

        emailDiv = outlook.CreateItem(0)

        # configurar as informaçoes do email
        emailDiv.To = f"{textFormatDiv}"

        emailDiv.Subject = "Titulo email"
        # formatado com linguagem html
        emailDiv.HTMLBody = f""" 
                                         <h5>  {textoEmaiDiv} </h5>  <br><br>
    
    
    
               
                                            """

        emailDiv.Send()

    messagebox.showinfo(title="Processo finalizado",message='Todos os Emails enviados')



    #funçoes para uso das tabelas
def InserirEmail():

    print(listaCombo.get())
    if email_entry.get() == "" or listaCombo.get() == "":
        messagebox.showwarning(title='erro',message='campo de digitação vazio ou nao foi selecionado a filial')

    #condições para adicionar email nas unidades
    if listaCombo.get() == "Brumadinho":
        tabelaBR.insert("",END,values=(email_entry.get()))

        #inserindo no banco de dados
        comando = f"""INSERT INTO Brumadinho(Brumadinho)
               VALUES   ('{email_entry.get()}')"""

        cursor.execute(comando)
        cursor.commit()


    if listaCombo.get() == "Juiz de Fora":
        tabelaJf.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Jf(Jf)
                      VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()


    if listaCombo.get() == "Governador Valadares":
        tabelaGV.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Gv(Gv)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()


    if listaCombo.get() == "Contagem":
        tabelaCont.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Contagem(Contagem)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    if listaCombo.get() == "Uberlandia":
        tabelaUbl.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Ubl(Ubl)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    if listaCombo.get() == "Montes Claros":
        tabelaMoc.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Moc(Moc)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    if listaCombo.get() == "Libertas":
        tabelaLib.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Libertas(Libertas)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    if listaCombo.get() == "Divinopolis":
        tabelaDiv.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Divinopolis(Divinopolis)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    if listaCombo.get() == "Diamantina":
        tabelaDia.insert("", END, values=(email_entry.get()))

        # inserindo no banco de dados
        comando = f"""INSERT INTO Diamantina(Diamantina)
                              VALUES   ('{email_entry.get()}')"""
        cursor.execute(comando)
        cursor.commit()

    #final das condições para adicionar email


#função para deletar emails Brumadinho
def DelEmailBR():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaBR.selection()[0]

        #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaBR.item(itemSelecionado, "values")
        vl = valores[0].replace(";","")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Brumadinho WHERE Brumadinho = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaBR.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função dell email brumadinho
def DelEmailCont():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaCont.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaCont.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Contagem WHERE Contagem = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()
        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaCont.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função dell email contagem

def DellEmailUbl():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaUbl.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaUbl.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Ubl WHERE Ubl = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaUbl.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email uberlandia

def DellEmailJf():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaJf.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaJf.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Jf WHERE Jf = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaJf.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Juiz de Fora

def DellEmailLib():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaLib.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaLib.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Libertas WHERE Libertas = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaLib.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Libertas

def DellEmailDiv():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaDiv.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaDiv.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Divinopolis WHERE Divinopolis = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaDiv.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Divinopolis

def DellEmailDia():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaDia.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaDia.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Diamantina WHERE Diamantina = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaDia.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Libertas

def DellEmailGV():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaGV.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaGV.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Gv WHERE Gv = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaGV.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Libertas

def DellEmailMoc():
    try:
        #SELECIONA UM ITEM, DENTRO DA TABELA
        itemSelecionado = tabelaMoc.selection()[0]

                #SELECIONA UM VALOR DENTRO DA TABELA
        valores = tabelaMoc.item(itemSelecionado, "values")
        vl = valores[0].replace(";", "")
        print(vl)
        # DELETA O ITEM SELECIONADO NO BANCO DE DADOS
        comando = f""" DELETE FROM Moc WHERE Moc = '{vl}'"""

        cursor.execute(comando)
        cursor.commit()

        #DELETA O ITEM SELECIONADO NA TABELA
        tabelaMoc.delete(itemSelecionado)
    except:
        messagebox.showwarning(title='erro',message='selecione um email para deletar')
#fim função deletar email Libertas

#criando interface grafica
janela = tk.Tk()
janela.title('Disparo de emails') #titulo janela

janela.geometry("{0}x{1}+0+0".format(janela.winfo_screenwidth(), janela.winfo_screenheight()))


#mtexto sobre emails pre definidos
emailPreLB = tk.Label(text='Emails pré definidos:', font='arial 20 bold')
emailPreLB.place(x=0,y=0)
EmailPre = tk.Label(text='Emails padrão de envio, coordenadora e a enfermeira de sua unidade',font='arial 10 italic')
EmailPre.place(x=0,y=35)


#criando estilos da treeview
style = ttk.Style()
style.configure('tabelaBR',background='silver')


#criando interface para lista de contatos
FrameBru = tk.Frame(master=janela,width=250,height=120,bg='red')#frame Brumadinho
FrameBru.place(x=450,y=7)
#treeview Brumadinho
tabelaBR = ttk.Treeview(FrameBru, selectmode='browse', columns=('column1'), show='headings')


tabelaBR.column('column1',width=250,minwidth=50,stretch=NO)
tabelaBR.heading('#1',text='Lista de emails Brumadinho')
tabelaBR.place(x=0,y=0)

#botao deletar email brumadinho
DellBr = tk.Button(text='Deletar Email Brumadinho', command=DelEmailBR)
DellBr.place(x=450,y=135)


FrameCont = tk.Frame(master=janela,width=250,height=120,bg='green')#frame contagem
FrameCont.place(x=450,y=200)

#treeview contagem
tabelaCont = ttk.Treeview(FrameCont,selectmode='browse',columns=('column1'),show='headings')



tabelaCont.column('column1',width=250,minwidth=50,stretch=NO)
tabelaCont.heading('#1',text='Lista de emails Contagem')

tabelaCont.place(x=0,y=0)
#botao deletar email contagem
DellCont = tk.Button(text='Deletar Email Contagem',command=DelEmailCont)
DellCont.place(x=450,y=340)

FrameUbl = tk.Frame(master=janela,width=250,height=120,bg='blue')#frame Uberlandia
FrameUbl.place(x=750,y=466)

#treeview uberlandia
tabelaUbl = ttk.Treeview(FrameUbl,selectmode='browse',columns=('column1'),show='headings')



tabelaUbl.column('column1',width=250,minwidth=50,stretch=NO)
tabelaUbl.heading('#1',text='Lista de emails Uberlandia')
tabelaUbl.place(x=0,y=0)
#botao deleter email uberlandia
Dellubl = tk.Button(text='Deletar Email Uberlandia',command=DellEmailUbl)
Dellubl.place(x=750,y=600)


FrameJf = tk.Frame(master=janela,width=250,height=120,bg='yellow')#frame Juiz de Fora
FrameJf.place(x=1050,y=7)

#treeview Juiz de fora
tabelaJf = ttk.Treeview(FrameJf,selectmode='browse',columns=('column1'),show='headings')



tabelaJf.column('column1',width=250,minwidth=50,stretch=NO)
tabelaJf.heading('#1',text='Lista de emails Juiz de Fora')
tabelaJf.place(x=0,y=0)
#botao deleter email Juiz de Fora
DellJF = tk.Button(text='Deletar Email Juiz de Fora',command=DellEmailJf)
DellJF.place(x=1050,y=135)

FrameLib = tk.Frame(master=janela,width=250,height=120,bg='orange')#frame Libertas
FrameLib.place(x=450,y=465)

#treeview Libertas
tabelaLib = ttk.Treeview(FrameLib,selectmode='browse',columns=('column1'),show='headings')



tabelaLib.column('column1',width=250,minwidth=50,stretch=NO)
tabelaLib.heading('#1',text='Lista de emails Libertas')
tabelaLib.place(x=0,y=0)
#botao deleter email Libertas
DellLib = tk.Button(text='Deletar Email Libertas',command=DellEmailLib)
DellLib.place(x=450,y=600)


FrameDiv = tk.Frame(master=janela,width=250,height=120,bg='black')#frame Divinopolis
FrameDiv.place(x=1050,y=466)

#treeview Divinopolis
tabelaDiv = ttk.Treeview(FrameDiv,selectmode='browse',columns=('column1'),show='headings')



tabelaDiv.column('column1',width=250,minwidth=50,stretch=NO)
tabelaDiv.heading('#1',text='Lista de emails Divinopolis')
tabelaDiv.place(x=0,y=0)
#botao deleter email Divinopolis
DellDiv = tk.Button(text='Deletar Email Divinopolis',command=DellEmailDiv)
DellDiv.place(x=1050,y=590)


FrameDia = tk.Frame(master=janela,width=250,height=120,bg='pink')#frame Diamantina
FrameDia.place(x=750,y=7)

#treeview Diamantina
tabelaDia = ttk.Treeview(FrameDia,selectmode='browse',columns=('column1'),show='headings')



tabelaDia.column('column1',width=250,minwidth=50,stretch=NO)
tabelaDia.heading('#1',text='Lista de emails Diamantina')
tabelaDia.place(x=0,y=0)
#botao deleter email Diamantina
DellDia = tk.Button(text='Deletar Email Diamantina',command=DellEmailDia)
DellDia.place(x=750,y=135)

FrameGV = tk.Frame(master=janela,width=250,height=120,bg='grey')#frame Governador Valadares
FrameGV.place(x=750,y=200)

#treeview Governador valadares
tabelaGV = ttk.Treeview(FrameGV,selectmode='browse',columns=('column1'),show='headings')


tabelaGV.column('column1',width=250,minwidth=50,stretch=NO)
tabelaGV.heading('#1',text='Lista de emails Governador Valadares')
tabelaGV.place(x=0,y=0)
#botao deleter email Diamantina
DellGV = tk.Button(text='Deletar Email Gv',command=DellEmailGV)
DellGV.place(x=750,y=330)

FrameMoc = tk.Frame(master=janela,width=250,height=120,bg='brown')#frame Montes claros
FrameMoc.place(x=1050,y=200)

#treview Montes claros
tabelaMoc = ttk.Treeview(FrameMoc,selectmode='browse',columns=('column1'),show='headings')



tabelaMoc.column('column1',width=250,minwidth=50,stretch=NO)
tabelaMoc.heading('#1',text='Lista de emails Montes Claros')
tabelaMoc.place(x=0,y=0)
#botao deleter email Diamantina
DellMoc = tk.Button(text='Deletar Email Montes Claros',command=DellEmailMoc)
DellMoc.place(x=1050,y=330)




#criando combobox para adicionar os emails
textCombo = tk.Label(text='Selecione a filial')
textCombo.place(x=235,y=100)
listaCombo = ttk.Combobox(janela,values=["","Brumadinho","Juiz de Fora","Governador Valadares","Contagem","Uberlandia","Montes Claros",
                                         "Libertas","Divinopolis","Diamantina"])
listaCombo.current(0)
listaCombo.place(x=235,y=125)

#configurando interface de digitação
email_lbl= tk.Label(text='Insira o Email:')
email_lbl.place(x=5,y=100)

email_entry = tk.Entry(width=30)#campo de digitação
email_entry.place(x=5,y=125)#campo de digitação

botao_email = tk.Button(text='Inserir email na lista',width=30,command=InserirEmail)#inserindo email na lista
botao_email.place(x=5,y=155)

texto_email = tkinter.Text(width=50,height=10)
texto_email.place(x=5,y=195)

btnEnviar = tk.Button(text='Enviar emails',font='arial 10',border=3,command=enviarEmail)#envia os emails
btnEnviar.place(x=5,y=375)

#acessando banco para pegar os dados ja gravados para a tabela
def MostrarDadosBru():
    consulta_sql = """SELECT * FROM Brumadinho"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()


    for linhas in linha:


        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'",""))

        tabelaBR.insert("", END, values=text, tag='1')








def mostrarDadosCont():
    consulta_sql = """SELECT * FROM Contagem"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaCont.insert("", END, values=text, tag='1')




def mostrarDadosMoc():
    consulta_sql = """SELECT * FROM Moc"""
    cursor.execute(consulta_sql)

    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaMoc.insert("", END, values=text, tag='1')




def mostrarDadosJF():
    consulta_sql = """SELECT * FROM Jf"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaJf.insert("", END, values=text, tag='1')




def mostrarDadosGv():
    consulta_sql = """SELECT * FROM Gv"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaGV.insert("", END, values=text, tag='1')




def mostrarDadosUbl():
    consulta_sql = """SELECT * FROM Ubl"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaUbl.insert("", END, values=text, tag='1')




def mostrarDadosLib():
    consulta_sql = """SELECT * FROM Libertas"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaLib.insert("", END, values=text, tag='1')




def mostrarDadosDiv():
    consulta_sql = """SELECT * FROM Divinopolis"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaDiv.insert("", END, values=text, tag='1')




def mostrarDadosDia():
    consulta_sql = """SELECT * FROM Diamantina"""
    cursor.execute(consulta_sql)
    linha = cursor.fetchall()

    for linhas in linha:
        text1 = str(linhas)

        text = str(text1.replace(",", "").replace("(", "").replace(")", "").replace("'", ""))

        tabelaDia.insert("", END, values=text, tag='1')



MostrarDadosBru()# chamando função mostrar dados da tabela Brumadinho

mostrarDadosDia() #chamando função mostrar dados da tabela Diamantina

mostrarDadosLib() #chamando função mostrar dados da tabela Libertas

mostrarDadosUbl() #chamando função mostrar dados da tabela Uberlandia

mostrarDadosGv()  #chamando função mostrar dados da tabela Governador Valadares

mostrarDadosJF()  #chamando função mostrar dados da tabela Juiz de Fora

mostrarDadosCont()#chamando função mostrar dados da tabela Contagem

mostrarDadosMoc() #chamando função mostrar dados da tabela Montes Claros

mostrarDadosDiv() #chamando função mostrar dados da tabela Divinopolis


janela.mainloop()






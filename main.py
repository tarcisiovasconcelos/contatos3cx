
from os import sep
import re
from unittest.mock import DEFAULT
import PySimpleGUI as sg
import pandas as pd
from datetime import datetime
import getpass
from configparser import ConfigParser
import banco
import os


sg.theme("Default1")

frame_layout = [
                [sg.T("\n",background_color="white", size=(1)),sg.T("Olá, essa interface vai ajudar você a atualizar a base de dados 3CX.\n\n1º Escolha o arquivo de Contatos do 3CX geralmente tem o nome Contacts e clique em enviar.\n2º Escolha o arquivo de Telefones do 3CX geralmente tem o nome Phones e clique em enviar.\n3º Clique em Atualizar BD para que o sistema faça a att do mesmo.\n\n- O sistema insere no banco se houver contato novo.\n- O sistema atualiza no banco se houver mudança no contato\n- O sistema deleta do banco se o contato não existe no 3CX",background_color="white")],
               ]
# Layout

layout = [
        [sg.Image(filename="cesmac.png")],
        [sg.T("\n",background_color="white")],
        [sg.T("\n",background_color="white", size=(10,10)),sg.Frame('Guia de Uso', frame_layout ,background_color='white', font='Any 12', title_color='black', size=(627,200)),sg.T("\n",background_color="white", size=(10,10))],
        [sg.T("\n",background_color="white")],
        [sg.T("\n",background_color="white", size=(10)),sg.Text("Escolha o arquivo 1 Contacts: ",background_color="white",size= 30), sg.Input(key="-IN2-" ,change_submits=True, size=(34)),sg.FileBrowse("Selecionar arquivo",key="-IN-", size=(15), file_types=(("Text Files", "*.ini"),))],
        [sg.T("\n",background_color="white", size=(10)),sg.Text("Escolha o arquivo 2 Phones: ",background_color="white",size=30), sg.Input(key="-IN3-" ,change_submits=True, size=(34)),sg.FileBrowse("Selecionar arquivo",key="-IN1-",disabled=True, size=(15), file_types=(("Text Files", "*.ini"),))],
        [sg.T("\n",background_color="white")],
        [sg.T("\n",background_color="white", size=(10)),sg.Button("Enviar Contatos", disabled=True, size=(15)),sg.T("",background_color="white",key='sendContact', size=(15)),sg.T("\n",background_color="white", size=(27)),sg.Button("Atualizar BD", disabled=True, size=(15))],
        [sg.T("\n",background_color="white", size=(10)),sg.Button("Enviar Telefones", disabled=True, size=(15)),sg.T("",background_color="white",key='sendPhones', size=(15)),sg.T("\n",background_color="white", size=(27)),sg.Button("Arquivo Rev", disabled=False, size=(15))]

]

# Janela
window = sg.Window('Contatos', layout, size=(848,600),background_color="white")

# Eventos   
while True:
    log = []
    #Inicia a janela
    event, values = window.read()

    #Descobre o username do user do windows
    username = getpass.getuser()

    #Variavel de caminho 
    fpath = f"C:\\Users\\{username}\\Downloads\\"  

    #Primeiro evento, aqui ele permite fechar a aplicação 
    if event == sg.WIN_CLOSED or event=="Exit":
        break

    #Só libera o btn de Enviar quando escolhe o arquivo
    elif event == '-IN2-':
        window['Enviar Contatos'].update(disabled=False)

    #Só libera o btn de Enviar quando escolhe o arquivo
    elif event == '-IN3-':
        window['Enviar Telefones'].update(disabled=False)

    #Evento responsável por tratar o arquivo de contato do 3CX    
    elif event == "Enviar Contatos":

        hora = str(datetime.today().strftime('%H:%M:%S')) 

        #Primeiro crio o contatos.ini a partir do arquivo de contato enviado para converter UTF-16 em LATIN1, só assim ele abre o configparser
        with open(values["-IN-"], 'rb') as source_file:
            with open('contatos.ini', 'w+b') as dest_file:
                contents = source_file.read()
                dest_file.write(contents.decode('UTF-16').encode('LATIN1'))

        #Abre o config parser
        configContatos = ConfigParser()
        configContatos.read('contatos.ini') 

        #Desativo o enviar arquivo de contatos e libero para inserir o segundo arquivo de telefones
        window['Enviar Contatos'].update(disabled=True)
        window['-IN1-'].update(disabled=False)
        window['sendContact'].update('Enviado ás:'+ str(hora))

        print('Contacts Enviado')

    #Evento responsável por tratar o arquivo de telefones do 3CX 
    elif event == "Enviar Telefones":

        hora = str(datetime.today().strftime('%H:%M:%S')) 

        #Primeiro crio o telefones.ini a partir do arquivo de telefone enviado para converter UTF-16 em LATIN1, só assim ele abre o configparser
        with open(values["-IN1-"], 'rb') as source_file:
            with open('telefones.ini', 'w+b') as dest_file:
                contents = source_file.read()
                dest_file.write(contents.decode('UTF-16').encode('LATIN1')) 

        #Abre o config parser
        configPhones = ConfigParser()
        configPhones.read('telefones.ini')
 
        #Desativo o enviar arquivo de telefones e libero o Gerar
        window['Atualizar BD'].update(disabled=False)
        window['Enviar Telefones'].update(disabled=True)
        window['sendPhones'].update('Enviado ás:'+ str(hora)) 

        print('Phones Enviado')

    #Evento responsável por Fazer nova carga no banco de dados com os dados dos arquivos de contato e telefone do 3CX 
    elif event == "Atualizar BD":

        hora = str(datetime.today().strftime('%H:%M:%S'))
        hora1 = str(datetime.today().strftime('%H:%M:%S').replace(":", "")) 
        data = str(datetime.today().strftime("%d%m%Y"))
         

        #Criando lista de phones pra indexar o ID AO TELEFONE e criar um dataframe para o merge a seguir
        listaPhones = []
        for telefone in configPhones.sections():
            lista = [str(configPhones[telefone]['contact']),str(telefone)]
            listaPhones.append(lista)

        #DATA FRAME DE TELEFONES BRUTO
        dfPhones = pd.DataFrame(listaPhones, columns=['id', 'telefone'])

        #Criando lista de contatos pra indexar o ID AO CONTATO e criar um dataframe para o merge a seguir
        listaContatos= []
        for id in configContatos.sections():
            lista = [str(id),str(configContatos[id]['displayName'])]
            listaContatos.append(lista)
            
        #DATA FRAME DE CONTATOS BRUTO    
        dfContatos = pd.DataFrame(listaContatos, columns=['id', 'contato'])

        #MERGGE DOS DOIS DATA FRAMES PARA DAR ORIGEN AO DATAFRAME QUE VOU USAR PARA INSERIR NO BANCO , FAZER UPDATES E VALIDAÇÕES VIA SQL's
        dfPerfil = pd.merge(dfContatos, dfPhones, on = "id")

        #SQL DE CONSULTA PARA ENCONTRAR QUEM ESTÁ DEPRECIADO NO BANCO 
        banco.cursor.execute(banco.sqlSelectALL)
        row = banco.cursor.fetchall()
        #Criei a lista de id de quem está no banco e no arquivo
        listaBanco = [] 
        listaArquivos = []
        for x in row:
            listaBanco.append(x[0])
        for x in dfPerfil['id']:
            listaArquivos.append(x)
        #Verifica se esta depreciado e deleta         
        for id in listaBanco:
            if id not in listaArquivos:
                banco.cursor.execute(banco.sqlSelect, id)
                itemDepreciado = banco.cursor.fetchone()
                log.append(str(' *DELETE* ' + ' -ID: ' + str( id) + ' -Setor Deletado: ' + str( itemDepreciado[0]) + ' -Telefone Deletado: ' + str( itemDepreciado[1]) + '\n'))
                banco.cursor.execute(banco.sqlDelete, id)
                banco.cnxn.commit()

        #Abrindo o DATAFRAME de perfil para rodar os SQL's  
        for index, perfil in dfPerfil.iterrows():
            #salvando o ID , TELEFONE , SETOR dos arquivos 3CX
            id = str(perfil[0])
            telefoneNovo = str(perfil[2])
            setorNovo = str(perfil[1])
            #SQL DE CONSULTA POR ID 
            banco.cursor.execute(banco.sqlSelect, id)
            #SE TIVER NO BANCO VAI RETORNAR OS CAMPOS SETOR E TELEFONE E SE NÃO ESTIVER VAI RETORNAR NONE
            row = banco.cursor.fetchone() 
            #Se o contato já estiver no banco ele vai entrar nesse if
            if row != None:
                #Checando se os setores são iguais
                setorAtual = row[0]
                if setorAtual != setorNovo:
                    #SQL DE UPDATE SETOR
                    banco.cursor.execute(banco.sqlUpdateSetor,setorNovo,id)
                    banco.cnxn.commit()
                    log.append(str(' *UPDATE* ' + ' -ID: ' + str(id) + ' -Setor Atualizado: ' + str(setorAtual) + ' => ' + str(setorNovo) + '\n'))
                #Checando se os telefones são iguais
                telefoneAtual = row[1]
                if telefoneAtual != telefoneNovo:
                    #SQL UPDATE TELEFONE
                    banco.cursor.execute(banco.sqlUpdateTelefone,telefoneNovo,str(perfil[0]))
                    banco.cnxn.commit()
                    log.append(str(' *UPDATE* ' + ' -ID: ' + str(id) + ' -Telefone Atualizado: ' + str(telefoneAtual) + ' => ' + str(telefoneNovo) + '\n'))
            #Se o contato não estiver no banco ele vai entrar nesse else e inserir 

            else:
                #SQL DE INSERT
                banco.cursor.execute(banco.sqlInsert, id,setorNovo,telefoneNovo)
                banco.cnxn.commit()
                log.append(str(' *INSERT* ' + ' -ID: ' + str(id) + ' -Setor: ' + str(setorNovo) + ' -Telefone: ' + str(telefoneNovo) + '\n'))
        for x in log:
            print(x)
        for i in range(1,100):
            sg.one_line_progress_meter('Contatos 3CX', i+1, 100, 'Atualizando Base de Dados' ,orientation='h')
        #window['sendATT'].update('Atualizado ás:'+ str(hora))
        with open(fpath + f"LOG_3CX_{data}_{hora1}.txt", "w+") as arquivo:
            arquivo.writelines("")
            arquivo.close
            for x in log:
                with open(fpath +f"LOG_3CX_{data}_{hora1}.txt", "a+") as arquivo:
                    arquivo.writelines(x)
                    arquivo.close 

        if os.path.exists("contatos.ini"):
            os.remove("contatos.ini")
 
        if os.path.exists("telefones.ini"):
            os.remove("telefones.ini")

    elif event == 'Arquivo Rev':

        lista = []
        listaContatoString = []
        listaTelefoneString = []

        data = str(datetime.today().strftime("%d%m%Y"))

        banco.cursor.execute(banco.sqlSelectALL)
        row = banco.cursor.fetchall()
        for contato in row:
            contatos = [contato[0],contato[1],contato[2]]
            lista.append(contatos)

        df = pd.DataFrame(lista, columns=['id', 'contato', 'telefone'])

        
        contatoConfig = ConfigParser()
        telefoneConfig = ConfigParser()
        
        for index, contato in df.iterrows():
            id = contato[0]
            setor = contato[1]
            telefone = contato[2]

            stringContato = f'[{id}]\nDisplayName={setor}\nFirstName=\nLastName=\neMail=\nDefault=\nImage=\nType=2\n'
            stringTelefone = f'[{telefone}]\nName=Phone\nContact={id}\nType=2\n'

            listaContatoString.append(stringContato)
            listaTelefoneString.append(stringTelefone)  

        
        with open('ContactsReverseTXT.txt', "w+", encoding='UTF-16') as arquivo:
            arquivo.writelines(listaContatoString)
            arquivo.close

        with open('PhonesReverseTXT.txt', "w+", encoding='UTF-16') as arquivo:
            arquivo.writelines(listaTelefoneString)
            arquivo.close
        
        with open('ContactsReverseTXT.txt', 'rb') as source_file:
            with open(fpath + f'ContactsReverse_{data}.ini', 'w+b') as dest_file:
                contents = source_file.read()
                dest_file.write(contents.decode('UTF-16').encode('LATIN1'))

        with open('PhonesReverseTXT.txt', 'rb') as source_file:
            with open(fpath + f'PhonesReverse_{data}.ini', 'w+b') as dest_file:
                contents = source_file.read()
                dest_file.write(contents.decode('UTF-16').encode('LATIN1'))
        
        if os.path.exists("ContactsReverseTXT.txt"):
            os.remove("ContactsReverseTXT.txt")

        if os.path.exists("PhonesReverseTXT.txt"):
            os.remove("PhonesReverseTXT.txt")
        sg.popup('Arquivo Criado em Downloads!')


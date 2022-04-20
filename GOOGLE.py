#Recursos
from distutils.util import execute
import pandas as pd
import numpy as np
import re
import os.path
import pickle
from datetime import date, datetime
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


class gapi:
    def __init__(self, number):
        self.number = number

    def call(self):
        #Si se modifican los SCOPES, borrar el archivo token.pickle.
        SCOPES = ['https://mail.google.com/','https://www.googleapis.com/auth/spreadsheets']
        #Proceso de conexión
        self.creds = None
        #Busqueda de token
        if os.path.exists('token.pickle'+str(self.number)+''):
            with open('token.pickle'+str(self.number)+'', 'rb') as token:
                self.creds = pickle.load(token)
        # If there are no (valid) credentials available, let the user log in.
        if not self.creds or not self.creds.valid:
            if self.creds and self.creds.expired and self.creds.refresh_token:
                self.creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    '/Users/keilamarin/Documents/Reserve/Bot_ZELLE/vast_creds.json', SCOPES)
                self.creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.pickle'+str(self.number)+'', 'wb') as token:
                pickle.dump(self.creds, token)

    def mail(self):
        #Llamar al API
        self.call()
        #Servicio del API
        gmail = build('gmail', 'v1', credentials=self.creds)

        #Encontrar LabelId de etiquta CHASE
        results = gmail.users().labels().list(
            userId='me'
        ).execute()
        labels = results.get('labels',[])
        
        data_list = []
        #Extraer LabelId
        for label in labels:
            if label['name'] == 'PENDIENTE':
                tag=label['id']
                print('Etiqueta', label['name'], 'encontrada', tag)

                #Filtrar mensajes por lable y no leidos
                results = gmail.users().messages().list(
                    userId='me', 
                    labelIds=[str(tag)],
                    q='is:unread'
                ).execute()
                messages = results.get('messages', [])
                #Extraer metadata
                for message in messages:
                    msg = gmail.users().messages().get(
                        userId='me', 
                        id=message['id'],  
                        format='metadata',
                        metadataHeaders=['Delivered-To','Received','Subject']
                    ).execute()
                    #Marcar como leido
                    results = gmail.users().messages().modify(
                        userId='me',
                        id=message['id'],
                        body={
                            "addLabelIds": [],
                            "removeLabelIds": ['UNREAD',str(tag)]
                        }
                    ).execute()


                    #Datos extraidos
                    recipient = str(msg['payload']['headers'][0]['value']) #Correo receptor
                    date = str((re.split(',|-|\n',str(msg['payload']['headers'][1]['value'])))[1]) #Fecha
                    name = str((str(msg['payload']['headers'][7]['value']).split('$'))[0]).replace(' te envió ','').upper() #Nombre
                    amount = str((str(msg['payload']['headers'][7]['value']).split('$'))[1]) #Monto
                    data = {
                        'MsgId':message['id'],
                        'Cuenta':recipient,
                        'Fecha':date,
                        'Remitente':name,
                        'Monto':amount,
                        'Estatus':'POR PAGAR'  
                    }
                    data_list.append(data)
        df = pd.DataFrame(data_list)
        print('Correos recogidos')
        return(df)
    
    def sheets(self, emails, column, row):
        self.emails = emails
        self.column = str(column).upper()
        #self.row = str(row).lower

        #Llamar al API
        self.call()
        #Servicio del API
        sheets = build('sheets','v4', credentials=self.creds)
        spreadsheet_id = '125DCd6QZ8AUdZ8hgFNxGQHMGXOi8QDcCYb6OlNMq9Do'

        #Armar lista y conseguir la primera fila vacia en sheets
        rows_sheets = sheets.spreadsheets().values().get(
            spreadsheetId = spreadsheet_id,
            majorDimension = 'ROWS',
            range = 'Log Operaciones!A1:A'
        ).execute()
        #Crear lista con MsgId ya existentes
        log = []
        for i in range(len(rows_sheets['values'])):
            log.append(rows_sheets['values'][i][0])
        #Encontrar última fila en el log
        if row=='last':
            last_row = (len(rows_sheets['values']))+1
            print("Ultima fila con datos en el log:",last_row)
        else:
            last_row = row
        
        #Comparar
        data_list= []
        data = emails.to_dict('Records')
        for r in data:
            if row=='last':
                if r['MsgId'] not in log:
                    #Crear lista a introducir en sheets
                    data_list.append(r)
            else:
                data_list.append(r)


        df = pd.DataFrame(data_list)
        entries = df.T.reset_index().T.values.tolist() #Trasponer Dataframe para exportar a sheets
        del entries[0] #Borrar headers de DataFrame

        #Crear diccionario para cumplir con parametro body
        dict = {
            'majorDimension' : 'ROWS',
            'values' : entries
        }
        #Limpiar fila
        clear = sheets.spreadsheets().values().clear(
            spreadsheetId = spreadsheet_id,
            range = 'Log Operaciones!'+str(self.column)+str(last_row)
        )
        clear.execute()
        #Exportar datos
        response = sheets.spreadsheets().values().update(
            spreadsheetId = spreadsheet_id,
            valueInputOption = 'USER_ENTERED',
            range = 'Log Operaciones!'+str(self.column)+str(last_row),
            body = dict
        )
        response.execute()
        print('Sheet actualizado')
    
    def credentials(self,type):
        self.type = type
        
        #Llamar al API
        self.call()
        #Servicio del API
        sheets = build('sheets','v4', credentials=self.creds)
        spreadsheet_id = '125DCd6QZ8AUdZ8hgFNxGQHMGXOi8QDcCYb6OlNMq9Do'

        #Credenciales Reserve
        credentials = sheets.spreadsheets().values().get(
            spreadsheetId = spreadsheet_id,
            majorDimension = 'COLUMNS',
            range = 'Credenciales!B3:B5'
        ).execute()

        #Definir credenciales por tipo de argumento
        if self.type == 'link':
            response = credentials['values'][0][0]
        elif self.type == 'user':
            response = credentials['values'][0][1]
        elif self.type == 'password':
            response = credentials['values'][0][2]
        return response
    
    def compare(self):
        #Llamar al API
        self.call()
        #Servicio del API
        sheets = build('sheets','v4', credentials=self.creds)
        spreadsheet_id = '125DCd6QZ8AUdZ8hgFNxGQHMGXOi8QDcCYb6OlNMq9Do'

        #Lista de nombres en el sheet
        log = []
        names = sheets.spreadsheets().values().get(
            spreadsheetId = spreadsheet_id,
            majorDimension = 'COLUMNS',
            range = 'Log Operaciones!A2:F'
        ).execute()
        #Crear Dataframe para iterar
        for i in range(len(names['values'][0])):
            if str(names['values'][5][i])=='POR PAGAR':
                data = {
                    'MsgId':names['values'][0][i],
                    'Correo':names['values'][1][i],
                    'Fecha':datetime.strptime(names['values'][2][i],'%d %b %Y %H:%M:%S').date(),
                    'Nombre':names['values'][3][i],
                    'Monto':names['values'][4][i]
                }
                log.append(data)
        df = pd.DataFrame(log)
        return df

    def lookup(self):
        #Llamar al API
        self.call()
        #Servicio del API
        sheets = build('sheets','v4', credentials=self.creds)
        spreadsheet_id = '125DCd6QZ8AUdZ8hgFNxGQHMGXOi8QDcCYb6OlNMq9Do'

        #Lista de transid en el sheet
        log = []
        ids = sheets.spreadsheets().values().get(
            spreadsheetId = spreadsheet_id,
            majorDimension = 'COLUMNS',
            range = 'Log Operaciones!A2:H'
        ).execute()

        #Crear Dataframe para iterar
        index = 0
        for i in range(len(ids['values'][0])):
            index+=1
            data = {
                'Trans Id':ids['values'][7][i],
                'Index':index
            }
            log.append(data)
        df = pd.DataFrame(log)
        return df
        
    
    def mark(self, message_id):
        self.message_id = message_id
        #Llamar API
        self.call()
        #Servicio del API
        gmail = build('gmail', 'v1', credentials=self.creds)
        
        #Encontrar LabelId de etiquta CHASE
        results = gmail.users().labels().list(
            userId='me'
        ).execute()
        labels = results.get('labels',[])
        data_list = []
        #Extraer LabelId
        for label in labels:
            if label['name'] == 'COMPLETADA POR BOT':
                tag=label['id']
                print('Etiqueta', label['name'], 'encontrada', tag)

        #Modificar etiquetas
        results = gmail.users().messages().modify(
            userId='me',
            id=str(message_id),
            body={
                "addLabelIds": [str(tag)],
                "removeLabelIds": ['UNREAD']
            }
        ).execute()
    








        


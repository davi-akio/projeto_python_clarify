from sys import exception
import pandas as pd
import re
import smtplib
import openpyxl
import email
from email.message import EmailMessage
from time import sleep

def nome():
    nome()



class Projeto_final:

    def iniciar(self):
        self.lista_tarefas = []
        self.email_destino()
        self.menu()
        self.criar_planilha()
        sleep(5)
        self.enviar_email()
    
    def email_destino(self):
        while True:
            self.email = input("Email de destino : ").lower()
            email_padrao = re.search("^[a-z0-9_.-]+@[a-z0-9]+.[a-z]+(.[a-z]+)?$" , self.email)
            # Regex ( " ^ " começa a validação " $ " termina a validação )
            # [] obrigatório () opcional


            if email_padrao:
                print("Email Válido !")
                break
            else: 
                print("Email inválido !")

    def menu(self):
        while True:
            menu_principal = int(input("""
        [1] Cadastrar
        [2] Visualizar
        [3] Sair
        Opção : """))
            
            match menu_principal:
                case 1:
                    self.cadastrar()
                case 2:
                    self.visualizar()
                case 3:
                    break     
                case _:
                    print("Opção Inválida")

    def cadastrar(self):
        while True :
            tarefa = input("Tarefa ou s para sair: ").capitalize()
            if tarefa == "S":
                break
            else:
                self.lista_tarefas.append(tarefa)

                try:
                    with open("Tarefas/histórico de tarefas.txt" , "a" , encoding="utf8") as arquivo :
                        arquivo.write(f"{tarefa} \n")

                except FileNotFoundError:
                    print("Arquivo não encontrado !")
    
    def visualizar(self):
        try:
            with open("Tarefas/histórico de tarefas.txt" , "r" , encoding="utf8") as arquivo:
                print(arquivo.read())
        except FileNotFoundError:
            print("Arquivo não encontrado !")

    def criar_planilha(self):
        if len(self.lista_tarefas) > 0:
            try:
                df = pd.DataFrame({"Tarefas" : self.lista_tarefas})
                self.nome_arquivo = input("Nome do arquivo (sem xlsx):")
                df.to_excel(f"Tarefas/{self.nome_arquivo}.xlsx" , index=False)
                print("Planilha está sendo criada ! Aguarde até ser enviada por email... ")

            except exception:
                print("Erro ao criar a planilha !")
                
    def enviar_email(self):
        endereco = "daviakio@gmail.com"
        senha = "qgad oawi upcn jodv"
        msg = EmailMessage()
        msg["From"] = endereco
        msg["To"] = self.email
        msg["Subject"] = "Planilha do programa"
        msg.set_content(
            "A planilha está pronta ! Por favor confira"
        )

        arquivos = [f"{self.nome_arquivo}.xlsx"]

        for arq in arquivos:
            with open(arq , "rb") as arq:
                dados = arq.read()
                nome_arquivo = arq.name

            msg.add_attachment(
                dados ,
                maintype = "application" ,
                subtype = "octet-stream" ,
                filename = nome_arquivo
            )
        server = smtplib.SMTP_SSL("smtp.gmail.com" , 465)
        server.login(endereco , senha , initial_response_ok = True)
        server.send_message(msg)

        print("Planilha foi enviada no email ")















start = Projeto_final()
start.iniciar()


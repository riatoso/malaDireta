# MODULO DE MALA DIRETA INTEGRADO COM OUTLOOK
# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------------
# Created By  : Ricardo Antonio Cardoso
# Created Date: Fev-2022
# version ='2.0'
# ---------------------------------------------------------------------------
import time
import win32com.client as win32
import pandas as pd
import PySimpleGUI as sg


class MalaDireta:
    def __init__(self):
        # PANDAS EMAIL
        self.path = "D:/GITHUB/Projetos/malaDireta/email.xlsx"
        self.tabela = pd.read_excel(self.path)
        # INTEGRAÇÃO COM OUTLOOK
        self.outlook = win32.Dispatch('outlook.application')
        # CRIA O LAYOUT
        sg.theme("Reddit")
        layout = [
            [sg.Text("Tela do sistema de Mala Direta", size=(50, 0))],
            [sg.Text("Campos obrigatorios (*) ", size=20)],
            [sg.Text(" * Assunto do Email: ", size=15), sg.Input(key="assunto", size=(37, 0))],
            [sg.Text(f"PATH: {self.path}", text_color="green", size=(39, 0))],
            [sg.Output(size=(55, 5), key="terminal")],
            [sg.Button('Enviar', disabled=False, size=10)],
            [sg.Button('Finalizar', size=10)]
        ]
        # CRIAR A JANELA
        self.janela = sg.Window('TELA DE ENVIO DE EMAIL MALA DIRETA.', layout=layout).finalize()

    def executa(self):

        # EXECUTA O ENVIO
        while True:
            # PEGA OS VALORES DA JANELA
            self.eventos, self.valores = self.janela.Read()

            if self.eventos == "Enviar":  # BOTÃO ENVIAR
                self.emails = []
                for i, email_excel in enumerate(self.tabela["email"]):
                    if email_excel not in self.emails:
                        # criar um email
                        email = self.outlook.CreateItem(0)
                        # configurar as informações do seu e-mail
                        email.To = email_excel
                        email.Subject = self.valores["assunto"]
                        email.HTMLBody = f"""
                        <p>Prezados, como vocês estão? </p>
                        <p>Vou fazer uma breve apresentação.</p>
                        <p>Meu nome é Ricardo Antonio Cardoso , trabalho a 23 anos no ramo de tecnologia.</p>
                        <p>Participei de implantação de Softwares , bem como de todo parque tecnológico em empresas.</p>
                        <p>Trabalhei em uma Faculdade local como coordenador de T.I (Operacional).</p>
                        <p>Estou buscando reingressar no mercado de T.I , DEVOPS Jr. , ou Pleno/Senior na parte de Infraestrutura.</p>
                        <p> </p>
                        <p>Segue anexo meu Currículo.</p>
                        <p>Abs,</p>
                        <p>---------------------------------</p>
                        <p>Ricardo Antonio Cardoso</p>
                        <p>Analista de TI/DEV</p>
                        <p>WhatsApp (16)99122.0875</p>
                        <p>---------------------------------</p>
                        <p># by Python / I9TI #</p>
                        """

                        # BLOCO QUE ENVIA O EMAIL
                        anexo = "D:/GITHUB/Projetos/malaDireta/CV_TI.PDF"
                        email.Attachments.Add(anexo)
                        email.Send()
                        time.sleep(3)
                        self.emails.append(email_excel)
                        print(f"{i + 1} # Email Enviado. {email_excel}")
                    else:
                        print(f"{i + 1} # Email Duplicado. {email_excel}")
                        continue
                print("ENVIO FINALIZADO.")

            if self.eventos == "Finalizar":  # BOTÃO FINALIZAR
                print("Finalizando o sistema...")
                time.sleep(4)
                break
            if self.eventos == sg.WIN_CLOSED:
                break
        self.janela.close()


email = MalaDireta()
email.executa()

# MODULO DE MALA DIRETA INTEGRADO COM OUTLOOK
# -*- coding: utf-8 -*-
# ----------------------------------------------------------------------------
# Created By  : Ricardo Antonio Cardoso
# Created Date: Fev-2022
# version ='1.1'
# ---------------------------------------------------------------------------
import time
import win32com.client as win32
import pandas as pd


class MalaDireta:

    def executa(self):
        # PANDAS EMAIL
        self.tabela = pd.read_excel('D:/GITHUB/Projetos/malaDireta/teste.xlsx')
        # INTEGRAÇÃO COM OUTLOOK
        self.outlook = win32.Dispatch('outlook.application')

        # EXECUTA O ENVIO
        self.emails = []
        for i, email_excel in enumerate(self.tabela["email"]):
            if email_excel not in self.emails:
                # criar um email
                email = self.outlook.CreateItem(0)
                # configurar as informações do seu e-mail
                email.To = email_excel
                email.Subject = "Curriculo - Ricardo Antonio Cardoso"
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
                print(f"{i + 1} # Email Enviado.")
            else:
                print(f"{i + 1} # Email Duplicado.")
                continue


email = MalaDireta()
email.executa()

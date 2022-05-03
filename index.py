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
        self.path = "email.xlsx"
        self.tabela = pd.read_excel(self.path)
        # INTEGRAÇÃO COM OUTLOOK
        self.outlook = win32.Dispatch('outlook.application')
        # CRIA O LAYOUT
        sg.theme("DarkBlue17")
        layout = [
            [sg.Text("Tela do sistema de Mala Direta", size=(50, 0))],
            [sg.Text("PREENCHA O ASSUNTO E ASSINALE O CORPO DO EMAIL E ANEXO.", size=60)],
            [sg.Text("Assunto do Email: ", size=15), sg.Input(key="assunto", size=(37, 0))],
            # CHECKBOX
            [
                sg.Text("Selecione o Curriculo:", size=20),
                sg.Checkbox("CV_ADM", key="cvadm"),
                sg.Checkbox("CV_TI", key="cvti")],
            [
                sg.Text("Selecione o Corpo do Email: ", size=20),
                sg.Checkbox("ADM", key="adm"),
                sg.Checkbox("TI", key="ti")],
            ##########
            [sg.Text(f"PATH: {self.path}", text_color="green", size=(39, 0))],
            [sg.Output(size=(75, 20), key="terminal")],
            [sg.Button('Enviar', disabled=False, size=10)],
            [sg.Button('Finalizar', size=10)]
        ]
        # CRIAR A JANELA
        self.janela = sg.Window('TELA DE ENVIO DE EMAIL MALA DIRETA.', layout=layout, icon="email.ico").finalize()
        self.anexo_ti = r"c:\MalaDireta\CV_TI.PDF"
        self.anexo_adm = r"c:\MalaDireta\CV_ADM.PDF"
        self.corpo_adm = f"""
                        <p>Meu nome é Ricardo Antonio Cardoso.</p>
                        <p>Trabalhei em uma Faculdade local como Coordenador Administrativo por mais de 3 anos,</p> 
                        <p>fazendo toda a gestão da unidade.</p>
                        <p></p>
                        <p>Coordenador Administrativo.</p>
                        <p>Coordenador de Projetos e Operações.</p>
                        <p>Gerencia administrativa.</p>
                        <p>Analista administrativo.</p>
                        <p></p>
                        <p>Estou buscando reingressar no mercado.<p>
                        <p></p>
                        <p>Segue anexo meu Currículo.</p>
                        <p>Abs,</p>
                        <p>---------------------------------</p>
                        <p>Ricardo Antonio Cardoso</p>
                        <p>Analista de TI/DEV</p>
                        <p>WhatsApp (16)99122.0875</p>
                        <p>---------------------------------</p>
                        <p># by Python / I9TI #</p>
                        """
        self.corpo_ti = f"""
                        <p>Meu nome é Ricardo Antonio Cardoso , trabalho a 23 anos no ramo de tecnologia.</p>
                        <p>Participei de implantação de Softwares , bem como de todo parque tecnológico em empresas.</p> 
                        <p>Trabalhei em uma Faculdade local como Coordenador T.I por mais de 3 anos,</p> 
                        <p>fazendo toda a gestão da unidade.</p>
                        <p></p>
                        <p>Desenvolvimento de sofwares.</p>
                        <p></p>
                        <p>Estou buscando reingressar no mercado.<p>
                        <p></p>
                        <p>Segue anexo meu Currículo.</p>
                        <p>Abs,</p>
                        <p>---------------------------------</p>
                        <p>Ricardo Antonio Cardoso</p>
                        <p>Analista de TI/DEV</p>
                        <p>WhatsApp (16)99122.0875</p>
                        <p>---------------------------------</p>
                        <p># by Python / I9TI #</p>
                        """

    def executa(self):
        # EXECUTA O ENVIO
        while True:
            # PEGA OS VALORES DA JANELA
            self.eventos, self.valores = self.janela.Read()
            if self.eventos == "Enviar":  # BOTÃO ENVIAR
                if self.valores["cvadm"] and self.valores["adm"]:
                    self.emails = []
                    for i, email_excel in enumerate(self.tabela["email"]):
                        if email_excel not in self.emails:
                            # criar um email
                            email = self.outlook.CreateItem(0)
                            # configurar as informações do seu e-mail
                            email.To = email_excel
                            email.Subject = self.valores["assunto"]
                            email.HTMLBody = self.corpo_adm
                            # BLOCO QUE ENVIA O EMAIL
                            anexo = self.anexo_adm
                            email.Attachments.Add(anexo)
                            email.Send()
                            time.sleep(3)
                            self.emails.append(email_excel)
                            print(f"{i + 1} # Email Enviado. {email_excel}")
                        else:
                            print(f"{i + 1} # Email Duplicado. {email_excel}")
                            continue
                    print("ENVIO FINALIZADO, CV_ADM.")
                if self.valores["cvti"] and self.valores["ti"]:
                    self.emails = []
                    for i, email_excel in enumerate(self.tabela["email"]):
                        if email_excel not in self.emails:
                            # criar um email
                            email = self.outlook.CreateItem(0)
                            # configurar as informações do seu e-mail
                            email.To = email_excel
                            email.Subject = self.valores["assunto"]
                            email.HTMLBody = self.corpo_ti
                            # BLOCO QUE ENVIA O EMAIL
                            anexo = self.anexo_ti
                            email.Attachments.Add(anexo)
                            email.Send()
                            time.sleep(3)
                            self.emails.append(email_excel)
                            print(f"{i + 1} # Email Enviado. {email_excel}")
                        else:
                            print(f"{i + 1} # Email Duplicado. {email_excel}")
                            continue
                    print("ENVIO FINALIZADO, CV_TI.")
                else:
                    print("Curriculo de ADM precisa estar no corpo de email de ADM, o mesmo ocorre com T.I.")
                    time.sleep(2)
                    continue

            if self.eventos == "Finalizar":  # BOTÃO FINALIZAR
                print("Finalizando o sistema...")
                time.sleep(4)
                break
            if self.eventos == sg.WIN_CLOSED:
                break
        self.janela.close()


email = MalaDireta()
email.executa()

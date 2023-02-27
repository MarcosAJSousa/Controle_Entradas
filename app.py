from PyQt6 import uic, QtWidgets 
from PyQt6.QtWidgets import QMessageBox
from PyQt6.QtGui import QIntValidator
from reportlab.pdfgen.canvas import Canvas
from reportlab.platypus import Table, TableStyle
from reportlab.lib import colors
from reportlab.platypus import BaseDocTemplate, Frame, Paragraph, PageBreak, PageTemplate, FrameBreak, NextPageTemplate
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import inch
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.enums import TA_CENTER
import webbrowser
from tkinter import filedialog
from database import Data_base
from datetime import datetime
import pandas as pd
import sqlite3

def botao_home():
    home.stackedWidget.setCurrentWidget(home.page)
    home.tableWidget.clearContents()
    home.filtro.setCurrentIndex(-2)
 
def botao_help():
    home.stackedWidget.setCurrentWidget(home.page_help)
    home.stackedWidget_4.setCurrentWidget(home.page_15)
    home.user.setText("")
    home.key.setText("")
 
def backup_lock():
    home.stackedWidget_4.setCurrentWidget(home.page_16)
    home.label_20.setText("")

def login():
    home.label_20.setText("")
    usuario = home.user.text()
    senha = home.key.text()
    if usuario == "SuporteSemu" and senha == "101909":
        home.stackedWidget_4.setCurrentWidget(home.page_17)
    else:
        home.label_20.setText("  Login Incorreto!  ")

def backup_registros():
    try:
        path_name  = filedialog.askdirectory()
        
        cnx = sqlite3.connect("system.db")
        result = pd.read_sql_query("""SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros""", cnx)

        result.to_excel(f"{path_name}/registros.xlsx", sheet_name='controle', index=False)
        
        QMessageBox.about(home, 'Mensagem', ' Planilha gerada com Sucesso! \n ')

        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
       
        cursor.execute("DELETE FROM registros")

        db.connection.commit()
        db.close_connection()

        QMessageBox.about(home, 'Mensagem', ' Os dados de Registros foi excluídos! \n ')
    except:
        QMessageBox.critical(home, 'Mensagem', ' Essa função não pode ser realizada! \n ')

def backup_pessoas():
    try:
        path_name  = filedialog.askdirectory()
       
        cnx = sqlite3.connect("system.db")
        result = pd.read_sql_query("""SELECT cpf, nome, orgao, municipio, telefone, email FROM visitantes ORDER BY nome""", cnx)
 
        result.to_excel(f"{path_name}/pessoas_cadastradas.xlsx", sheet_name='controle', index=False)
       
        QMessageBox.about(home, ' Mensagem ', ' Planilha gerada com Sucesso! \n ')

        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
       
        cursor.execute("DELETE FROM visitantes")

        db.connection.commit()
        db.close_connection()

        QMessageBox.about(home, 'Mensagem', ' Todas as pessoas cadastras foram excluídos! \n ')
    except:
        QMessageBox.critical(home, 'Mensagem', ' Essa função não pode ser realizada! \n ')

def botao_registrar():
    home.stackedWidget.setCurrentWidget(home.page_2)
    data_hoje = datetime.now()
    home.line_data.setText(str(data_hoje.strftime("%d/%m/%Y")))
 
def botao_historico():
    home.stackedWidget.setCurrentWidget(home.page_3)
    home.filtro.setPlaceholderText(" Buscar por:")
    home.stackedWidget_2.setCurrentWidget(home.page_13)
    home.stackedWidget_3.setCurrentWidget(home.page_14)
 
    home.tableWidget.setColumnWidth(0,140)
    home.tableWidget.setColumnWidth(1,500)
    home.tableWidget.setColumnWidth(2,250)
    home.tableWidget.setColumnWidth(3,250)
    home.tableWidget.setColumnWidth(4,150)
    home.tableWidget.setColumnWidth(5,400)
    home.tableWidget.setColumnWidth(6,100)
    home.tableWidget.setColumnWidth(7,120)
 
def botao_novo():
    home.stackedWidget.setCurrentWidget(home.page_cadastro)
 
def cadastrar_novo(fullDataSet):
    try:
        db = Data_base()
        db.connect()
       
        input_cpf = home.line_CPF_2.text()
        input_nome = home.line_nome_2.text().upper()
        input_orgao = home.line_orgao_2.text().upper()
        input_municipio = home.line_muni_2.text().upper()
        input_telfone = home.line_tel_2.text()
        input_email = home.line_email_2.text().lower()
       
        if home.line_CPF_2.text() == "" or home.line_nome_2.text() == "" or home.line_orgao_2.text() == "" or home.line_muni_2.text() == ""  or home.line_tel_2.text() == "" or home.line_email_2.text() == "":
            raise Exception(QMessageBox.warning(home, ' Mensagem ', 'Verifique se todos os dados estão preenchidos corretamente!'))
 
        fullDataSet = ( str(input_cpf),  str(input_nome), str(input_orgao), str(input_municipio), str(input_telfone), str(input_email))
       
        #CADASTRAR NO BANCO DE DADOS
        resp = db.insert_table(fullDataSet)
 
        if resp == "OK":
            QMessageBox.about(home, ' Mensagem ', ' Cadastro realizado com sucesso!  ')
            db.close_connection()
            home.line_CPF_2.setText('')
            home.line_nome_2.setText('')
            home.line_orgao_2.setText('')
            home.line_muni_2.setText('')
            home.line_tel_2.setText('')
            home.line_email_2.setText('')
            return
        else:
            QMessageBox.warning(home, ' Atenção! ', ' Essa CPF já foi cadastrado no sistema!  ')
            db.close_connection()
    except:
        QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
 
def completar():
    try:
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        cpf = home.line_CPF.text()
        cursor.execute("SELECT nome, orgao, municipio, telefone, email FROM visitantes WHERE cpf LIKE '{}' ORDER BY nome ASC;".format(cpf))
        dados = cursor.fetchall()
               
        home.line_nome.setText(str(dados[0][0]))
        home.line_orgao.setText(str(dados[0][1]))
        home.line_muni.setText(str(dados[0][2]))
        home.line_tel.setText(str(dados[0][3]))
        home.line_email.setText(str(dados[0][4]))
         
    except:
        QMessageBox.warning(home, ' Atenção! ', ' Esse CPF não costa no sistema     ')
       
def registrar(fullDataSet):
    try:
        db = Data_base()
        db.connect()
 
        input_cpf = home.line_CPF.text()
        input_nome = home.line_nome.text().upper()
        input_orgao = home.line_orgao.text().upper()
        input_municipio = home.line_muni.text().upper()
        input_telfone = home.line_tel.text()
        input_email = home.line_email.text().lower()
        input_data = home.line_data.text()
        input_destino = home.line_destino.text().upper()
       
        if home.line_CPF.text() == "" or home.line_nome.text() == "" or home.line_orgao.text() == "" or home.line_muni.text() == ""  or home.line_tel.text() == "" or home.line_email.text() == "" or home.line_data.text() == "" or home.line_destino.text() == "":
            raise Exception(QMessageBox.warning(home, ' Mensagem ', 'Verifique se todos os dados estão preenchidos corretamente!'))
                 
        fullDataSet = ( str(input_cpf),  str(input_nome), str(input_orgao), str(input_municipio), str(input_telfone), str(input_email), str(input_data), str(input_destino))
       
        cursor = db.connection.cursor()
        input_cpf = home.line_CPF.text()
        cursor.execute("SELECT cpf FROM visitantes WHERE cpf LIKE '{}'".format(input_cpf) )
        dados = cursor.fetchall()
 
        resp = db.insert_table_2(fullDataSet)
 
        if dados == [] and resp == "OK":
            quetion.show()
           
        elif resp == "OK":
            QMessageBox.about(home, ' Mensagem ', ' Registro realizado com sucesso! ')
            db.close_connection()
            home.line_CPF.setText('')
            home.line_nome.setText('')
            home.line_orgao.setText('')
            home.line_muni.setText('')
            home.line_tel.setText('')
            home.line_email.setText('')
            home.line_data.setText('')
            home.line_destino.setText('')
           
            data_hoje = datetime.now()
            home.line_data.setText(str(data_hoje.strftime("%d/%m/%Y")))
        else:
            QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
            db.close_connection()
    except:
        pass
 
def Consulta_all():
    try:
        db = Data_base()
        db.connect()
        dados_lidos = db.select_all()
       
        home.tableWidget.setRowCount(len(dados_lidos))
       
        for i in range(0, len(dados_lidos)):
            for j in range (0,8):
                home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))
        home.stackedWidget_2.setCurrentWidget(home.page_4)
    except:
        QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
 
def Consulta_filtro():
    try:
        if home.filtro.currentText() == " CPF":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.localizar_cpf.text()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE cpf LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
                           
                home.stackedWidget_2.setCurrentWidget(home.page_5)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
 
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " NOME":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.buscar.text().upper()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE nome LIKE '%{}%'".format(sua_busca))
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_6)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
               
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " ÓRGÃO/EMPRESA":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.buscar.text().upper()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE orgao LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_7)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " MINICÍPIO":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.buscar.text().upper()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE municipio LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_10)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " TELEFONE":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.localizar_tel.text()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE telefone LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_8)
               
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " EMAIL":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.localizar_email.text().lower()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE email LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_9)
                               
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', 'Algo não saiu como planejado!')
       
        elif home.filtro.currentText() == " DATA":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.localizar_data.text()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE data LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_11)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        elif home.filtro.currentText() == " DESTINO":
            try:
                db = Data_base()
                db.connect()
                cursor = db.connection.cursor()
                sua_busca = home.buscar.text().upper()
                cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE destino LIKE '%{}%'".format(sua_busca) )
                dados = cursor.fetchall()
               
                home.tableWidget.setRowCount(len(dados))
                home.tableWidget.setColumnCount(8)
                for i in range(0, len(dados)):
                    for j in range (0,8):
                        home.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(str(dados[i][j])))
               
                home.stackedWidget_2.setCurrentWidget(home.page_12)
 
                if dados == []:
                    QMessageBox.about(home, ' Mensagem ', '\n   Não foi encontrado dados referente a sua pesquisa!\n  \n Verifique os dados de sua pesquisa e a caixa de seleção       ')
                    home.stackedWidget_2.setCurrentWidget(home.page_13)
            except:
                QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
       
        else:
            QMessageBox.about(home, ' Mensagem ', '     Selecione um filtro para fazer a pesquisa!     ')
            home.tableWidget.clearContents()
            home.stackedWidget_2.setCurrentWidget(home.page_13)
 
    except:
        QMessageBox.critical(home, ' Mensagem ', ' Algo não saiu como planejado! ')
        pass
 
def validar_edit(fullDataSet):
    try:
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        busca = home.line_CPF.text()
        cursor.execute("SELECT nome, orgao, municipio, telefone, email, cpf FROM visitantes WHERE cpf LIKE '{}';".format(busca))
        dados = cursor.fetchall()
               
        nome = (str(dados[0][0]))
        orgao = (str(dados[0][1]))
        municipio =(str(dados[0][2]))
        telefone = (str(dados[0][3]))
        email = (str(dados[0][4]))
        cpf = (str(dados[0][5]))
 
        input_nome = home.line_nome.text().upper()
        input_orgao = home.line_orgao.text().upper()
        input_municipio = home.line_muni.text().upper()
        input_telfone = home.line_tel.text()
        input_email = home.line_email.text().lower()
 
        if input_nome != nome or input_orgao != orgao or input_municipio != municipio or input_telfone != telefone or input_email != email:
            quetion2.show()
            quetion2.label_4.setText(f'Alterações foram feitas no cadastro do cpf: {cpf}')
            quetion2.nome_old.setText(nome)  
            quetion2.orgao_old.setText(orgao)  
            quetion2.muni_old.setText(municipio)  
            quetion2.tel_old.setText(telefone)  
            quetion2.email_old.setText(email)
 
            if input_nome != nome:
                quetion2.nome_new.setText(input_nome)
            if input_orgao != orgao:
                quetion2.orgao_new.setText(input_orgao)
            if input_municipio != municipio:
                quetion2.muni_new.setText(input_municipio)
            if input_telfone != telefone:
                quetion2.tel_new.setText(input_telfone)
            if input_email != email:
                quetion2.email_new.setText(input_email)
        else:
            registrar(fullDataSet)
    except:
        registrar(fullDataSet)
 
def nao_nao():
    try:
        quetion2.close()
        completar()
        quetion2.nome_new.setText('')
        quetion2.orgao_new.setText('')
        quetion2.muni_new.setText('')
        quetion2.tel_new.setText('')
        quetion2.email_new.setText('')
    except:
        pass
 
def nao_sim(fullDataSet):
    try:
        quetion2.close()
        registrar(fullDataSet)
        quetion2.nome_new.setText('')
        quetion2.orgao_new.setText('')
        quetion2.muni_new.setText('')
        quetion2.tel_new.setText('')
        quetion2.email_new.setText('')
    except:
        pass
 
def sim_sim(fullDataSet):
    try:
        quetion2.close()
 
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        busca = home.line_CPF.text()
        cursor.execute("SELECT id FROM visitantes WHERE cpf LIKE '{}';".format(busca))
        dados = cursor.fetchall()
               
        id = (str(dados[0][0]))
 
        nome = home.line_nome.text().upper()
        orgao = home.line_orgao.text().upper()
        municipio = home.line_muni.text().upper()
        telefone = home.line_tel.text()
        email = home.line_email.text().lower()
 
        cursor.execute("UPDATE visitantes SET nome = '{}', orgao = '{}', municipio = '{}', telefone = '{}', email = '{}' WHERE id = '{}' ".format(nome, orgao, municipio, telefone, email, id) )    
        db.connection.commit()
        db.close_connection()
        QMessageBox.about(home, ' Mensagem ', ' Cadastro Atualizado com Sucesso! ')
 
        registrar(fullDataSet)
       
        quetion2.nome_new.setText('')
        quetion2.orgao_new.setText('')
        quetion2.muni_new.setText('')
        quetion2.tel_new.setText('')
        quetion2.email_new.setText('')
 
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
        pass
 
def sim_question(fullDataSet):
    try:
        db = Data_base()
        db.connect()
           
        input_cpf = home.line_CPF.text()
        input_nome = home.line_nome.text().upper()
        input_orgao = home.line_orgao.text().upper()
        input_municipio = home.line_muni.text().upper()
        input_telfone = home.line_tel.text()
        input_email = home.line_email.text().lower()
           
        fullDataSet = ( str(input_cpf),  str(input_nome), str(input_orgao), str(input_municipio), str(input_telfone), str(input_email))
       
        #CADASTRAR NO BANCO DE DADOS
        db.insert_table(fullDataSet)
        quetion.close()
        QMessageBox.about(home, ' Mensagem ', ' Registro realizado com sucesso! ')
        db.close_connection()
        home.line_CPF.setText('')
        home.line_nome.setText('')
        home.line_orgao.setText('')
        home.line_muni.setText('')
        home.line_tel.setText('')
        home.line_email.setText('')
        home.line_data.setText('')
        home.line_destino.setText('')
       
        data_hoje = datetime.now()
        home.line_data.setText(str(data_hoje.strftime("%d/%m/%Y")))
        #db.connection.commit()
        #db.close_connection()
    except:
        QMessageBox.warning(home, ' Atenção! ', ' Algo não saiu como planejado! ')
        db.close_connection()
 
def no_question():
    try:
        quetion.close()
        db.close_connection()
        QMessageBox.about(home, ' Mensagem ', ' Registro realizado com sucesso! ')
        home.line_CPF.setText('')
        home.line_nome.setText('')
        home.line_orgao.setText('')
        home.line_muni.setText('')
        home.line_tel.setText('')
        home.line_email.setText('')
        home.line_data.setText('')
        home.line_destino.setText('')
 
        data_hoje = datetime.now()
        home.line_data.setText(str(data_hoje.strftime("%d/%m/%Y")))
    except:
        QMessageBox.warning(home, ' Atenção! ', ' Algo não saiu como planejado! ')
       
def pdf_all():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        result = db.select_all()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_cpf():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE cpf LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_nome():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE nome LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_orgao():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE orgao LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_municipio():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE municipio LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_telefone():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE telefone LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_email():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE email LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_data():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE data LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def pdf_destino():
    try:
        webbrowser.open("Relatório.pdf")
        file_name = 'Relatório.pdf'
        document_title = 'Controle de Funcionários'
        title = ' RECEPÇÃO SECRETARIA DA MULHER '
        day = datetime.now().strftime("%d/%m/%Y")
        fecha_actual = f'Data de emissão: {day}'
       
        canvas = Canvas(file_name, pagesize=landscape(A4))
       
        doc = BaseDocTemplate(file_name)
        contents =[]
        width,height = landscape(A4)
       
        frame_later = Frame(
            0.11*inch,
            0.6*inch,
            (width-0.5*inch)+0.38*inch,
            height-1*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
       
        frame_table= Frame(
            0.1*inch,
            0.7*inch,
            (width-0.5*inch)+0.38*inch,
            height-2*inch,
            leftPadding = 0,
            topPadding=0,
            id='col'
            )
 
        laterpages = PageTemplate(id='laterpages',frames=[frame_later],pagesize=landscape(A4))
        firstpage = PageTemplate(id='firstpage',frames=[frame_later,frame_table],pagesize=landscape(A4))
       
        styleSheet = getSampleStyleSheet()
        style_title = styleSheet['Heading1']
        style_title.fontSize = 18
        style_title.fontName = 'Helvetica-Bold'
        style_title.alignment=TA_CENTER
       
        style_data = styleSheet['Normal']
        style_data.fontSize = 16
        style_data.fontName = 'Helvetica'
        style_data.alignment=TA_CENTER
       
        style_date = styleSheet['Normal']
        style_date.fontSize = 14
        style_date.fontName = 'Helvetica'
        style_date.alignment=TA_CENTER
       
        canvas.setTitle(document_title)
       
        contents.append(Paragraph(title, style_title))
        contents.append(Paragraph(fecha_actual, style_date))
        contents.append(FrameBreak())
       
        db = Data_base()
        db.connect()
        cursor = db.connection.cursor()
        sua_busca = home.buscar.text().upper()
        cursor.execute("SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros WHERE destino LIKE '%{}%'".format(sua_busca))
        result = cursor.fetchall()
 
        column1Heading = "CPF"
        column2Heading = "NOME"
        column3Heading = "ORGÃO/EMPRESA"
        column4Heading = "MUNICÍPIO"
        column5Heading = "TELEFONE"
        column6Heading = "EMAIL"
        column7Heading = "DATA"
        column8Heading = "DESTINO"
 
        data = [(column1Heading,column2Heading,column3Heading,column4Heading,column5Heading,column6Heading,column7Heading,column8Heading)]
        for i in range(0, len(result)):
            data.append([str(result[i][0]),Paragraph(str(result[i][1])),Paragraph(str(result[i][2])),Paragraph(str(result[i][3])),str(result[i][4]),Paragraph(str(result[i][5])),str(result[i][6]),Paragraph(str(result[i][7]))])
               
        tableThatSplitsOverPages = Table(data, colWidths=(1.2*inch, 2.2*inch, None, 1.4*inch, None, 1.9*inch, 0.9*inch, None))
        tableThatSplitsOverPages.hAlign = 'CENTER'
        tblStyle = TableStyle([('TEXTCOLOR',(0,0),(-1,-1),colors.black),
                            ('VALIGN',(0,0),(-1,-1),'TOP'),
                            ('LINEBELOW',(0,0),(-1,-1),1,colors.black),
                            ('GRID',(0,0),(-1,-1),1,colors.black)])
        tblStyle.add('BACKGROUND',(0,0),(7,0),colors.lightblue)
        tblStyle.add('BACKGROUND',(0,1),(-1,-1),colors.white)
        tableThatSplitsOverPages.setStyle(tblStyle)
        contents.append(tableThatSplitsOverPages)  
        contents.append(NextPageTemplate('laterpages'))
 
        contents.append(PageBreak())
 
        doc.addPageTemplates([firstpage,laterpages])
        doc.build(contents)
        db.close_connection()
    except:
        QMessageBox.critical(home, ' Atenção! ', ' Algo não saiu como planejado! ')
 
def index_changed(i):
    index = i
   
    if index == 0:
        home.localizar_cpf.setText('')
        home.stackedWidget_3.setCurrentWidget(home.page_19)
    elif index == 4:
        home.localizar_tel.setText('')
        home.stackedWidget_3.setCurrentWidget(home.page_21)
    elif index == 5:
        home.localizar_email.setText('')
        home.stackedWidget_3.setCurrentWidget(home.page_23)
    elif index == 6:
        home.localizar_data.setText('')
        home.stackedWidget_3.setCurrentWidget(home.page_22)
    else:
        home.stackedWidget_3.setCurrentWidget(home.page_20)
    pass
 
def excel_registro():
    try:
        path_name  = filedialog.askdirectory()
       
        cnx = sqlite3.connect("system.db")
        result = pd.read_sql_query("""SELECT cpf, nome, orgao, municipio, telefone, email, data, destino FROM registros""", cnx)
 
        result.to_excel(f"{path_name}/registros.xlsx", sheet_name='controle', index=False)
       
        QMessageBox.about(home, ' Mensagem ', '\n  Planilha gerada com Sucesso! \n ')
    except:
        QMessageBox.critical(home, ' Atenção! ', '\n A planilha não pode ser gerada! \n ')
 
def excel_pessoas():
    try:
        path_name  = filedialog.askdirectory()
       
        cnx = sqlite3.connect("system.db")
        result = pd.read_sql_query("""SELECT cpf, nome, orgao, municipio, telefone, email FROM visitantes ORDER BY nome""", cnx)
 
        result.to_excel(f"{path_name}/pessoas_cadastradas.xlsx", sheet_name='controle', index=False)
       
        QMessageBox.about(home, ' Mensagem ', ' Planilha gerada com Sucesso! \n ')
    except:
        QMessageBox.critical(home, ' Atenção! ', ' A planilha não pode ser gerada! \n ')

app = QtWidgets.QApplication([])

# Janelas
home = uic.loadUi('open.ui')
quetion = uic.loadUi('quetion.ui')
quetion2 = uic.loadUi('quetion_2.ui')

# BOTÕES - FUNÇÕES
home.btn_registar.clicked.connect(botao_registrar)
home.btn_historico.clicked.connect(botao_historico)
home.btn_cadastro.clicked.connect(botao_novo)
home.btn_home.clicked.connect(botao_home)
home.btn_help.clicked.connect(botao_help)

# ----- HELP PAGE -----
home.excel_registro.clicked.connect(excel_registro)
home.excel_pessoas.clicked.connect(excel_pessoas)
home.backup_lock.clicked.connect(backup_lock)
home.entrar.clicked.connect(login)
home.close.clicked.connect(botao_help)

# -----------------
home.backup_registro.clicked.connect(backup_registros)
home.backup_pessoa.clicked.connect(backup_pessoas)

# ---- NOVO ----
home.brn_salvar_3.clicked.connect(cadastrar_novo)

# ----- INPUT MASK -----
home.line_CPF_2.setValidator(QIntValidator())
home.line_CPF_2.setInputMask('999.999.999-99')
home.line_CPF_2.mouseReleaseEvent=lambda event:home.line_CPF_2.setCursorPosition(0)
home.line_tel_2.setInputMask("(99)\\  99990-9999")
home.line_tel_2.mouseReleaseEvent=lambda event:home.line_tel_2.setCursorPosition(0)
   
home.line_CPF.setValidator(QIntValidator())
home.line_CPF.setInputMask('999.999.999-99')
home.line_CPF.mouseReleaseEvent=lambda event:home.line_CPF.setCursorPosition(0)
home.line_tel.setInputMask("(99)\\  99990-9999")
home.line_tel.mouseReleaseEvent=lambda event:home.line_tel.setCursorPosition(0)
home.line_data.setValidator(QIntValidator()) 
home.line_data.setInputMask('99/99/9999')
home.line_data.mouseReleaseEvent=lambda event:home.line_data.setCursorPosition(0)

#  ------ MASK FILTRO ------
#  CPF
home.localizar_cpf.setValidator(QIntValidator()) 
home.localizar_cpf.setInputMask('999.999.999-99')
home.localizar_cpf.mouseReleaseEvent=lambda event:home.localizar_cpf.setCursorPosition(0)

# TELEFONE
home.localizar_tel.setValidator(QIntValidator()) 
home.localizar_tel.setInputMask("(99)\\  99990-9999")
home.localizar_tel.mouseReleaseEvent=lambda event:home.localizar_tel.setCursorPosition(0)

# DATA
home.localizar_data.setValidator(QIntValidator()) 
home.localizar_data.setInputMask('99/99/9999')
home.localizar_data.mouseReleaseEvent=lambda event:home.localizar_data.setCursorPosition(0)

   
home.line_data.setValidator(QIntValidator())
# ---- PAGE REGISTRO ----
home.btn_complete.clicked.connect(completar)
home.brn_salvar.clicked.connect(validar_edit)

# ---- PAGE HISTÓRICO ----
home.all.clicked.connect(Consulta_all)
home.busca.clicked.connect(Consulta_filtro)
home.filtro.currentIndexChanged.connect(index_changed)  

home.relatorio_4.clicked.connect(pdf_all)
home.relatorio_5.clicked.connect(pdf_cpf)
home.relatorio_6.clicked.connect(pdf_nome)
home.relatorio_8.clicked.connect(pdf_orgao)
home.relatorio_9.clicked.connect(pdf_municipio)
home.relatorio_10.clicked.connect(pdf_telefone)
home.relatorio_11.clicked.connect(pdf_email)
home.relatorio_12.clicked.connect(pdf_data)
home.relatorio_13.clicked.connect(pdf_destino)

# ---- PAGE QUESTION ----
quetion.brn_salvar.clicked.connect(sim_question)
quetion.btn_complete.clicked.connect(no_question)

# ---- PAGE QUESTION 2 ----
quetion2.nao_nao.clicked.connect(nao_nao)
quetion2.nao_sim.clicked.connect(nao_sim)
quetion2.sim_sim.clicked.connect(sim_sim)

db = Data_base()
db.connect()
db.create_table()
db.create_table_2()
db.close_connection()

home.stackedWidget.setCurrentWidget(home.page)
home.showMaximized()
app.exec()
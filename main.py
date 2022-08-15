# Script de automação de inclusão de dados de XML de notas fiscais emitidas em planilha.
import os
import xml.etree.ElementTree as ET
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = 'Emitidas'
ws.append(['Número da NF', 'Data de emissão', 'Natureza da Operação', 'Destinatário', 'CNPJ', 'Valor'])
emitidas = './Emitidas'
for arquivos in os.walk(emitidas):
    for arquivo in arquivos[2]:
        if arquivo != '.DS_Store':
            root = ET.parse(f'{emitidas}/{arquivo}').getroot()
            nsNFE = {'ns': "http://www.portalfiscal.inf.br/nfe"}
            try:
                numero_nfe = (root.find('ns:NFe/ns:infNFe/ns:ide/ns:nNF', nsNFE)).text
                natOp = (root.find('ns:NFe/ns:infNFe/ns:ide/ns:natOp', nsNFE)).text
                dhEmi = (((root.find('ns:NFe/ns:infNFe/ns:ide/ns:dhEmi', nsNFE)).text)).replace('-', '/')[0:10]
                dest_nome = (root.find('ns:NFe/ns:infNFe/ns:dest/ns:xNome', nsNFE)).text
                if root.find('ns:NFe/ns:infNFe/ns:dest/ns:CPF', nsNFE) != None:
                    dest_cpf = (root.find('ns:NFe/ns:infNFe/ns:dest/ns:CPF', nsNFE)).text
                else:
                    dest_cpf = (root.find('ns:NFe/ns:infNFe/ns:dest/ns:CNPJ', nsNFE)).text
                valor = (root.find('ns:NFe/ns:infNFe/ns:total/ns:ICMSTot/ns:vNF', nsNFE)).text
                dados = []
                dados = [numero_nfe, dhEmi, natOp, dest_nome, dest_cpf, valor]
                ws.append(dados)
            except AttributeError:
                pass
            except xml.etree.ElementTree.ParseError:
                pass
wb.create_sheet('Recebidas')
ws2 = wb['Recebidas']
recebidas = './Recebidas'
ws2.append(['Número da NF', 'Data de emissão', 'Natureza da Operação', 'Emitente', 'CNPJ', 'Valor'])
for arquivos2 in os.walk(recebidas):
    for arquivo2 in arquivos2[2]:
        if arquivo2 != '.DS_Store':
            root = ET.parse(f'{recebidas}/{arquivo2}').getroot()
            nsNFE = {'ns': "http://www.portalfiscal.inf.br/nfe"}
            try:
                numero_nfe2 = (root.find('ns:NFe/ns:infNFe/ns:ide/ns:nNF', nsNFE)).text
                natOp2 = (root.find('ns:NFe/ns:infNFe/ns:ide/ns:natOp', nsNFE)).text
                dhEmi2 = (((root.find('ns:NFe/ns:infNFe/ns:ide/ns:dhEmi', nsNFE)).text)).replace('-', '/')[0:10]
                rem_nome = (root.find('ns:NFe/ns:infNFe/ns:emit/ns:xNome', nsNFE)).text
                rem_cnpj = (root.find('ns:NFe/ns:infNFe/ns:emit/ns:CNPJ', nsNFE)).text
                valor2 = (root.find('ns:NFe/ns:infNFe/ns:total/ns:ICMSTot/ns:vNF', nsNFE)).text
                dados2 = [numero_nfe2, dhEmi2, natOp2, rem_nome, rem_cnpj, valor2]
                ws2.append(dados2)
            except AttributeError:
                pass
            except xml.etree.ElementTree.ParseError:
                pass
filename_input = input('Digite o mês e o ano (Ex.: Jun22): ')
wb.save(f'./Fechamento {filename_input}.xlsx')

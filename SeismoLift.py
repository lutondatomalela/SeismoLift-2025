# -*- coding: utf-8 -*-
"""
Created on Sat Jun 14 23:45:15 2025

@author: luton
"""

import pandas as pd
from unidecode import unidecode
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx2pdf import convert
from docx.shared import Pt, Inches
from datetime import datetime
from docx.oxml import parse_xml
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import nsdecls
from docx.oxml import OxmlElement
import os

def menu():
    while True:
        print("\n==================== SeismoLift 2025 ====================\n")
        print("1 - Iniciar Análise")
        print("2 - Sair")
        opcao = input("\nEscolha uma opção (1 ou 2): ").strip()
        if opcao in ("1", "2"):
            return opcao
        print("\nEntrada inválida. Por favor, pressione 1 para iniciar ou 2 para sair.\n")

def Copyright():
    print("\n==================== SeismoLift 2025 ====================")
    try:
        with open("LAT.txt", "r") as file:
            assinatura = file.read()
        print(assinatura)
    except FileNotFoundError:
        print("\n[Erro] O ficheiro de direitos autorais não foi encontrado.\n")
    except Exception as e:
        print(f"\n[Erro] Ocorreu um erro ao ler o ficheiro:\n {e}\n")

def SeismoLift():
    # Caminhos relativos
    base_dir = os.path.dirname(__file__)
    ficheiro_excel = os.path.join(base_dir, "../1_IN/Zonas_Sismicas_PT.xlsx")
    localidades = ["Portugal Continental", "Arquipélago da Madeira", "Arquipélago dos Açores"]

    def normalizar(texto):
        return unidecode(str(texto)).lower().strip()

    def pedir_localidade():
        while True:
            print("\nInsira o nome do Concelho (ex.: 'Faro', 'Porto', 'Ponta Delgada').")
            local = input("Concelho (ou 'sair' para terminar): ").strip()
            if local.lower() == "sair":
                print("\nPrograma terminado.")
                return None
            return local

    def zona_sismica(localidade):
        localidade_normalizado = normalizar(localidade)
        for local in localidades:
            try:
                df = pd.read_excel(ficheiro_excel, sheet_name=local, header=0)
            except Exception as e:
                print(f"\n[Erro] Não foi possível ler a folha '{local}': {e}")
                continue

            col_b = df.iloc[1:, 1]  # designação (coluna B, 2.a linha em diante)
            for idx, concelho in col_b.items():
                if pd.isna(concelho):
                    continue
                if localidade_normalizado == normalizar(concelho):
                    print(f"\n {concelho}, {local}")

                    if local == "Portugal Continental":
                        aceleracoes = df.iloc[idx, 2:6].tolist()
                        z_1 = aceleracoes[0]; z_2 = aceleracoes[2]
                        agR_1 = float(aceleracoes[1]); agR_2 = float(aceleracoes[3])
                        return agR_1, agR_2, localidade, concelho, local, z_1, z_2

                    elif local == "Arquipélago da Madeira":
                        aceleracoes = df.iloc[idx, 2:4].tolist()
                        z_1 = float(aceleracoes[0]); z_2 = None
                        agR_1 = float(aceleracoes[-1]); agR_2 = None
                        return agR_1, agR_2, localidade, concelho, local, z_1, z_2

                    elif local == "Arquipélago dos Açores":
                        aceleracoes = df.iloc[idx, 2:4].tolist()
                        z_1 = None; z_2 = float(aceleracoes[0])
                        agR_1 = None; agR_2 = float(aceleracoes[-1])
                        return agR_1, agR_2, localidade, concelho, local, z_1, z_2

        return None, None, None, None, None, None, None

    while True:
        localidade = pedir_localidade()
        if not localidade:
            return
        agR_1, agR_2, localidade, concelho, local, z_1, z_2 = zona_sismica(localidade)
        if agR_1 is None and agR_2 is None:
            print("⚠️ Concelho não encontrado, tente novamente.")
        else:
            break

    def classe_importancia():
        descricoes_classe = {
            "I":  "Edifícios de importância menor.",
            "II": "Edifícios correntes.",
            "III":"Edifícios importantes (escolas, etc).",
            "IV": "Edifícios críticos (hospitais, etc)."
        }
        while True:
            print("\nClasses de importância a considerar:")
            for k, v in descricoes_classe.items():
                print(f"{k} - {v}")
            classe = input("\nInsira a classe (I, II, III ou IV): ").strip().upper()
            if classe in descricoes_classe:
                return classe
            print("\n⚠️ Entrada inválida. Tente novamente.")

    classe = classe_importancia()

    def coeficientes_importancia(classe):
        if classe == 'I':
            gamma_l = 0.80; gamma_a = 1.00
        elif classe == 'II':
            gamma_l = 1.00; gamma_a = 1.00
        elif classe == 'III':
            gamma_l = 1.20; gamma_a = 1.00
        else:
            gamma_l = 1.40; gamma_a = 1.50
        return gamma_l, gamma_a

    gamma_l, gamma_a = coeficientes_importancia(classe)

    def aceleracao_superficie(gamma_l, agR_1, agR_2, local):
        # Correção normativa: Continente/Madeira → Tipo 1 ; Açores → Tipo 2
        if agR_1 is None and agR_2 is None:
            agR = None
        elif local == "Arquipélago dos Açores":
            agR = agR_2
        else:
            agR = agR_1
        ag = gamma_l * agR
        alfa = ag / 9.81
        return ag, alfa, agR

    ag, alfa, agR = aceleracao_superficie(gamma_l, agR_1, agR_2, local)

    def tipo_terreno():
        descricoes_terreno = {
            "A": "Rocha ou similar.",
            "B": "Areia muito compacta, cascalho ou argila rija.",
            "C": "Areia compacta ou média, argila pouco coesa.",
            "D": "Solos soltos ou pouco compactos.",
            "E": "Perfil aluvionar superficial fraco sobre solo rijo."
        }
        while True:
            print("\nTipos de solo a considerar:")
            for k, v in descricoes_terreno.items():
                print(f"{k} - {v}")
            terreno = input("\nInsira o tipo de solo (A-E): ").strip().upper()
            if terreno in descricoes_terreno:
                return terreno
            print("\n⚠️ Entrada inválida.")

    terreno = tipo_terreno()

    def coeficiente_solo(terreno):
        if terreno == 'A':
            S = 1.00
        elif terreno == 'B':
            S = 1.35
        elif terreno == 'C':
            S = 1.60
        elif terreno == 'D':
            S = 2.00
        elif terreno == 'E':
            S = 1.80
        else:
            S = 1.60
        return S

    S = coeficiente_solo(terreno)

    def tipo_estrutura():
        descricoes_estrutura = {
            "1": "Pórticos metálicos.",
            "2": "Pórticos de betão ou contraventados.",
            "3": "Estruturas em geral.",
            "":  ""
        }
        while True:
            print("\nTipos de estrutura a considerar:")
            for k, v in descricoes_estrutura.items():
                print(f"{k} - {v}")
            estrutura = input("\nInsira o tipo (1, 2, 3 ou 'Enter' para ignorar.): ").strip()
            if estrutura in descricoes_estrutura:
                return estrutura
            print("\n⚠️ Entrada inválida.")

    estrutura = tipo_estrutura()

    def coeficiente_forma(estrutura):
        if estrutura == '1':
            Ct = 0.085
        elif estrutura == '2':
            Ct = 0.075
        elif estrutura == '3':
            Ct = 0.050
        else:
            Ct = 0.050
        qa = 2.0
        return Ct, qa

    Ct, qa = coeficiente_forma(estrutura)

    def input_altura(msg):
        while True:
            print(msg)
            entrada = input("\nAltura (em metros): ").strip().replace(",", ".")
            try:
                altura = float(entrada)
                if altura > 0:
                    return altura
                print("\n⚠️ Valor deve ser positivo.")
            except ValueError:
                print("\n⚠️ Entrada inválida.")

    def alturas():
        H = input_altura("\nNota 1: Altura total do edifício (fundação/topo de cave rígida → cobertura).")
        while True:
            z = input_altura("\nNota 2: Altura do elemento não estrutural acima do nível de aplicação da ação sísmica (z ≤ H).")
            if z <= H:
                break
            print("⚠️ z não pode ser superior a H. Tente novamente.")
        return H, z

    H, z = alturas()

    def periodo_vibracao_fundamental(Ct, H):
        T1 = Ct * H ** (3/4)
        Ta = 0.00
        return T1, Ta

    T1, Ta = periodo_vibracao_fundamental(Ct, H)

    def calcular_categoria_sismica(qa, Ta, T1, H, z):
        Sa = alfa * S * (3 * (1 + z/H) / (1 + (1 - Ta / T1)**2) - 0.5)
        Sa = max(Sa, alfa * S)  # condição Sa ≥ α·S (consistente com EN 81-77 Remissão EC8)
        ad = Sa * (gamma_a / qa) * 9.81
        if ad <= 1:
            categoria = 0
        elif ad <= 2.5:
            categoria = 1
        elif ad <= 4:
            categoria = 2
        else:
            categoria = 3
        return Sa, ad, categoria

    Sa, ad, categoria = calcular_categoria_sismica(qa, Ta, T1, H, z)

    def resumo_resultados(categoria):
        print(f"\n Resultados:")
        print(f"Local: {concelho}, {local}")
        print(f"γI     [-]:          {gamma_l:.3f}".replace('.', ',')) # Coeficiente de importância (EC8: 4.2.5(5)P. NOTA)
        print(f"agR    [m/s²]:       {agR:.3f}".replace('.', ',')) # Ação sísmica relevante
        print(f"ag     [m/s²]:       {ag:.3f}".replace('.', ',')) # Valor de cálculo da aceleração à superfície de um terreno do tipo A
        print(f"S      [-]:          {S:.3f}".replace('.', ',')) # Coeficiente de solo (Quadro 3.2/3.3 EC8)
        print(f"Ct     [-]:          {Ct:.3f}".replace('.', ',')) # Coeficiente de forma .EC8 - 4.3.3.2.2 (4.6)
        print(f"H      [m]:          {H:.3f}".replace('.', ',')) # Altura do edifício desde a fundação ou desde o nível superior de uma cave rígida
        print(f"z      [m]:          {z:.3f}".replace('.', ',')) # Altura do elemento não estrutural acima do nível de aplicação da ação sísmica (fundação ou nível superior de uma cave rígida
        print(f"α      [-]:          {alfa:.3f}".replace('.', ',')) # Quociente entre a aceleração para solos tipo A e a aceleração gravítica
        print(f"T1     [s]:          {T1:.3f}".replace('.', ',')) # Período de vibração fundamental do edifício na direção considerada
        print(f"qa     [-]:          {qa}".replace('.', ',')) # Coeficientes de comportamento (EC8 - 4.3.5.4: Quadro 4.4)
        print(f"γa     [-]:          {gamma_a}".replace('.', ',')) # Coeficientes de importância (EC8 - 4.3.5.3)
        print(f"Sa     [-]:          {Sa:.3f}".replace('.', ',')) # Coeficiente sísmico aplicável aos elementos não estruturais (EC8 - 4.3.5.2 (4.25))
        print(f"ad     [m/s²]:       {ad:.3f}".replace('.', ',')) # Aceleração de dimensionamento (EN 81-77)
        print(f"\nCategoria sísmica: {categoria}")

    resumo_resultados(categoria)

    def relatorio_de_calculo(classe, estrutura, terreno, caminho_saida="SeismoLift_report.docx"):
        doc = Document()

        style = doc.styles['Normal']
        style.font.name = 'Courier New'
        style.font.size = Pt(10)

        custom_style = doc.styles.add_style('CustomStyle', 1)
        sub_custom_style = doc.styles.add_style('sub_CustomStyle', 1)
        custom_style.font.name = 'Courier New'
        custom_style.font.size = Pt(16)
        sub_custom_style.font.name = 'Courier New'
        sub_custom_style.font.size = Pt(12)

        title = doc.add_heading('Relatório - Categoria sísmica do elevador', level=0)
        title.style = custom_style
        for run in title.runs:
            run.bold = True
        title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        now = datetime.now()
        formatted_time = now.strftime("%Y-%m-%d %H:%M:%S")
        section = doc.sections[0]
        footer = section.footer
        footer_para = footer.paragraphs[0]
        run = footer_para.add_run(f"SeismoLift 2025 - Relatório automático - Gerado em: {formatted_time}")
        run.font.size = Pt(8)
        footer_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        doc.add_paragraph('')
        titulo_1 = doc.add_heading('1. Localização e zonamento sísmico', level=1)
        titulo_1.style = sub_custom_style
        for run in titulo_1.runs:
            run.bold = True
        doc.add_paragraph('')
        doc.add_paragraph(f'Local: {concelho}, {local}')

        if z_1 == None:
            doc.add_paragraph(f'Zona sísmica tipo 2 (NP EN 1998-1 2009: ANEXO NA.I): {z_2}')
            doc.add_paragraph(f'Ação sísmica tipo 2 (NP EN 1998-1 2009: ANEXO NA.I) (agR): {f"{agR_2:.3f}".replace(".", ",")} m/s²')
        elif z_2 == None:
            doc.add_paragraph(f'Zona sísmica tipo 1 (NP EN 1998-1 2009: ANEXO NA.I): {z_1}')
            doc.add_paragraph(f'Ação sísmica tipo 1 (NP EN 1998-1 2009: ANEXO NA.I) (agR): {f"{agR_1:.3f}".replace(".", ",")} m/s²')
        else:
            doc.add_paragraph(f'Zona sísmica tipo 1 (NP EN 1998-1 2009: ANEXO NA.I): {z_1}')
            doc.add_paragraph(f'Ação sísmica tipo 1 (NP EN 1998-1 2009: ANEXO NA.I) (agR): {f"{agR_1:.3f}".replace(".", ",")} m/s²')
            doc.add_paragraph(f'Zona sísmica tipo 2 (NP EN 1998-1 2009: ANEXO NA.I): {z_2}')
            doc.add_paragraph(f'Ação sísmica tipo 2 (NP EN 1998-1 2009: ANEXO NA.I) (agR): {f"{agR_2:.3f}".replace(".", ",")} m/s²')

        doc.add_paragraph(f'Espectro utilizado para o cálculo: {"Tipo 2" if local=="Arquipélago dos Açores" else "Tipo 1"}')
        doc.add_paragraph(f'Valor de cálculo da aceleração à superfície do terreno (NP EN 1998-1 2009: 3.2.1 (3)) — ag = γI·agR: {f"{ag:.3f}".replace(".", ",")} [m/s²]')

        doc.add_paragraph('')
        titulo_2 = doc.add_heading('2. Parâmetros de cálculo', level=1)
        titulo_2.style = sub_custom_style
        for run in titulo_2.runs:
            run.bold = True
        doc.add_paragraph('')

        doc.add_paragraph('Classe e coeficientes de importância (NP EN 1998-1 2009: 4.2.5; NA-4.2.5(5)P. Ver Nota 4.3.5.3): '
                          + f'{classe} - Edifícios cuja integridade em caso de sismo é de importância vital (hospitais, quartéis de bombeiros, centrais eléctricas, etc).')
        doc.add_paragraph(f'(γI): {f"{gamma_l:.3f}".replace(".", ",")} [-]')
        doc.add_paragraph(f'(γa): {f"{gamma_a:.3f}".replace(".", ",")} [-]')

        doc.add_paragraph(f'Valor de cálculo da aceleração à superfície do terreno (NP EN 1998-1 2009: 3.2.1 (3)) (ag): {f"{ag:.3f}".replace(".", ",")} [m/s²]')

        doc.add_paragraph(f'Tipo de terreno (NP EN 1998-1 2009: Quadro 3.1): {terreno}')
        doc.add_paragraph(f'Coeficiente de solo (NP EN 1998-1 2009: Quadro 3.2/3.3 EC8) (S): {f"{S:.3f}".replace(".", ",")} [-]')

        if estrutura == '1':
            doc.add_paragraph('Tipo de estrutura e coeficiente de forma (NP EN 1998-1 2009: 4.3.3.2.2 (4.6)): Pórticos espaciais metálicos.')
        elif estrutura == '2':
            doc.add_paragraph('Tipo de estrutura e coeficiente de forma (NP EN 1998-1 2009: 4.3.3.2.2 (4.6)): Pórticos espaciais de betão e/ou pórticos metálicos com contraventamentos excêntricos.')
        else:
            doc.add_paragraph('Tipo de estrutura e coeficiente de forma (NP EN 1998-1 2009: 4.3.3.2.2 (4.6)): Estruturas em geral.')

        doc.add_paragraph(f'(Ct): {f"{Ct:.3f}".replace(".", ",")} [-]')
        doc.add_paragraph(f'Altura total do edifício (H): {f"{H:.3f}".replace(".", ",")} m')
        doc.add_paragraph(f'Altura do elemento considerado (z): {f"{z:.3f}".replace(".", ",")} m')
        doc.add_paragraph(f'Período fundamental de vibração do elemento não estrutural (EN 81-77:2022 - Annex B) (Ta): {f"{Ta:.3f}".replace(".", ",")} [s]')
        doc.add_paragraph(f'Período fundamental de vibração do edifício na direcção relevante (EN 81-77:2022 - Annex B) (T1): {f"{T1:.3f}".replace(".", ",")} [s]')

        doc.add_paragraph('')
        titulo_3 = doc.add_heading('3. Resultados', level=1)
        titulo_3.style = sub_custom_style
        for run in titulo_3.runs:
            run.bold = True
        doc.add_paragraph('')

        tabela = doc.add_table(rows=2, cols=3)
        tabela.alignment = WD_TABLE_ALIGNMENT.CENTER

        hdr_cells = tabela.rows[0].cells
        hdr_cells[0].text = 'Aceleração de projeto (ad) [m/s²]                        (EN 81-77:2022 (B.1))'
        hdr_cells[1].text = 'Categoria sísmica do elevador                              (EN 81-77:2022 (B.2))'
        hdr_cells[2].text = 'Nota'

        row_cells = tabela.rows[1].cells
        if categoria == 0:
            row_cells[0].text = f'{ad:.3f} < 1'.replace(".", ",")
            row_cells[1].text = '0'
            row_cells[2].text = 'Não são requeridas ações adicionais.'
        elif categoria == 1:
            row_cells[0].text = f'1 ≤ {ad:.3f} < 2,5'.replace(".", ",")
            row_cells[1].text = '1'
            row_cells[2].text = 'Ações corretivas de baixa expressão.'
        elif categoria == 2:
            row_cells[0].text = f'2,5 ≤ {ad:.3f} < 4'.replace(".", ",")
            row_cells[1].text = '2'
            row_cells[2].text = 'Ações corretivas de média expressão.'
        else:
            row_cells[0].text = f'{ad:.3f} ≥ 4'.replace(".", ",")
            row_cells[1].text = '3'
            row_cells[2].text = 'São requeridas ações corretivas importantes.'

        # Cor da célula da categoria
        cores_categoria = {0: "90EE90", 1: "FFFF00", 2: "FFA500", 3: "FF6B6B"}
        cor_categoria = cores_categoria.get(categoria, "FFFFFF")
        shading_categoria = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), cor_categoria))
        row_cells[1]._tc.get_or_add_tcPr().append(shading_categoria)

        for row in tabela.rows:
            for cell in row.cells:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                for paragraph in cell.paragraphs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Bordas da tabela: correção 
        tbl_pr = tabela._tbl.tblPr
        if tbl_pr is None:
            tbl_pr = OxmlElement('w:tblPr')
            tabela._tbl.append(tbl_pr)

        tbl_borders = parse_xml(r'''
        <w:tblBorders %s>
            <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideH w:val="single" w:sz="6" w:space="0" w:color="000000"/>
            <w:insideV w:val="single" w:sz="6" w:space="0" w:color="000000"/>
        </w:tblBorders>
        ''' % nsdecls('w'))
        tbl_pr.append(tbl_borders)

        try:
            out_path = os.path.join(base_dir, "../2_OUT/SeismoLift_report.docx")
            doc.save(out_path)
            #print(f"\n📄 Relatório gerado: {out_path}")
        except Exception as e:
            print(f"\n⚠️ Não foi possível gerar o relatório .docx.\nErro: {e}")

    relatorio_de_calculo(classe, estrutura, terreno)

if __name__ == "__main__":
    opcao = menu()
    if opcao == "1":
        SeismoLift()
        Copyright()
    else:
        print("\nPrograma encerrado.\n")
        Copyright()

import pandas as pd
from datetime import time
import pandas as pd
from openpyxl import load_workbook
from descricao_tra import define_descricao
from xlsxwriter import Workbook

estados_regioes = {
    'AC': 'Norte',
    'AL': 'Nordeste',
    'AP': 'Norte',
    'AM': 'Norte',
    'BA': 'Nordeste',
    'CE': 'Nordeste',
    'DF': 'Centro-Oeste',
    'ES': 'Sudeste',
    'GO': 'Centro-Oeste',
    'MA': 'Nordeste',
    'MT': 'Centro-Oeste',
    'MS': 'Centro-Oeste',
    'MG': 'Sudeste',
    'PA': 'Norte',
    'PB': 'Nordeste',
    'PR': 'Sul',
    'PE': 'Nordeste',
    'PI': 'Nordeste',
    'RJ': 'Sudeste',
    'RN': 'Nordeste',
    'RS': 'Sul',
    'RO': 'Norte',
    'RR': 'Norte',
    'SC': 'Sul',
    'SP': 'Sudeste',
    'SE': 'Nordeste',
    'TO': 'Norte'
}

lider = {
    'DF': 'Daniel',
    'GO': 'Daniel',
    'MT': 'Bruno Martins',
    'MS': 'Bruno Martins',
    'AL': 'Luiz Bolzon',
    'BA': 'Luiz Bolzon',
    'CE': 'Luiz Bolzon',
    'MA': 'Erivaldo',
    'PB': 'Luiz Bolzon',
    'PE': 'Luiz Bolzon',
    'PI': 'Gilvan',
    'RN': 'Luiz Bolzon',
    'SE': 'Luiz Bolzon',
    'AC': 'Erivaldo',
    'AP': 'Erivaldo',
    'AM': 'Erivaldo',
    'PA': 'Gilvan',
    'RO': 'Erivaldo',
    'RR': 'Erivaldo',
    'TO': 'Gilvan',
    'ES': 'Dolôr',
    'MG': 'Dolôr',
    'RJ': 'Dolôr',
    'SP': 'Elesandro',
    'PR': 'Elesandro',
    'RS': 'Elesandro',
    'SC': 'Elesandro'
}

causa_dicionario = {
    'Compressor': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da evaporadora ': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da evaporadora': 'Infraestrutura - Sistema de Climatização',
    'FAN-10': 'Infraestrutura - Sistema de Climatização',
    'Módulo de expansão 2157': 'Infraestrutura - Sistema de Climatização',
    'Transformador 220 - 24 ': 'Infraestrutura - Sistema de Climatização',
    'AGST - MP5000 M00 08000004P': 'Infraestrutura - Sistema de Climatização',
    'Válvula de expansão': 'Infraestrutura - Sistema de Climatização',
    'Ventilador da condensadora ': 'Infraestrutura - Sistema de Climatização',
    ': ZIEHL - Modelo: ZEIHL ABEGG': 'Infraestrutura - Sistema de Climatização',
    'PLC  SANRIO': 'Infraestrutura - Sistema de Climatização',
    'Módulo AGST - MP5000 M02 00080800': 'Infraestrutura - Sistema de Climatização',
    'Placa de circuito interno do Ar-condicionado SANRIO': 'Infraestrutura - Sistema de Climatização',
    'Tubulação de gás para ar-condicionado.': 'Infraestrutura - Sistema de Climatização',
    'AGST - MP5000 M00 08000004P': 'Infraestrutura - Sistema de Climatização',
    '8 unidade  Bateria Sec Power HMA-12 110/12V 110Ah - SEC POWER': 'Infraestrutura - Sistema de Climatização',
    'Ziehl Abegg 113237 230vts // 60h // 15a/  40ac': 'Infraestrutura - Sistema de Climatização',
    '4 Ventiladores  Ziel  ': 'Infraestrutura - Sistema de Climatização',
    'Controladora PCIR-04': 'Infraestrutura - Sistema de Climatização',
    'Reposição de gás': 'Infraestrutura - Sistema de Climatização',
    'Controladora PCIR-04': 'Infraestrutura - Sistema de Climatização',
    'Falha na tubulação do AC4': 'Infraestrutura - Sistema de Climatização',
    'Controladora TIC-17 e MF6': 'Infraestrutura - Sistema de Climatização',
    'Controladora TCIR04': 'Infraestrutura - Sistema de Climatização',
    'Falha na tubulação do AC1': 'Infraestrutura - Sistema de Climatização',
    'PLC -  TCIR04': 'Infraestrutura - Sistema de Climatização',
    'SSC10 BBAA': 'Infraestrutura - Sistema de Climatização',
    'Máquina de ar-condicionado': 'Infraestrutura - Sistema de Climatização',
    'PLC SANRIO PCIR03G': 'Infraestrutura - Sistema de Climatização',
    'Tampa da caixa R1': 'Infraestrutura - Sistema Retificador FCC',
    'Kit Ventilado': 'Infraestrutura - Sistema Retificador FCC',
    'Placa OMPI': 'Infraestrutura - Sistema Retificador FCC',
    'RAU2X6U/A26R6AUKL40179/A26': 'Rádio - Equipamento Ericsson',
    'RAU2x7/A15Ericsson': 'Rádio - Equipamento Ericsson',
    'NPU 3C': 'Rádio - Equipamento Ericsson',
    'NPU3 C ROJR 211 006/2 R2A': 'Rádio - Equipamento Ericsson',
    'DVR/Câmeras/Fonte12V': 'Infraestrutura - Sistema de CFTV',
    'NVR-2 camera 22': 'Infraestrutura - Sistema de CFTV',
    'NVR': 'Infraestrutura - Sistema de CFTV',
    'KoDo PRO - Modelo: KCX-5700N': 'Infraestrutura - Sistema de CFTV',
    ': Isotrafo  - Modelo: 45KVA - 13,8 KV': 'Infraestrutura - Sistema de CFTV',
    'HIK VISION  - Modelo: DS-2CD2120F-IS': 'Infraestrutura - Sistema de CFTV',
    'MCE - Modelo: 300W': 'Infraestrutura - Sistema de CFTV',
    'Cabos para CFTV': 'Infraestrutura - Sistema de CFTV',
    'SPVL-4': 'DWDM - Equipamento PADTEC',
    'SPVL-4SM': 'DWDM - Equipamento PADTEC',
    "2 Supervisores SMARTPACWEB / SNMP6 UR's 50A FLATPAC21 SPVL-901 Base para DPS3 BANDEJAS para RETIFICADORES FCC ELTEK": 'DWDM - Equipamento PADTEC',
    'SPVL-90': 'DWDM - Equipamento PADTEC',
    'Subbastidor 14uTM400# sobressalente OCM  com defeito SPVL-91# com defeito SSC  sobressalente': 'DWDM - Equipamento PADTEC',
    'SPVL 90': 'DWDM - Equipamento PADTEC',
    'SPVL-91 ': 'DWDM - Equipamento PADTEC',
    'T100DCT-4JRT2L':'DWDM - Equipamento PADTEC',
    'T100DCT-4JRT2L	':'DWDM - Equipamento PADTEC',
    'TR400C93-QBF-QBF':'DWDM - Equipamento PADTEC',
    'T100DCT-4PTT2L': 'DWDM - Equipamento PADTEC',
    'SSC-BBAA - FAN 10':'DWDM - Equipamento PADTEC',
    'TM400C92QBFXHACA':'DWDM - Equipamento PADTEC',
    'TM400C92-DBF-XHF-CA.':'DWDM - Equipamento PADTEC',
    'TM400-9B':'DWDM - Equipamento PADTEC',
    'T100DCT-4JT2L':'DWDM - Equipamento PADTEC',
    'TM400C92-DBF-XHF-CA':'DWDM - Equipamento PADTEC',
    'TR400-9B':'DWDM - Equipamento PADTEC',
    'LOA4C211AYAHA':'DWDM - Equipamento PADTEC',
    'SCME-4DP e CVA-4SRA':'DWDM - Equipamento PADTEC',
    '- SPVL-4SM':'DWDM - Equipamento PADTEC',
    'FAN-G8':'DWDM - Equipamento PADTEC',
    'SCMD3S1A':'DWDM - Equipamento PADTEC',
    'T100DCT-4JTMYL':'DWDM - Equipamento PADTEC',
    '- Sobressalente necessário: Placa SSC-AAAA ':'DWDM - Equipamento PADTEC',
    'CVA-4SRA':'DWDM - Equipamento PADTEC',
    'Amplificador Óptico de Linha - LOAP14B244AA':'DWDM - Equipamento PADTEC',
    ': SCME-4DP':'DWDM - Equipamento PADTEC',
    'VOAB-2A16AA':'DWDM - Equipamento PADTEC',
    '- Modelo: VOAB-2A16AA':'DWDM - Equipamento PADTEC',
    'MDSADC21401ST3':'DWDM - Equipamento PADTEC',
    'BOA4C241BDAHA':'DWDM - Equipamento PADTEC',
    '3x XFP 10G Base-LR/LW 1310nm':'DWDM - Equipamento PADTEC',
    'LightPad i1600G - Canal de Voz - CVA-4SRA':'DWDM - Equipamento PADTEC',
    'POA4C141AHAH':'DWDM - Equipamento PADTEC',
    'PADTEC - LOAP14B244AA#268':'DWDM - Equipamento PADTEC',
    'TR400-9B#':'DWDM - Equipamento PADTEC',
    'CVA-4SSA':'DWDM - Equipamento PADTEC',
    'TCX11-4P-A1#':'DWDM - Equipamento PADTEC',
    'SFP 10 Gb  Modelo: 1200-SM-LL-L': 'IP - Equipamento Datacom',
    'DM4000 - MPU512 ': 'IP - Equipamento Datacom',
    '1KVA NB HDS LM S2': 'Infraestrutura - Nobreak',
    'QCAB': 'Infraestrutura - Balizamento de Torre',
    'Bomba injetora do GMG ': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria MARCA -  DISBAL ; MODELO - S 150MD': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria de GMG': 'Infraestrutura - Grupo Motor Gerador',
    'Gerador completo': 'Infraestrutura - Grupo Motor Gerador',
    'bateria do gerador  Optima - Modelo: Gel - Red Top 35 / 12 volts - 44ha - 720a (-18Cº) 910a 0Cº - RC90min': 'Infraestrutura - Grupo Motor Gerador',
    'DEEP SEA MODELO:DSE-7320': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria para o GMG': 'Infraestrutura - Grupo Motor Gerador',
    'USCA - AJX TELECOM - AP Control GMG 7320 8-36 VCC': 'Infraestrutura - Grupo Motor Gerador',
    'USCA CUNNINS': 'Infraestrutura - Grupo Motor Gerador',
    'Retificaro GMG': 'Infraestrutura - Grupo Motor Gerador',
    'Contactora Stemac EK370': 'Infraestrutura - Grupo Motor Gerador',
    'Bateria ': 'Infraestrutura - Grupo Motor Gerador',
    'Estemac MWMD229': 'Infraestrutura - Grupo Motor Gerador',
    'Fonte da Telemetria Siemens': 'Infraestrutura - Sistema de Alarmes',
    'FTLB in 48V / out 24 V': 'Infraestrutura - Sistema de Alarmes',
    'OM-SMR100BR-TM-N': 'Infraestrutura - Sistema Retificador FCC',
    ': Emerson - Modelo: R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'UR 37A/48Vcc/412 Modelo OM - 1S37': 'Infraestrutura - Sistema Retificador FCC',
    'EMERSON - Modelo: R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'EMERSON R48 - 3200': 'Infraestrutura - Sistema Retificador FCC',
    'R48-3200': 'Infraestrutura - Sistema Retificador FCC',
    'Omibra OM1S50N': 'Infraestrutura - Sistema Retificador FCC',
    'FLATPAC 2 - 3kW': 'Infraestrutura - Sistema Retificador FCC',
    'OM-SMR100BR-TM-N': 'Infraestrutura - Sistema Retificador FCC',
    'Novus - RHT Modelo: RHT-WM Transmitter': 'Infraestrutura - Sistema Retificador FCC',
    'FLATPACK 2 3000W 5A/-48V/4.1.2 - PN 241119.903': 'Infraestrutura - Sistema Retificador FCC',
    'Aguilera - AE/PX2-F': 'Infraestrutura - Sistema Retificador FCC',
    'UR 37A/48Vcc': 'Infraestrutura - Sistema Retificador FCC',
    'FTLB 4824 S': 'Infraestrutura - Sistema Retificador FCC',
    'Duas unidades SCMD3S1A#': 'Infraestrutura - Sistema Retificador FCC',
    'Bateria Estacionaria Freedom DF2000 115Ah': 'Infraestrutura - Banco de Baterias',
    'DELTA - Modelo: GES161B1057000-N': 'Infraestrutura - Banco de Baterias',
    'SECPower - Modelo: HMA110': 'Infraestrutura - Banco de Baterias',
    'TecPower - Modelo: HNA 12/110': 'Infraestrutura - Banco de Baterias',
    'Modelo: pig-tail de alta E2000 / cordão óptico SC/APC E2000/APC': 'Outros',
    'Cordões com conector E2000': 'Outros',
    'Cordão óptico': 'Outros',
    'MCE  - Modelo: 300W': 'Outros',
    'ODU 8giga hertz ': 'Rádio - Equipamento Digitel',
    'Radio IDU ': 'Rádio - Equipamento Digitel'
}

arquivo = "Lista_modelos_bilhete.xls"
arquivo2 = "Base VDS.xlsx"
abertura_data = []
abertura_hora = []
abertura_mes = []
termino_data = []
termino_hora = []
termino_mes = []
estacao = []
descricao = []
uf = []
regiao = []
tipo_site = []
lider_regiao = []
categoria = []
subcategoria = []
prioridade = []
causa = []

def dividi_data(data_entrada, array_data, array_hora, array_mes):
    if data_entrada == '-':
        array_data.append('-')
        array_hora.append('')
        array_mes.append('-')
    else:
        data_dia, hora = map(str, data_entrada.split(' '))
        h, m ,s = map(int, hora.split(':'))
        array_data.append(data_dia)
        array_hora.append(time(h, m, s))
        array_mes.append(mes_ano(data_dia))

def mes_ano(data_entrada):
    meses = {
        '01': 'janeiro',
        '02': 'fevereiro',
        '03': 'março',
        '04': 'abril',
        '05': 'maio',
        '06': 'junho',
        '07': 'julho',
        '08': 'agosto',
        '09': 'setembro',
        '10': 'outubro',
        '11': 'novembro',
        '12': 'dezembro'
    }
    mes = data_entrada[3:5]
    ano = data_entrada[8:10]
    return f'{meses.get(mes)}-{ano}'

def estacao_id(entrada, entrada_2):
    if isinstance(entrada, float):
        estacao.append(entrada_2.lstrip()[0:11])
        descricao.append(define_descricao(entrada_2.lstrip()[0:11]))
    else:
        estacao.append(entrada)
        descricao.append(define_descricao(entrada))

def estacao_info(estacao_entrada):
    uf.append(estacao_entrada[0:2])
    regiao.append(estados_regioes.get(estacao_entrada[0:2]))
    tipo_site.append(estacao_entrada[6:8])
    lider_regiao.append(lider.get(estacao_entrada[0:2]))

def define_categoria(entrada, entrada_cat, entrada_sub):
    if isinstance(entrada, float) or entrada == None:
        if(entrada_cat, float) or entrada_cat == None:
            categoria.append('')
            subcategoria.append('')
        else:
            categoria.append(entrada_cat)
            subcategoria.append(entrada_sub)
    elif str(entrada) == 'Cancelado':
        categoria.append('Cancelado')
        subcategoria.append('Cancelado')
    elif str(entrada) == 'Outros' or str(entrada) == 'outros' or str(entrada) == 'Outro' or str(entrada) == 'outro':
        categoria.append('Outros')
        subcategoria.append('Outros')
    elif ' - ' in entrada:
        cat, sub = map(str, entrada.split(' - '))
        categoria.append(cat)
        subcategoria.append(sub)
    else:
        categoria.append('')
        subcategoria.append('')

def define_prioridade(entrada):
    if isinstance(entrada, float):
        prioridade.append('Baixa')
    else:
        prioridade.append(entrada)

def define_causa(entrada, entrada2, entrada_base, entrada_nome):
    if '-' in str(entrada):
        causa.append(entrada)
    elif entrada_base == 'Cancelado':
        causa.append('Cancelado')
    else:
        entrada_tratada = str(entrada2).lstrip()
        entrada_tratada = entrada_tratada.upper()
        if 'RETIFICADOR' in entrada_tratada or 'ELTEK' in entrada_tratada or 'EMERSON' in entrada_tratada or 'OMIBRA' in entrada_tratada or '(FONTEOMIBRA)' in entrada_tratada or '(FONTEDELTA)' in entrada_tratada or 'SSC-10' in entrada_tratada or 'ROA4C301AWAHA' in entrada_tratada:
            causa.append('Infraestrutura - Sistema Retificador FCC')
        elif 'TSDA' in entrada_tratada or 'TELEMETRIA' in entrada_tratada or 'SIEMENS' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Alarmes')
        elif 'INVERSOR' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Inversores')
        elif 'NOBREAK' in entrada_tratada or 'NDHBSLM72' in entrada_tratada:
            causa.append('Infraestrutura - Nobreak')
        elif 'BATERIAS' in entrada_tratada:
            causa.append('Infraestrutura - Banco de Baterias')
        elif 'RAU' in entrada_tratada or 'RAU2' in entrada_tratada or 'ERICSSON' in entrada_tratada:
            causa.append('Rádio - Equipamento Ericsson')
        elif 'DIGITEL' in entrada_tratada or 'DSR – ODU' in entrada_tratada:
            causa.append('Rádio - Equipamento Digitel')
        elif 'MOTOR' in entrada_tratada and ('MWM' in entrada_tratada or 'MWM-D-229-4' in entrada_tratada or 'GERADOR' in entrada_tratada) or entrada_tratada == 'BATERIA' or 'GMG' in entrada_tratada:
            causa.append('Infraestrutura - Grupo Motor Gerador')
        elif 'KIT VENTILADOR' in entrada_tratada or 'KIT DE VENTILADOR' in entrada_tratada or 'COMPRESSOR' in entrada_tratada or 'COMPRESSOR.' in entrada_tratada or 'COMPRESSO' in entrada_tratada or 'FAN' in entrada_tratada:
            causa.append('Infraestrutura - Sistema de Climatização')
        elif 'HD' in entrada_tratada or 'DVR' in entrada_tratada or 'INTELBRAS' in entrada_tratada or 'SEAGATE' in entrada_tratada or 'CÂMERA' in str(entrada2).upper() or 'CAMERA' in str(entrada2).upper():
            causa.append('Infraestrutura - Sistema de CFTV')
        elif 'CISCO' in entrada_tratada:
            causa.append('IP - Equipamento Cisco')
        elif 'DATACOM' in entrada_tratada or 'DM4001' in entrada_tratada:
            causa.append('IP - Equipamento Datacom')
        elif 'HUAWEI' in entrada_tratada or 'AR1220EV' in entrada_tratada or 'AR2200E' in entrada_tratada:
            causa.append('IP - Equipamento Huawei')
        elif 'HP' in entrada_tratada:
            causa.append('IP - Equipamento HP')
        elif 'CPE' in entrada_tratada:
            causa.append('IP - Equipamento CPE')
        elif str(entrada_base) != 'nan':
            causa.append(str(entrada_base))
        else:
            if 'PADTEC' in str(entrada_nome).upper():
                causa.append('DWDM - Equipamento PADTEC')
            elif 'CLIMATIZAÇÃO' in str(entrada_nome).upper() or 'AR-CONDICIONADO' in str(entrada_nome).upper() or 'AR CONDICIONADO' in str(entrada_nome).upper():
                causa.append('Infraestrutura - Sistema de Climatização')
            elif 'CFTV' in str(entrada_nome).upper():
                causa.append('Infraestrutura - Sistema de CFTV')
            elif 'MOTOR' in str(entrada_nome).upper() or 'BATERIA' in str(entrada_nome).upper() or 'GERADOR' in str(entrada_nome).upper():
                causa.append('Infraestrutura - Grupo Motor Gerador')
            elif 'ERICSSON' in str(entrada_nome).upper():
                causa.append('Rádio - Equipamento Ericsson')
            elif 'DIGITEL' in str(entrada_nome).upper():
                causa.append('Rádio - Equipamento Digitel')
            elif 'ALARMES' in str(entrada_nome).upper() or 'SIEMENS' in str(entrada_nome).upper() or 'TELEMETRIA' in str(entrada_nome).upper():
                causa.append('Infraestrutura - Sistema de Alarmes')
            else:
                causa.append(causa_dicionario.get(str(entrada2).lstrip()))

arquivo_pd = pd.read_excel(arquivo)
base_pd = pd.read_excel(arquivo2)

for index, row in arquivo_pd.iterrows():
    dividi_data(row['Data abertura'], abertura_data, abertura_hora, abertura_mes)
    dividi_data(row['Término'], termino_data, termino_hora, termino_mes)
    estacao_id(row['Id da Estação'], row['Nome'])
    estacao_info(estacao[index])
    if row['Estado'] == 'Cancelado':
        define_causa(row['Causa do alerta'], row['Sobressalente a ser verificado'], 'Cancelado', row['Nome'])
        define_categoria(causa[index], base_pd['Categoria'][index], base_pd['Subcategoria'][index])
        if isinstance(row['Sobressalente a ser verificado'], float):
            row['Sobressalente a ser verificado'] = 'Cancelado'
    elif  index < base_pd.shape[0]:
        define_causa(row['Causa do alerta'], row['Sobressalente a ser verificado'], base_pd['Causa do alerta'][index], row['Nome'])
        define_categoria(causa[index], base_pd['Categoria'][index], base_pd['Subcategoria'][index])
    else:
        define_causa(row['Causa do alerta'], row['Sobressalente a ser verificado'], '', row['Nome'])
        define_categoria(causa[index], '', '')
    define_prioridade(row['Nível de prioridade da VDS'])
        
df = {
    'Id ordem serviço': arquivo_pd['Id ordem serviço'],
    'Nome': arquivo_pd['Nome'],
    'Estado': arquivo_pd['Estado'],
    'Data abertura': abertura_data,
    'Hora Abertura': abertura_hora,
    'Mês abertura': abertura_mes,
    'Término': termino_data,
    'Hora Término': termino_hora,
    'Mês término': termino_mes,
    'Data ultima transição': arquivo_pd['Data ultima transição'],
    'Estação': estacao,
    'Descrição': descricao,
    'UF': uf,
    'Região': regiao,
    'Tipo Site': tipo_site,
    'Líder de Campo': lider_regiao,
    'Sobressalente a ser verificado': arquivo_pd['Sobressalente a ser verificado'],
    'Nível de prioridade da VDS': prioridade,
    'Causa do alerta': causa,
    'Categoria': categoria,
    'Subcategoria': subcategoria
}

pd_df = pd.DataFrame(df)

print(base_pd)

with pd.ExcelWriter('saida.xlsx', engine='xlsxwriter') as writer:
    pd_df.to_excel(writer, sheet_name='BaseVDS', index=False)
    
    workbook = writer.book
    worksheet = writer.sheets['BaseVDS']
    
    header_format = workbook.add_format({
        'bold': True,
        'text_wrap': True,
        'valign': 'top',
        'fg_color': '#002060',
        'font_color': '#FFFFFF',
        'border': 1
    })

    worksheet.autofilter(0, 0, 0, len(pd_df.columns) - 1)

    center_format = workbook.add_format({'align': 'center'})

   #worksheet.set_column(0, len(pd_df.columns) - 1, cell_format=center_format)

    for col_num, value in enumerate(pd_df.columns.values):
        worksheet.write(0, col_num, value, header_format)

book = load_workbook('saida.xlsx')
sheet = book['BaseVDS']

for col in sheet.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        if len(str(cell.value)) > max_length:
            max_length = len(str(cell.value))
    adjusted_width = (max_length + 2)
    sheet.column_dimensions[column].width = adjusted_width

book.save('saida.xlsx')
import pymssql
from datetime import date
import globalconf as GC


def initDB(rs):
    conn = pymssql.connect(server=GC.server, user=GC.user, password=GC.password, database=GC.database)
    cursor = conn.cursor()
    
    select = f'''
    
    SELECT * FROM cockpit_mds_endosso A
    LEFT JOIN produtoAutomovel_mds_endosso B
    ON B.rs = A.rs
    WHERE A.rs = {rs}

    '''
    cursor.execute(select)
    row = cursor.fetchone()

    conn.close()
    for x in range(0,len(row)):
        if (row[x] == 'Não Informado' or row[x] == 'Nao Informado' ):
            row[x] = ''
    
    if(row[1] != None):
        seguradora = row[1].strip()
    else:
        seguradora = ''
    
    if(row[2] != None):
        tipoEndosso = row[2].strip()
    else:
        tipoEndosso = ''
    
    if(row[3] != None):
        subtipoEndosso = row[3].strip()
    else:
        subtipoEndosso = ''
        
    if(row[4] != None):
        nuapolice = row[4].strip()
    else:
        nuapolice = ''
    
    if(row[5] != None):
        nupropostacia = row[5].strip()
    else:
        nupropostacia = ''
    
    if(row[6] != None):
        segurado = row[6].strip()
    else:
        segurado = ''
    
    if(row[7] != None):
        cpf_cnpj = row[7].strip()
    else:
        cpf_cnpj = ''
    
    if(row[8] != None):
        consultor = row[8].strip()
    else:
        consultor = ''
    
    if(row[9] != None):
        datasolicitacao = row[9]
    else:
        datasolicitacao = ''
    
    if(row[10] != None):
        inicioVigencia = row[10].strip()
    else:
        inicioVigencia = ''
    
    if(row[11] != None):
        fimVigencia = row[11].strip()
    else:
        fimVigencia = ''
    
    if(row[12] != None):
        email = row[12].strip()
    else:
        email = ''
    
    if(row[13] != None):
        cep = row[13].strip()
    else:
        cep = ''
    
    if(row[14] != None):
        endereco = row[14].strip()
    else:
        endereco  = ''
    
    if(row[15] != None):
        cidade = row[15].strip()
    else:
        cidade = ''
    
    if(row[16] != None):
        estado = row[16].strip()
    else:
        estado = ''
    
    if(row[17] != None):
        numero = row[17].strip()
    else:
        numero = ''
    
    if(row[18] != None):
        complemento = row[18].strip()
    else:
        complemento = ''
    
    if(row[19] != None):
        ddd = row[19].strip()
    else:
        ddd = ''
    
    if(row[20] != None):
        telefone = row[20].strip()
    else:
        telefone = ''
    
    if(row[21] != None):
        dddc = row[21].strip()
    else:
        dddc = ''
    
    if(row[22] != None):
        celular = row[22].strip()
    else:
        celular = ''
    
    if(row[23] != None):
        produto = row[23].strip()
    else:
        produto = ''
    
    if(row[24] != None):
        premioLiq= row[24].strip()
    else:
        premioLiq = ''
    
    if(row[25] != None):
        adicional = row[25].strip()
    else:
        adicional = ''
    
    if(row[26] != None):
        formarPgto = row[26].strip()
    else:
        formarPgto = ''
    
    if(row[27] != None):
        qtdParcelas = row[27].strip()
    else:
        qtdParcelas = ''
    
    if(row[28] != None):
        dataVencimento = row[28].strip()
    else:
        dataVencimento = ''
    
    if(row[29] != None):
        valorIOF = row[29].strip()
    else:
        valorIOF = ''
    
    if(row[30] != None):
        premioTotal = row[30].strip()
    else:
        premioTotal = ''
    
    if(row[31] != None):
        nupropostaquiver = row[31].strip()
    else:
        nupropostaquiver = ''

    if(row[32] != None):
        datacadastroquiver = row[32].strip()
    else:
        datacadastroquiver = ''

    if(row[33] != None):
        status = row[33].strip()
    else:
        status = ''

    if(row[34] != None):
        statusRPA = row[34].strip()
    else:
        statusRPA = ''

    if(row[35] != None):
        caminho_doc = row[35].strip()
        #caminho_doc = caminho_doc.replace('\\','/')
    else:
        caminho_doc = ''

    if(row[36] != None):
        IOF = row[36].strip()
    else:
        IOF = ''

    if(row[37] != None):
        jurosmelhorData = row[37].strip()
    else:
        jurosmelhorData = ''

    if(row[38] != None):
        nome_arq = row[38].strip()
    else:
        nome_arq = ''

    if(row[40] != None):
        classeBonus = row[40].strip()
    else:
        classeBonus = ''

    if(row[41] != None):
        codCI = row[41].strip()
    else:
        codCI = ''

    if(row[42] != None):
        categoria = row[42].strip()
    else:
        categoria = ''

    if(row[43] != None):
        franquia_reduzida = row[43].strip()
    else:
        franquia_reduzida= ''

    if(row[44] != None):
        premio_liquido_coberturas = row[44].strip()
    else:
        premio_liquido_coberturas = ''
    
    if(row[45] != None):
        detalhamento_franquia_vidro = row[45].strip()
    else:
        detalhamento_franquia_vidro = ''

    if(row[46] != None):
        nota_fiscal = row[46].strip()
    else:
        nota_fiscal = ''
    if(row[47] != None):
        renavam = row[47].strip()
    else:
        renavam = ''
    if(row[48] != None):
        data_saida = row[48].strip()
    else:
        data_saida = ''
    if(row[49] != None):
        usuario = row[49].strip()
    else:
        usuario = ''

    if (row[52] != None):
        Fabricante = row[52].strip()
    else:
        Fabricante = ''
    

    if (row[53] != None):
        Chassi = row[53].strip()
    else:
        Chassi = ''

    if (row[54] != None):
        FIPE = row[54].strip()
    else:
        FIPE = ''

    if (row[55] != None):
        Placa = row[55].strip()
    else:
        Placa = ''

    if (row[56] != None):
        ano = row[56].strip()
    else:
        ano = ''

    if (row[57] != None):
        Modelo = row[57].strip()
    else:
        Modelo = ''

    if (row[58] != None):
        CEPRISCO = row[58].strip()
    else:
        CEPRISCO = ''
    if (row[60] != None):
        cep_pernoite = row[60].strip()
    else:
        cep_pernoite = ''

    if (row[62] != None):
        idadeCondutor = row[62].strip()
    else:
        idadeCondutor = ''

    if (row[63] != None):
        cpfCondutor = row[63].strip()
    else:
        cpfCondutor = ''

    if (row[64] != None):
        nomeCondutor = row[64].strip()
    else:
        nomeCondutor = ''

    if (row[65] != None):
        estadoCivilCondutor = row[65].strip()
    else:
        estadoCivilCondutor = ''
    if (row[66] != None):
        genero = row[66].strip()
    else:
        genero = ''
    if (row[67] != None):
        combustivel = row[67].strip()
    else:
        combustivel = ''
    

    dataSubs = date.today()
    dataSubs = dataSubs.strftime('%y/%m/%Y')
    custo = '0,0'

    return (
    seguradora
    ,tipoEndosso
    ,subtipoEndosso
    ,nuapolice
    ,nupropostacia
    ,segurado
    ,cpf_cnpj
    ,consultor
    ,datasolicitacao
    ,inicioVigencia
    ,fimVigencia
    ,email
    ,cep
    ,endereco
    ,cidade
    ,estado
    ,numero
    ,complemento
    ,ddd
    ,telefone
    ,dddc
    ,celular
    ,produto
    ,premioLiq
    ,adicional
    ,custo
    ,formarPgto
    ,qtdParcelas
    ,dataVencimento
    ,valorIOF
    ,premioTotal
    ,nupropostaquiver
    ,datacadastroquiver
    ,status
    ,statusRPA
    ,caminho_doc   
    ,Fabricante
    ,Modelo
    ,ano
    ,Placa
    ,Chassi
    ,FIPE
    ,CEPRISCO
    ,dataSubs
    ,classeBonus
    ,codCI
    ,categoria
    ,franquia_reduzida
    ,premio_liquido_coberturas
    ,nomeCondutor
    ,cpfCondutor
    ,estadoCivilCondutor
    ,idadeCondutor
    ,detalhamento_franquia_vidro
    ,nota_fiscal
    ,nome_arq
    ,renavam
    ,genero
    ,cep_pernoite
    ,data_saida
    ,combustivel
    )
import sys

sys.path.append(r"C:\rpa\Python")
from Classes.Oracle.Oracle.ConectaOracle import ConectaOracle



class GeraSQLProdutos:
    def __init__(self):
        pass


    def gera_sql_produtos():
        sql = f"""
        SELECT
        To_Char(A.CD_MERCADORIA||NR_DV_MERCADORIA, 'FM000000') as CODIGO_MERCADORIA,
        NM_MERCADORIA,
        DS_APRESENTACAO_MERCADORIA,
        LISTAGG(NROS_S || ' - ' || NOMS_S || ' - ' || QT_CONCENTRACAO, ' | ')
        WITHIN GROUP (ORDER BY NROS_S) as DADOS_CONCATENADOS,
        A.DS_UNIDADE_MEDIDA,
        NR_NCM,
        CD_EAN_VENDA,
        DS_MINISTERIO_SAUDE,
        ID_SITUACAO_MERCADORIA,
        Decode(DS_NIVEL_ECNM, 1, 'Medicamento', 2, 'Não medicamento', 3, 'Não medicamento', 4, 'Não medicamento'),
        NVL(FAT.DS_FAT_MENSAGEM, 'SEM ISENÇÃO') as DS_FAT_MENSAGEM
        FROM PRDDM.DC_MERCADORIA A
        LEFT JOIN PRDDM.DCAMS AMS ON A.CD_MERCADORIA = AMS.NROM_A
        LEFT JOIN PRDDM.DCSAL SAL ON AMS.NROS_A = SAL.NROS_S
        LEFT JOIN PRDDM.DC_FAT_MERCADORIA_MENSAGEM MENS ON A.CD_MERCADORIA = MENS.CD_MERCADORIA
        AND MENS.ID_ESTADO_ORIGEM = 'SC'
        AND MENS.ID_ESTADO_DESTINO = 'SC'
        LEFT JOIN PRDDM.DC_FAT_MENSAGEM FAT ON MENS.CD_FAT_MENSAGEM = FAT.CD_FAT_MENSAGEM
        WHERE A.DS_APRESENTACAO_MERCADORIA NOT LIKE '%CANCELADO%'
        AND A.DS_NIVEL_ECNM IN (1, 2, 3, 4)
        AND A.ID_SITUACAO_MERCADORIA = 'A'
        GROUP BY
        A.CD_MERCADORIA,
        NR_DV_MERCADORIA,
        NM_MERCADORIA,
        DS_APRESENTACAO_MERCADORIA,
        A.DS_UNIDADE_MEDIDA,
        NR_NCM,
        CD_EAN_VENDA,
        DS_MINISTERIO_SAUDE,
        ID_SITUACAO_MERCADORIA,
        DS_NIVEL_ECNM,
        FAT.DS_FAT_MENSAGEM
        ORDER BY To_Char(A.CD_MERCADORIA||NR_DV_MERCADORIA, 'FM000000')
        """

        oracle = ConectaOracle(sql=sql)
        tabela = oracle.conecta_oracle()

        return tabela
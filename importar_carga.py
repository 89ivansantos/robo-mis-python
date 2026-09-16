import os
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd
import pyodbc


# ============================================================
# CONFIGURAÇÕES
# ============================================================

# Pasta onde os arquivos Excel chegam.
PASTA_ENTRADA = Path(r"C:\arquivos\robo_mis")

# Pasta para arquivos processados com sucesso.
PASTA_PROCESSADOS = PASTA_ENTRADA / "processados"

# Pasta para arquivos que apresentaram erro.
PASTA_ERRO = PASTA_ENTRADA / "error"

# String de conexão obtida por variável de ambiente.
# Não deixar usuário, senha ou servidor diretamente no GitHub.
CONNECTION_STRING = os.getenv("CONNECTION_STRING")

# Nome da aba esperada no Excel.
ABA_EXCEL = "Dados"

# Colunas obrigatórias do arquivo de entrada.
COLUNAS_ESPERADAS = [
    "cliente",
    "produto",
    "valor",
    "data"
]

# Identificador relacionado ao processo.
ID_REFERENCIA = 1313


# ============================================================
# 1. BUSCAR ARQUIVOS
# ============================================================

def buscar_arquivos(pasta):
    """
    Localiza os arquivos .xlsx disponíveis na pasta de entrada.
    """

    return list(Path(pasta).glob("*.xlsx"))


# ============================================================
# 2. VALIDAR NOMENCLATURA
# ============================================================

def validar_nomenclatura(arquivo):
    """
    Valida se o arquivo segue o padrão:

        robo_mis_ddmmaaaa.xlsx

    Além da estrutura do nome, valida se a data informada
    realmente existe.
    """

    nome = arquivo.name

    # Verifica o prefixo.
    if not nome.startswith("robo_mis_"):
        return False

    # Verifica a extensão.
    if arquivo.suffix.lower() != ".xlsx":
        return False

    # Remove prefixo e extensão.
    data = nome[len("robo_mis_"):-5]

    # A data deve possuir exatamente 8 caracteres.
    if len(data) != 8:
        return False

    # A data deve conter somente números.
    if not data.isdigit():
        return False

    # Valida se a data realmente existe.
    try:
        datetime.strptime(data, "%d%m%Y")
    except ValueError:
        return False

    return True


# ============================================================
# 3. VERIFICAR DUPLICIDADE
# ============================================================

def arquivo_existe_banco(nome_arquivo, conexao):
    """
    Verifica se o arquivo já foi processado anteriormente.

    A consulta é feita na tabela tab_mis_arquivo.
    """

    cursor = conexao.cursor()

    sql = """
        SELECT 1
        FROM tab_mis_arquivo
        WHERE nome_arquivo = ?
    """

    cursor.execute(sql, nome_arquivo)

    resultado = cursor.fetchone()

    cursor.close()

    return resultado is not None


# ============================================================
# 4. LER EXCEL
# ============================================================

def ler_excel(arquivo):
    """
    Lê a aba 'Dados' do arquivo Excel.
    """

    df = pd.read_excel(
        arquivo,
        sheet_name=ABA_EXCEL
    )

    return df


# ============================================================
# 5. VALIDAR LAYOUT
# ============================================================

def validar_layout(df):
    """
    Verifica se todas as colunas obrigatórias estão presentes.

    A posição das colunas no Excel não precisa ser a mesma.
    """

    colunas = list(df.columns)

    return all(
        coluna in colunas
        for coluna in COLUNAS_ESPERADAS
    )


# ============================================================
# 6. PREPARAR DADOS
# ============================================================

def preparar_dados(df, data_importacao, nome_arquivo, id_referencia):
    """
    Organiza as colunas do arquivo e adiciona as informações
    de controle da importação.

    Resultado final:
        cliente
        produto
        valor
        data
        data_importacao
        nome_arquivo
        id_referencia
    """

    # Mantém somente as colunas esperadas na ordem correta.
    df = df[COLUNAS_ESPERADAS].copy()

    # Adiciona a data/hora da importação.
    df["data_importacao"] = data_importacao

    # Adiciona o nome do arquivo de origem.
    df["nome_arquivo"] = nome_arquivo

    # Adiciona o identificador do processo.
    df["id_referencia"] = id_referencia

    return df


# ============================================================
# 7. INSERIR DADOS NO SQL SERVER
# ============================================================

def inserir_banco(df, conexao):
    """
    Insere os dados na tabela do SQL Server.

    A quantidade de registros inseridos é retornada
    para utilização no log.
    """

    if df.empty:
        raise ValueError(
            "O arquivo não possui registros para importação."
        )

    cursor = conexao.cursor()

    # Permite maior desempenho em inserções em lote.
    cursor.fast_executemany = True

    sql = """
        INSERT INTO tab_mis_dados (
            cliente,
            produto,
            valor,
            data,
            data_importacao,
            nome_arquivo,
            id_referencia
        )
        VALUES (?, ?, ?, ?, ?, ?, ?)
    """

    registros = []

    for _, linha in df.iterrows():

        registros.append(
            (
                None if pd.isna(linha["cliente"])
                else linha["cliente"],

                None if pd.isna(linha["produto"])
                else linha["produto"],

                None if pd.isna(linha["valor"])
                else linha["valor"],

                None if pd.isna(linha["data"])
                else linha["data"],

                linha["data_importacao"],

                linha["nome_arquivo"],

                linha["id_referencia"]
            )
        )

    cursor.executemany(sql, registros)

    quantidade = len(registros)

    cursor.close()

    return quantidade


# ============================================================
# 8. REGISTRAR ARQUIVO PROCESSADO
# ============================================================

def registrar_arquivo(nome_arquivo, conexao):
    """
    Registra o arquivo processado na tabela tab_mis_arquivo.

    Esse registro permite identificar posteriormente se o
    mesmo arquivo já foi importado.
    """

    cursor = conexao.cursor()

    sql = """
        INSERT INTO tab_mis_arquivo (
            nome_arquivo
        )
        VALUES (?)
    """

    cursor.execute(sql, nome_arquivo)

    cursor.close()


# ============================================================
# 9. REGISTRAR LOG
# ============================================================

def registrar_log(arquivo, status, quantidade, erro=None):
    """
    Registra o resultado do processamento.

    O log é mantido em arquivo local para acompanhamento
    da execução do robô.
    """

    pasta_log = PASTA_ENTRADA / "logs"

    pasta_log.mkdir(
        parents=True,
        exist_ok=True
    )

    arquivo_log = (
        pasta_log /
        f"log_{datetime.now():%Y%m%d}.txt"
    )

    data_hora = datetime.now().strftime(
        "%d/%m/%Y %H:%M:%S"
    )

    mensagem = (
        f"{data_hora} | "
        f"{arquivo.name} | "
        f"{status} | "
        f"Registros: {quantidade}"
    )

    if erro:
        mensagem += f" | Erro: {erro}"

    with open(
        arquivo_log,
        "a",
        encoding="utf-8"
    ) as log:

        log.write(
            mensagem + "\n"
        )


# ============================================================
# 10. MOVER ARQUIVO
# ============================================================

def mover_arquivo(arquivo, pasta_destino):
    """
    Move o arquivo para a pasta correspondente ao resultado
    do processamento.
    """

    pasta_destino.mkdir(
        parents=True,
        exist_ok=True
    )

    destino = pasta_destino / arquivo.name

    shutil.move(
        str(arquivo),
        str(destino)
    )


# ============================================================
# 11. PROCESSAR ARQUIVO
# ============================================================


def processar_arquivo(arquivo, conexao):
    """
    Executa todas as etapas de processamento de um arquivo.

    O banco é confirmado antes da movimentação física do arquivo.
    Dessa forma, caso a movimentação apresente erro após o COMMIT,
    os dados não serão indevidamente considerados como não importados.
    """

    quantidade_importada = 0
    transacao_confirmada = False

    try:

        # ----------------------------------------------------
        # Validação da nomenclatura
        # ----------------------------------------------------

        if not validar_nomenclatura(arquivo):

            raise ValueError(
                "Nomenclatura do arquivo inválida."
            )


        # ----------------------------------------------------
        # Verificação de duplicidade
        # ----------------------------------------------------

        if arquivo_existe_banco(
            arquivo.name,
            conexao
        ):

            raise ValueError(
                "Arquivo já processado anteriormente."
            )


        # ----------------------------------------------------
        # Leitura do Excel
        # ----------------------------------------------------

        df = ler_excel(arquivo)


        # ----------------------------------------------------
        # Validação do layout
        # ----------------------------------------------------

        if not validar_layout(df):

            raise ValueError(
                "Layout do arquivo inválido."
            )


        # ----------------------------------------------------
        # Validação de arquivo vazio
        # ----------------------------------------------------

        if df.empty:

            raise ValueError(
                "Arquivo Excel sem registros."
            )


        # ----------------------------------------------------
        # Preparação dos dados
        # ----------------------------------------------------

        data_importacao = datetime.now()

        df = preparar_dados(
            df,
            data_importacao,
            arquivo.name,
            ID_REFERENCIA
        )


        # ----------------------------------------------------
        # Inserção dos dados
        # ----------------------------------------------------

        quantidade_importada = inserir_banco(
            df,
            conexao
        )


        # ----------------------------------------------------
        # Registro do arquivo processado
        # ----------------------------------------------------

        registrar_arquivo(
            arquivo.name,
            conexao
        )


        # ----------------------------------------------------
        # Confirma a transação
        # ----------------------------------------------------

        conexao.commit()

        transacao_confirmada = True


        # ----------------------------------------------------
        # Movimentação após confirmação do banco
        # ----------------------------------------------------

        try:

            mover_arquivo(
                arquivo,
                PASTA_PROCESSADOS
            )

        except Exception as erro_movimentacao:

            # O banco já foi confirmado.
            # Portanto, NÃO executar rollback aqui.

            registrar_log(
                arquivo,
                "SUCESSO_BANCO_ERRO_MOVIMENTACAO",
                quantidade_importada,
                str(erro_movimentacao)
            )

            return False


        # ----------------------------------------------------
        # Registro de sucesso
        # ----------------------------------------------------

        registrar_log(
            arquivo,
            "SUCESSO",
            quantidade_importada
        )

        return True


    except Exception as erro:

        # ----------------------------------------------------
        # Só executa rollback se o COMMIT ainda não ocorreu.
        # ----------------------------------------------------

        if not transacao_confirmada:

            conexao.rollback()


        # ----------------------------------------------------
        # Movimenta o arquivo para a pasta de erro.
        # ----------------------------------------------------

        try:

            if arquivo.exists():

                mover_arquivo(
                    arquivo,
                    PASTA_ERRO
                )

        except Exception as erro_movimentacao:

            registrar_log(
                arquivo,
                "ERRO",
                quantidade_importada,
                (
                    f"{erro} | "
                    f"Erro ao mover arquivo: "
                    f"{erro_movimentacao}"
                )
            )

            return False


        # ----------------------------------------------------
        # Registro do erro.
        # ----------------------------------------------------

        registrar_log(
            arquivo,
            "ERRO",
            quantidade_importada,
            str(erro)
        )

        return False

# ============================================================
# 12. PROCESSO PRINCIPAL
# ============================================================

def main():
    """
    Controla a execução completa do robô.

    A conexão com o SQL Server é aberta uma única vez
    e utilizada durante todo o processamento.
    """

    # Localiza os arquivos disponíveis.
    arquivos = buscar_arquivos(
        PASTA_ENTRADA
    )

    conexao = None

    try:

        # ----------------------------------------------------
        # Validação da conexão
        # ----------------------------------------------------

        if not CONNECTION_STRING:

            raise ValueError(
                "A variável de ambiente "
                "'CONNECTION_STRING' não foi configurada."
            )


        # ----------------------------------------------------
        # Abre uma única conexão com o banco.
        # ----------------------------------------------------

        conexao = pyodbc.connect(
            CONNECTION_STRING
        )


        # ----------------------------------------------------
        # Processa cada arquivo individualmente.
        # ----------------------------------------------------

        for arquivo in arquivos:

            processar_arquivo(
                arquivo,
                conexao
            )


    finally:

        # ----------------------------------------------------
        # Fecha a conexão ao finalizar o processo.
        # ----------------------------------------------------

        if conexao:

            conexao.close()


# ============================================================
# EXECUÇÃO
# ============================================================

if __name__ == "__main__":
    main()

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

# Pastas de destino conforme o resultado do processamento.
PASTA_PROCESSADOS = PASTA_ENTRADA / "processados"
PASTA_ERRO = PASTA_ENTRADA / "error"

# A conexão é obtida por variável de ambiente.
# Assim, usuário e senha não ficam expostos no GitHub.
CONNECTION_STRING = os.getenv("CONNECTION_STRING")

# Aba esperada no arquivo Excel.
ABA_EXCEL = "Dados"

# Colunas esperadas no arquivo de entrada.
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

    Também verifica se a data informada é uma data válida.
    """

    nome = arquivo.name

    if not nome.startswith("robo_mis_"):
        return False

    if arquivo.suffix.lower() != ".xlsx":
        return False

    # Retira o prefixo "robo_mis_" e a extensão ".xlsx".
    data = nome[len("robo_mis_"):-5]

    if len(data) != 8:
        return False

    if not data.isdigit():
        return False

    # Garante que DDMMYYYY representa uma data real.
    try:
        datetime.strptime(data, "%d%m%Y")
    except ValueError:
        return False

    return True


# ============================================================
# 3. VERIFICAR DUPLICIDADE NO BANCO
# ============================================================

def verificar_arquivo_banco(arquivo, conexao):
    """
    Verifica se o nome do arquivo já existe no banco.

    Retorna:
        True  -> arquivo já existe
        False -> arquivo não existe
    """

    nome = arquivo.name

    try:
        cursor = conexao.cursor()

        sql = """
            SELECT 1
            FROM tab_mis_arquivo
            WHERE nome_arquivo = ?
        """

        cursor.execute(sql, nome)

        resultado = cursor.fetchone()

        return resultado is not None

    except Exception as erro:
        print(
            f"Erro ao verificar arquivo no banco: {erro}"
        )

        # Propaga o erro para que o main() faça o tratamento.
        raise


# ============================================================
# 4. LER EXCEL
# ============================================================

def ler_excel(arquivo):
    """
    Lê a aba 'Dados' do arquivo Excel e retorna um DataFrame.

    Se o arquivo estiver corrompido ou a aba não existir,
    a exceção será tratada pelo main().
    """

    return pd.read_excel(
        arquivo,
        sheet_name=ABA_EXCEL
    )


# ============================================================
# 5. VALIDAR LAYOUT
# ============================================================

def validar_layout(df):
    """
    Verifica se o DataFrame possui as colunas esperadas.

    A comparação ignora:
        - maiúsculas/minúsculas;
        - espaços no início e no final.

    A ordem das colunas não é considerada.
    """

    colunas = [
        str(coluna).lower().strip()
        for coluna in df.columns
    ]

    return set(colunas) == set(COLUNAS_ESPERADAS)


# ============================================================
# 6. PREPARAR DADOS
# ============================================================

def preparar_dados(
    df,
    data_importacao,
    nome_arquivo,
    id_referencia
):
    """
    Acrescenta ao DataFrame as informações de controle
    utilizadas no processo de importação.
    """

    df["data_importacao"] = data_importacao
    df["nome_arquivo"] = nome_arquivo
    df["id_referencia"] = id_referencia

    return df


# ============================================================
# 7. INSERIR DADOS NO BANCO
# ============================================================

def inserir_banco(df, conexao):
    """
    Insere os registros do DataFrame no SQL Server.

    A inserção é realizada em lote com executemany().

    Em caso de sucesso:
        COMMIT

    Em caso de erro:
        ROLLBACK

    Retorna a quantidade de registros do DataFrame.
    """

    try:
        cursor = conexao.cursor()

        # Converte cada linha do DataFrame em uma tupla.
        dados = list(
            df.itertuples(
                index=False,
                name=None
            )
        )

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

        # Inserção em lote.
        cursor.executemany(
            sql,
            dados
        )

        # Confirma a transação somente após a inserção
        # ser concluída sem exceções.
        conexao.commit()

        return len(df)

    except Exception:
        # Se qualquer registro provocar erro,
        # desfazemos a transação.
        conexao.rollback()

        # O erro será tratado pelo main().
        raise


# ============================================================
# 8. MOVER ARQUIVO
# ============================================================

def mover_arquivo(arquivo, pasta_destino):
    """
    Move o arquivo para a pasta de destino.

    Sucesso -> Processados
    Erro    -> Error
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
# 9. REGISTRAR LOG
# ============================================================

def registrar_log(arquivo, status, mensagem):
    """
    Registra o resultado do processamento.

    Neste projeto didático, o log é exibido no console.
    Essa função pode posteriormente ser adaptada para
    gravar em uma tabela de log no SQL Server.
    """

    print(
        f"Arquivo: {arquivo.name} | "
        f"Status: {status} | "
        f"Informação: {mensagem}"
    )


# ============================================================
# 10. PROCESSO PRINCIPAL
# ============================================================

def main():
    """
    Orquestra todo o processo de importação.
    """

    # Localiza os arquivos antes de iniciar o processamento.
    arquivos = buscar_arquivos(
        PASTA_ENTRADA
    )

    # Inicializamos como None para saber se a conexão
    # chegou a ser criada.
    conexao = None

    try:
        # A conexão é aberta uma única vez e reutilizada
        # para todos os arquivos.
        if not CONNECTION_STRING:
            raise ValueError(
                "A variável CONNECTION_STRING não foi configurada."
            )

        conexao = pyodbc.connect(
            CONNECTION_STRING
        )

        # Processa cada arquivo individualmente.
        for arquivo in arquivos:

            try:
                # ------------------------------------------------
                # 1. VALIDAR NOMENCLATURA
                # ------------------------------------------------

                if not validar_nomenclatura(arquivo):
                    continue

                # ------------------------------------------------
                # 2. VERIFICAR DUPLICIDADE
                # ------------------------------------------------

                if verificar_arquivo_banco(
                    arquivo,
                    conexao
                ):
                    mover_arquivo(
                        arquivo,
                        PASTA_ERRO
                    )

                    registrar_log(
                        arquivo,
                        "ERRO",
                        "Arquivo já existe no banco"
                    )

                    continue

                # ------------------------------------------------
                # 3. LER EXCEL
                # ------------------------------------------------

                df = ler_excel(arquivo)

                # ------------------------------------------------
                # 4. VALIDAR LAYOUT
                # ------------------------------------------------

                if not validar_layout(df):
                    mover_arquivo(
                        arquivo,
                        PASTA_ERRO
                    )

                    registrar_log(
                        arquivo,
                        "ERRO",
                        "Layout inválido"
                    )

                    continue

                # ------------------------------------------------
                # 5. PREPARAR DADOS
                # ------------------------------------------------

                data_importacao = datetime.now()

                df = preparar_dados(
                    df,
                    data_importacao,
                    arquivo.name,
                    ID_REFERENCIA
                )

                # ------------------------------------------------
                # 6. INSERIR NO BANCO
                # ------------------------------------------------

                quantidade_importada = inserir_banco(
                    df,
                    conexao
                )

                # ------------------------------------------------
                # 7. MOVER PARA PROCESSADOS
                # ------------------------------------------------

                mover_arquivo(
                    arquivo,
                    PASTA_PROCESSADOS
                )

                # ------------------------------------------------
                # 8. REGISTRAR SUCESSO
                # ------------------------------------------------

                registrar_log(
                    arquivo,
                    "SUCESSO",
                    f"{quantidade_importada} linhas importadas"
                )

            except Exception as erro:
                # ------------------------------------------------
                # ERRO DURANTE O PROCESSAMENTO DO ARQUIVO
                # ------------------------------------------------

                mover_arquivo(
                    arquivo,
                    PASTA_ERRO
                )

                registrar_log(
                    arquivo,
                    "ERRO",
                    str(erro)
                )

                # O erro deste arquivo não interrompe
                # o processamento dos arquivos seguintes.
                continue

    except Exception as erro:
        # --------------------------------------------------------
        # ERRO GERAL DO PROCESSO
        # --------------------------------------------------------

        print(
            f"Erro geral do processo: {erro}"
        )

    finally:
        # A conexão é encerrada uma única vez ao final
        # do processamento.
        if conexao:
            conexao.close()


# ============================================================
# PONTO DE ENTRADA
# ============================================================

if __name__ == "__main__":
    main()

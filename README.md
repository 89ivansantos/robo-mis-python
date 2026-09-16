# 🤖 Automação de ETL — Excel → SQL Server

Projeto de automação desenvolvido em Python para processamento, validação e importação de arquivos Excel em um banco de dados SQL Server.

O fluxo foi estruturado pensando em uma rotina de MIS/BI que precisa processar múltiplos arquivos de forma controlada, evitando duplicidades, validando os dados antes da carga e mantendo rastreabilidade do processamento.

---

## 🎯 Objetivo

Automatizar o processo de importação de arquivos no padrão:

`robo_mis_ddmmaaaa.xlsx`

O robô realiza as seguintes etapas:

1. Localização dos arquivos `.xlsx`
2. Validação da nomenclatura
3. Verificação de arquivo já existente no banco
4. Leitura da aba esperada do Excel
5. Validação do layout
6. Preparação dos dados
7. Inserção em lote no SQL Server
8. Controle transacional com `commit` e `rollback`
9. Movimentação do arquivo
10. Registro do resultado em log

---

## 🔄 Fluxo do processo

```text
Pasta de Entrada
       ↓
Buscar arquivos
       ↓
Validar nomenclatura
       ↓
Arquivo válido?
   ├── NÃO → Ignora
   └── SIM
          ↓
    Verificar no banco
          ↓
      Já existe?
      ├── SIM → ERROR + LOG
      └── NÃO
             ↓
          Ler Excel
             ↓
        Validar layout
             ↓
         Layout válido?
         ├── NÃO → ERROR + LOG
         └── SIM
                ↓
        Preparar dados
                ↓
        Inserir no banco
                ↓
          ┌─────┴─────┐
       SUCESSO       ERRO
          ↓            ↓
       COMMIT       ROLLBACK
          ↓            ↓
    Processados      ERROR
          ↓            ↓
          └──── LOG ───┘
```

---

## 🧩 Principais componentes

### `buscar_arquivos()`

Localiza os arquivos `.xlsx` disponíveis na pasta de entrada e retorna uma lista de arquivos para processamento.

### `validar_nomenclatura()`

Verifica se o arquivo segue o padrão:

`robo_mis_ddmmaaaa.xlsx`

Exemplo:

- `robo_mis_01092026.xlsx` → válido
- `teste.xlsx` → inválido

Também é validada a existência de uma data real no formato `DDMMYYYY`.

### `verificar_arquivo_banco()`

Consulta o SQL Server para verificar se o nome do arquivo já foi registrado.

A consulta utiliza parâmetro:

```sql
SELECT 1
FROM tab_mis_arquivo
WHERE nome_arquivo = ?
```

A conexão é recebida pela função e reutilizada durante o processamento.

### `ler_excel()`

Realiza a leitura da aba `Dados` utilizando Pandas e retorna um `DataFrame`.

### `validar_layout()`

Compara as colunas recebidas no Excel com o conjunto de colunas esperadas.

A validação considera diferenças de maiúsculas/minúsculas e espaços no início ou no final do nome.

### `preparar_dados()`

Adiciona ao `DataFrame` três informações de controle:

- `data_importacao`
- `nome_arquivo`
- `id_referencia`

### `inserir_banco()`

Realiza a carga dos dados no SQL Server utilizando `executemany()` para inserção em lote.

A operação utiliza controle transacional:

```text
INSERT
  ↓
Sucesso → COMMIT
Erro    → ROLLBACK
```

A função retorna a quantidade de registros processados.

### `mover_arquivo()`

Movimenta o arquivo de acordo com o resultado:

```text
Sucesso → Processados
Erro    → Error
```

### `registrar_log()`

Registra o resultado do processamento, incluindo o arquivo, status e informação sobre a operação.

---

## 🛡️ Tratamento de erros

O projeto utiliza dois níveis de tratamento de exceção.

### Erro individual

Um erro durante o processamento de um arquivo não interrompe os demais.

```text
Arquivo 01 → OK
Arquivo 02 → OK
Arquivo 03 → ERRO → ERROR + LOG
Arquivo 04 → OK
Arquivo 05 → OK
```

### Erro geral

Falhas que impedem o processo como um todo, como uma falha na abertura da conexão com o banco, são tratadas separadamente.

---

## 💾 Controle de transação

A inserção dos dados utiliza:

```python
conexao.commit()
```

quando a operação é concluída com sucesso.

Em caso de exceção:

```python
conexao.rollback()
```

Isso evita confirmar uma carga que tenha apresentado falha durante a transação.

---

## 📊 Cenários de processamento

| Arquivo | Situação | Resultado |
|---|---|---|
| `robo_mis_01092026.xlsx` | Arquivo válido | Processado + LOG |
| `teste.xlsx` | Nomenclatura inválida | Descartado |
| `robo_mis_02092026.xlsx` | Já existe no banco | ERROR + LOG |
| `robo_mis_03092026.xlsx` | Layout inválido | ERROR + LOG |
| `robo_mis_04092026.xlsx` | Tudo correto | Processado + LOG |

---

## 🛠️ Tecnologias

- **Python**
- **Pandas**
- **PyODBC**
- **SQL Server**
- **Excel**

---

## 📁 Estrutura do projeto

```text
robo-mis-python/
│
├── importar_carga.py
├── README.md
└── .gitignore
```

---

## 🔐 Segurança

Credenciais de acesso ao banco não devem ser armazenadas diretamente no código ou no GitHub.

A conexão utiliza uma variável de ambiente:

```python
CONNECTION_STRING = os.getenv("CONNECTION_STRING")
```

Dessa forma, informações sensíveis permanecem fora do código-fonte versionado.

---

## 📌 Conceitos aplicados

Este projeto demonstra aplicação prática de:

- Automação de processos
- ETL
- Manipulação de dados com Pandas
- Integração Python + SQL Server
- SQL parametrizado
- Inserção em lote
- Controle transacional
- Tratamento de exceções
- Validação de arquivos
- Validação de layout
- Controle de duplicidade
- Logging
- Organização e modularização de código

---

## 👨‍💻 Sobre o projeto

Projeto desenvolvido como estudo prático de Python aplicado à automação de processos de dados e rotinas de MIS/BI.

O foco foi estruturar um processo que, além de executar a carga, possua validações, tratamento de falhas e rastreabilidade das operações.

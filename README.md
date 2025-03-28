Este script tem como objetivo processar um arquivo Excel contendo números de telefone, filtrar e validar esses números, e então salvar os resultados em arquivos CSV divididos em lotes de até **850.000 registros.** 

Para o caso em especifico, 850.000 foi o ideal tendo em vista ao todo (6 milhões). 
**Importante pontuar** que o script apresentou uma divergência no momento que foi tentado com mais de 850.000 registros. 

O script é essencial para a limpeza e organização de grandes volumes de dados de contato, garantindo que apenas números válidos sejam mantidos e formatados corretamente.


## Explicação do Código

##### Importação das Bibliotecas


```
import pandas as pd
import os
import re
```
*   `pandas` é utilizado para manipulação de dados.
*   `os` é utilizado para interagir com o sistema de arquivos.
*   `re` é utilizado para operações com expressões regulares.


## Definição de Constantes


```
CAMINHO_ENTRADA = r'c:\script\nome-do-seu-arquivo.csv'
PREFIXO_SAIDA = r'c:\script\arquivo-que-serao-registrados-apos-o-capability-check.csv'
LIMITE_POR_CSV = 850000
```

*   `CAMINHO_ENTRADA`: Caminho do arquivo Excel/txt com os contatos.
*   `PREFIXO_SAIDA`: Prefixo para os arquivos CSV de saída.
*   `LIMITE_POR_CSV`: Limite de registros por arquivo CSV.



## Função para Filtrar Telefones


```
def filtrar_telefones(valor):
    """Remove cabeçalhos como 'Telefone 1' e valida números, adicionando prefixo 55."""
    valor = str(valor).strip()
    if re.match(r'^Telefone\s*\d+$', valor, re.IGNORECASE):
        return None
    numero_limpo = re.sub(r'[^\d]', '', valor)
    if len(numero_limpo) >= 8 and not numero_limpo.startswith(('55', '0')):
        return f"55{numero_limpo}"
    elif len(numero_limpo) >= 8:
        return numero_limpo
    else:
        return None
```


*   Remove cabeçalhos como "Telefone 1".
*   Remove caracteres não numéricos.
*   Valida números de telefone.
*   Adiciona o prefixo "55" para números válidos.

## Função Principal


```
def main():
    try:
        if not os.path.exists(CAMINHO_ENTRADA):
            raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_ENTRADA}")
        print("🔍 Lendo todas as abas do arquivo Excel...")
        abas = pd.read_excel(CAMINHO_ENTRADA, sheet_name=None, header=None, engine='openpyxl')
        
        contatos = []
        total_abas = len(abas)
        abas_processadas = set()
        for i, (nome_aba, dados_aba) in enumerate(abas.items(), 1):
            if nome_aba in abas_processadas:
                continue
            abas_processadas.add(nome_aba)
            print(f"\n📊 Processando aba {i}/{total_abas}: '{nome_aba}'...")
            for coluna in [0, 1, 2, 3]:
                numeros = [
                    telefone
                    for valor in dados_aba[coluna].dropna()
                    if (telefone := filtrar_telefones(valor)) is not None
                ]
                contatos.extend(numeros)
                print(f" ✅ Coluna {coluna + 1}: +{len(numeros)} telefones")
        
        contatos = list(set(contatos))
        print(f"\n🔢 Total de telefones únicos encontrados: {len(contatos):,}")
        
        for parte in range(0, len(contatos), LIMITE_POR_CSV):
            lote = contatos[parte : parte + LIMITE_POR_CSV]
            caminho_saida = f"{PREFIXO_SAIDA}_parte_{parte // LIMITE_POR_CSV + 1}.csv"
            pd.DataFrame(lote).to_csv(caminho_saida, index=False, header=False, encoding='utf-8-sig')
            print(f"💾 Arquivo salvo: '{caminho_saida}' ({len(lote):,} registros)")
        
        print("\n✅ Concluído! Todos os lotes foram gerados.")
    except Exception as e:
        print(f"\n❌ Erro durante o processamento:\n{str(e)}")
```

*   Verifica se o arquivo de entrada existe.
*   Lê todas as abas do arquivo Excel.
*   Processa cada aba e coluna, filtrando e validando os números de telefone.
*   Remove duplicados.
*   Divide os resultados em lotes e salva em arquivos CSV.


## Execução do Script


```
if __name__ == "__main__":
    main()
```

*   Chama a função principal `main` quando o script é executado diretamente.

# **Como Usar**

#### Ambiente de Desenvolvimento

Recomenda-se o uso do Visual Studio Code (VSCode) como IDE para desenvolver e executar este script. O VSCode é uma ferramenta poderosa e flexível, ideal para desenvolvimento em Python.

##### Instalação do Visual Studio Code

1.  **Baixe o VSCode**: [Visual Studio Code](https://code.visualstudio.com/)
2.  **Instale o Python Extension**:
    *   Abra o VSCode.
    *   Vá para a aba de extensões (ícone de quadradinho no lado esquerdo).
    *   Pesquise por "Python" e instale a extensão desenvolvida pela Microsoft.

##### Criação e Ativação do Ambiente Virtual

Para garantir que todas as dependências sejam instaladas corretamente e evitar conflitos com outras bibliotecas Python, recomenda-se o uso de um ambiente virtual (`venv`).
1.  **Crie o ambiente virtual**:
    *   No terminal do VSCode, navegue até a pasta do script:


     `cd c:\script`

*   **Crie o ambiente virtual:**

     `python -m venv .venv`

2.  **Ative o ambiente virtual**:
    *   No Windows PowerShell:

     `.\.venv\Scripts\Activate.ps1`

*   No terminal, você verá algo como `(.venv)`, indicando que o ambiente virtual está ativo.

*   Execute o script:

     `python main.py`

#Script completo:

[Validação de contatos Capability - Google Drive](https://drive.google.com/drive/u/0/folders/1g5VNAn0QmpF4-D1ITnkHSwZ-2WlBZzX4)


**Se todos os passos foram executados corretamente, verá o seguinte no terminal:**

![image.png](/.attachments/image-ed086553-1a97-4eb0-a78e-2e811f27c422.png)

Após ler todas as abas e colunas, irá contabilizar e gerar os novos arquivos:
![image.png](/.attachments/image-5ec0edde-a029-4a81-873c-9f325ff3660b.png)



De acordo com os parâmetros passados, os arquivos serão registrados e poderá vê-los prontos para passar para o capability check validar:

![image.png](/.attachments/image-058fcd0c-9990-4c35-bd99-bc6f7ea9d1fc.png)


Após esse tratamento, terá os dados e poderá rodar o script do **capability check**.

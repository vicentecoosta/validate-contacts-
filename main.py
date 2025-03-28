import pandas as pd
import os
import re

CAMINHO_ENTRADA = r'c:\script\publico_alvo_nordeste.xlsx'
PREFIXO_SAIDA = r'c:\script\contatos_consolidados'
LIMITE_POR_CSV = 850000  # Divide os resultados a cada 600k registros

def filtrar_telefones(valor):
    """Remove cabeçalhos como 'Telefone 1' e valida números, adicionando prefixo 55."""
    valor = str(valor).strip()
    
    # Ignora cabeçalhos como "Telefone 1"
    if re.match(r'^Telefone\s*\d+$', valor, re.IGNORECASE):
        return None
    
    # Remove caracteres não numéricos
    numero_limpo = re.sub(r'[^\d]', '', valor)
    
    # Verifica se é um número válido (mínimo 8 dígitos, sem DDI)
    if len(numero_limpo) >= 8 and not numero_limpo.startswith(('55', '0')):
        return f"55{numero_limpo}"  # Adiciona prefixo 55 se não tiver
    elif len(numero_limpo) >= 8:
        return numero_limpo  # Mantém se já tiver 55 ou 0 no início
    else:
        return None  # Ignora números muito curtos

def main():
    try:
        # Verifica se o arquivo existe
        if not os.path.exists(CAMINHO_ENTRADA):
            raise FileNotFoundError(f"Arquivo não encontrado: {CAMINHO_ENTRADA}")
        
        print("🔍 Lendo todas as abas do arquivo Excel...")
        abas = pd.read_excel(
            CAMINHO_ENTRADA,
            sheet_name=None,  # Lê todas as abas
            header=None,      # Sem cabeçalho
            engine='openpyxl'
        )
        
        # Processa todas as abas e colunas (garantindo que cada aba seja única)
        contatos = []
        total_abas = len(abas)
        abas_processadas = set()  # Armazena abas já processadas
        
        for i, (nome_aba, dados_aba) in enumerate(abas.items(), 1):
            if nome_aba in abas_processadas:
                continue  # Pula se a aba já foi processada
                
            abas_processadas.add(nome_aba)  # Marca como processada
            print(f"\n📊 Processando aba {i}/{total_abas}: '{nome_aba}'...")
            
            for coluna in [0, 1, 2, 3]:  # Colunas A, B, C, D (0 a 3)
                numeros = [
                    telefone
                    for valor in dados_aba[coluna].dropna()
                    if (telefone := filtrar_telefones(valor)) is not None
                ]
                contatos.extend(numeros)
                print(f" ✅ Coluna {coluna + 1}: +{len(numeros)} telefones")
        
        # Remove duplicados
        contatos = list(set(contatos))
        print(f"\n🔢 Total de telefones únicos encontrados: {len(contatos):,}")
        
        # Divide em lotes de 850K e salva em CSVs separados
        for parte in range(0, len(contatos), LIMITE_POR_CSV):
            lote = contatos[parte : parte + LIMITE_POR_CSV]
            caminho_saida = f"{PREFIXO_SAIDA}_parte_{parte // LIMITE_POR_CSV + 1}.csv"
            
            # Salva sem cabeçalho (header=False)
            pd.DataFrame(lote).to_csv(
                caminho_saida,
                index=False,
                header=False,  # SEM CABEÇALHO
                encoding='utf-8-sig'
            )
            print(f"💾 Arquivo salvo: '{caminho_saida}' ({len(lote):,} registros)")
        
        print("\n✅ Concluído! Todos os lotes foram gerados.")
    
    except Exception as e:
        print(f"\n❌ Erro durante o processamento:\n{str(e)}")

if __name__ == "__main__":
    main()
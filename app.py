import pandas as pd
print("--- Automatizador de Planilhas v1 ---")
try:
    df = pd.read_excel("dados_brutos.xlsx")
except FileNotFoundError:
    print("Criando dados_brutos.xlsx de exemplo...")
    pd.DataFrame([['joao silva ',100],['JOAO SILVA',100],[' maria ',None],['pedro',50],['Ana',200]], columns=['nome','valor']).to_excel('dados_brutos.xlsx', index=False)
    df = pd.read_excel("dados_brutos.xlsx")
print(f"Antes: {len(df)} linhas")
df.columns = df.columns.str.strip().str.lower()
df['nome'] = df['nome'].astype(str).str.strip().str.title().replace('Nan', '')
df['valor'] = pd.to_numeric(df['valor'], errors='coerce').fillna(0)
df = df[df['nome'] != '']
df = df.drop_duplicates()
df.to_excel("dados_limpos.xlsx", index=False)
print(f"Depois: {len(df)} linhas")
print(df)
print("\n✅ Pronto! Arquivo dados_limpos.xlsx criado.")

import pandas as pd
import matplotlib.pyplot as plt

# Carregar o arquivo Excel
df = pd.read_excel('DataSet v2.xlsx')
df['Data'] = pd.to_datetime(df['Data']).dt.date
df['Taxa Indicativa'] = pd.to_numeric(df['Taxa Indicativa'],errors = 'coerce')

print(df.dtypes)
# Função para plotar o gráfico baseado no indicador
def plot_taxa_indicativa(indicador, titulo_grafico, cor):
    # Filtrar os dados para o indicador
    df_filtered = df[df['Indexador'] == indicador]
    
    # Verificar se o DataFrame não está vazio
    if df_filtered.empty:
        print(f"Nenhum dado encontrado com o indexador {indicador}.")
    else:
        # Calculando a taxa indicativa média por data
        taxa_media = df_filtered.groupby('Data')['Taxa Indicativa'].mean()
        
        # Plotando o gráfico de linha apenas com os dias que possuem valores
        plt.figure(figsize=(10,6))
        plt.plot(taxa_media.index, taxa_media, marker='o', color=cor)  # Gráfico de linha com marcadores
        plt.title(titulo_grafico)
        plt.xlabel('Data')
        plt.ylabel('Taxa Indicativa Média')
        plt.xticks(taxa_media.index, rotation=45)  # Mostrar apenas as datas que possuem valores
        plt.grid(True)  # Adiciona uma grade para facilitar a leitura do gráfico

        # Adicionando os valores exatos no gráfico
        for data, taxa in taxa_media.items():
            plt.text(data, taxa, f'{taxa:.4f}', ha='center', va='bottom', fontsize=10, color='black')

        # Mostrar o gráfico
        plt.tight_layout()  # Ajusta automaticamente o layout para evitar cortes nos rótulos
        plt.show()

# Plotando os gráficos para os diferentes indicadores
plot_taxa_indicativa('%DI', 'Taxa Indicativa Média por Data (Indicador % DI)', 'blue')
plot_taxa_indicativa('DI+', 'Taxa Indicativa Média por Data (Indicador DI+)', 'green')
plot_taxa_indicativa('IPCA+', 'Taxa Indicativa Média por Data (Indicador IPCA+)', 'red')
    
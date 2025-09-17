import pandas as pd


df = pd.read_csv("Pasta1.csv", encoding="latin1", sep=";")
novo_orcamento = [{'fornecedor': 'RETIFICADORA TIETE', 'orçamento': 33638, 'frota': 4200640},
                  {'fornecedor': 'UNIMAQ ASSIS', 'orçamento': 44500, 'frota': 4300199},
                  {'fornecedor': 'DENER RADIADORES', 'orçamento': 2425, 'frota': 4100624},
                  {'fornecedor': 'AUTO CAPAS PACOCA', 'orçamento': 5250, 'frota': 4100535},
                  {'fornecedor': 'AOKI', 'orçamento': 22141, 'frota': 4100288},
                  {'fornecedor': 'PB LOPES', 'orçamento': 206473, 'frota': 4100500},
                  {'fornecedor': 'GERMANICA', 'orçamento': 21973, 'frota': 4506364},                  
                  {'fornecedor': 'LAPONIA', 'orçamento': 53394343, 'frota': 4100501},
                  {'fornecedor': 'TORNEARIA VALVERDE', 'orçamento': 15643, 'frota': 4402021},                
]

df_novos = pd.DataFrame(novo_orcamento)

df = pd.concat([df, df_novos], ignore_index=True)

print(df)

df.to_csv("Pasta1.csv", sep=";", index=False, encoding="latin1")
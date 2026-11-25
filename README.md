# validador-termodinamico-ia

Repositório complementar ao Trabalho de Conclusão de Curso (Bacharelado em Engenharia de Software – Univille): **"O Impacto do Machine Learning na Mitigação de Gargalos Analíticos na Validação de Ensaios Termodinâmicos"**.

# Sobre o Projeto
Ferramenta *desktop* em Python (`Tkinter` + `Scikit-Learn`) que utiliza o algoritmo **Random Forest** para classificar automaticamente resultados de ensaios termodinâmicos de compressores em quatro categorias (`AVG`, `A`, `B` e `discard`) sob arquitetura *Human-in-the-Loop*.

# Estrutura do Repositório
- `analise_tcc.ipynb`: pipeline de ETL, treinamento do modelo (`Random Forest`, `n_estimators=100`, `random_state=42`, `class_weight='balanced'`), validação cruzada agrupada por planilha (`StratifiedGroupKFold`, $k=5$) e explicabilidade via **SHAP** (`TreeExplainer`).
- `app.py`: interface gráfica do aplicativo *Preenchimento_IAutomático* com mapeamento dinâmico de colunas e injeção em bloco no Excel (`win32com`).
- `requirements.txt`: dependências e versões utilizadas (`scikit-learn==1.9.1`).
- `historico_testes/`: planilhas anonimizadas no formato esperado (aba `CAL`) para testes de reprodutibilidade no treinamento da IA.
- `testes_app/`: planilhas anonimizadas com a coluna de decisão em branco para testar o preenchimento automático pelo aplicativo (`app.py`).

# Como Executar
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Execute o notebook `analise_tcc.ipynb` para gerar o modelo (`cerebro_ia.pkl` e `colunas_ia.pkl`) e reproduzir os logs de validação cruzada e SHAP.
3. Execute a aplicação desktop e selecione uma das planilhas da pasta `testes_app/`:
   ```bash
   python app.py
   ```

## Principais Resultados
- **Eficiência operacional:** redução média de ~97% no tempo de ciclo analítico (de 135 min para 4 min).
- **Validação cruzada (`StratifiedGroupKFold`, k=5):** acurácia média de 97,34% (± 1,98%) e F1-Score macro de 57,39% (± 25,03%), refletindo o desbalanceamento das classes minoritárias.
- **Explicabilidade (SHAP):** `Var_COP` (17,99%) e `Var_Qo` (14,42%) concentraram 32,41% do peso decisório.

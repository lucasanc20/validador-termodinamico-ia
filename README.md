## validador-termodinamico-ia

Repositório complementar ao Trabalho de Conclusão de Curso (Bacharelado em Engenharia de Software – Univille): **"O Impacto do Machine Learning na Mitigação de Gargalos Analíticos na Validação de Ensaios Termodinâmicos"**.

## Sobre o Projeto
Ferramenta *desktop* em Python (`Tkinter` + `Scikit-Learn`) que utiliza o algoritmo **Random Forest** para classificar automaticamente resultados de ensaios termodinâmicos de compressores em quatro categorias (`AVG`, `A`, `B` e `discard`) sob arquitetura *Human-in-the-Loop*.

## Estrutura do Repositório
- `analise_tcc.ipynb`: pipeline de ETL, treinamento do modelo (`Random Forest`, `n_estimators=100`, `random_state=42`, `class_weight='balanced'`), validação cruzada agrupada por planilha (`StratifiedGroupKFold`, $k=5$) e explicabilidade via **SHAP** (`TreeExplainer`).
- `app.py`: interface gráfica do aplicativo *Preenchimento_IAutomático* com mapeamento dinâmico de colunas e injeção em bloco no Excel (`win32com`).
- `requirements.txt`: dependências e versões utilizadas (`scikit-learn==1.9.1`).
- `historico_testes/`: planilhas anonimizadas no formato esperado (aba `CAL`) para testes de reprodutibilidade no treinamento da IA.
- `testes_app/`: planilhas anonimizadas com a coluna de decisão em branco para testar o preenchimento automático pelo aplicativo (`app.py`).

## Pré-requisitos
- **Python 3.10+** instalado.
- **Microsoft Excel** instalado em ambiente Windows (necessário para a biblioteca `pywin32` / `win32com` realizar a injeção em bloco no arquivo).

## Como Executar

1. **Baixe o repositório:**
   - **Via Git:**
     ```bash
     git clone [https://github.com/SEU-USUARIO/validador-termodinamico-ia.git](https://github.com/SEU-USUARIO/validador-termodinamico-ia.git)
     cd validador-termodinamico-ia
     ```
   - **Ou via Download ZIP:** clique no botão verde **<> Code** no topo desta página → **Download ZIP**, extraia a pasta no seu computador e abra o terminal dentro dela.

2. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```

3. **Treine o modelo e reproduza as métricas:**
   - Abra o arquivo `analise_tcc.ipynb` (no VS Code ou Jupyter) e clique em **Executar Tudo (Run All)**.
   - O notebook lerá automaticamente as planilhas da pasta `historico_testes/`, exibirá os logs da validação cruzada (`StratifiedGroupKFold`, $k=5$) e do **SHAP**, e criará na pasta raiz os arquivos `cerebro_ia.pkl` e `colunas_ia.pkl`.

4. **Execute o aplicativo desktop:**
   ```bash
   python app.py
   ```
   - Na janela aberta, clique no botão de análise e selecione uma das planilhas em branco da pasta `testes_app/`.
   - Ao finalizar, abra a planilha selecionada no Excel para conferir as decisões preenchidas automaticamente pela IA.

## Principais Resultados
- **Eficiência operacional:** redução média de ~97% no tempo de ciclo analítico (de 135 min para 4 min).
- **Validação cruzada (`StratifiedGroupKFold`, k=5):** acurácia média de 97,34% (± 1,98%) e F1-Score macro de 57,39% (± 25,03%), refletindo o desbalanceamento das classes minoritárias.
- **Explicabilidade (SHAP):** `Var_COP` (17,99%) e `Var_Qo` (14,42%) concentraram 32,41% do peso decisório.

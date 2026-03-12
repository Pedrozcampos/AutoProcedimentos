#  ProcAuditoria: Automação de Auditoria Contábil com Python

O **ProcAuditoria** é uma ferramenta de Auditoria Contábil Inteligente desenvolvida para automatizar a análise de Razões Contábeis (arquivos Excel/CSV). O sistema utiliza processamento de dados para identificar riscos críticos e gera relatórios formatados prontos para análise técnica, incluindo um dashboard visual. Ele segue um padrão de procedimentos chamado Journal Entries, ele é comum em análises de auditoria e contabilidade.

## O Problema: Lentidão no procedimento causando demora em entrega
No processo de auditoria tradicional (ex: Journal Entries), a identificação de lançamentos atípicos é frequentemente manual, resultando em:
* **Baixa Eficiência:** Processamento lento de milhares de registros.
* **Risco de Erro:** Falhas na aplicação de filtros ou fórmulas complexas.
* **Despadronização:** Dificuldade em manter a consistência dos relatórios entre diferentes auditores.

## A Solução: Inteligência na limpeza e Análise de Dados
O ProcAuditoria agiliza a vida do profissional. Um trabalho que ia levar horas ou até dias para limpar e fazer os procedimentos você faz em poucos minutos no ProcAuditoria. Pensei em arquivos de 50, 250 ou ate 600 mil linhas, então, agilidade e otimização forão levados em consideração para ser feito a automação.
1. **Tratamento de Dados "Sujos":** O código possui lógica de localização dinâmica que ignora automaticamente cabeçalhos inúteis e linhas de "Saldo Anterior", encontrando o início real da tabela de dados.
2. **Mapeamento Flexível de Colunas:** Através de busca por palavras-chave, o sistema identifica colunas de Data, Histórico e Valores, independente do layout ou sistema ERP do cliente.
3. **Padronização e Rigor Técnico:** O software gera um relatório final estruturado onde cada teste de auditoria (Redondos, FDS, ET) possui sua própria aba, já acompanhada de Objetivo, Procedimento e Conclusão descritos.
4. **UX e Estabilidade:** Interface moderna com uso de *threading* para garantir que o processamento de grandes bases de dados ocorra sem travamentos na aplicação.

## Procedimentos de Auditoria Automatizados ( Journal Entries )
O software aplica os seguintes testes:
* **Erro Tolerável (ET):** Filtra lançamentos que excedem o Erro Tolerável definido pelo usuário.
* **Valores Redondos:** Identifica lançamentos que podem indicar estimativas ou falta de precisão.
* **Final De Semana:** Detecta movimentações realizadas em finais de semana (Sábados e Domingos).
* **Palavras Chaves:** Busca termos como "ajuste", "estorno", "erro", "manual" e "urgente" no histórico.
* **Abate Débito/Crédito:** Verifica o equilíbrio matemático entre os totais de Débito e Crédito.

## Tecnologias Utilizadas
* **Python**: Linguagem do projeto.
* **Pandas**: Manipulação e análise de dados de alta performance.
* **CustomTkinter**: Interface moderna (GUI) com suporte a Dark Mode.
* **Matplotlib**: Geração de gráficos para o Dashboard de riscos.
* **Openpyxl**: Estilização avançada e formatação de planilhas Excel.
* **Regex**: Utilizado para buscar um padrão de forma flexível.

##  Como Executar
1. Instale as dependências:

   pip install pandas customtkinter matplotlib openpyxl python-calamine regex


## EXE
Tem suporte para se tornar um .exe e a pessoa que for usar não precisa ter o vs code e as bibliotecas etc... Facilitando para o cliente.
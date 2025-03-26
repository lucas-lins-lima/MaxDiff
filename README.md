**1. Introdução**

**Descrição do Projeto**

O intuito deste projeto é realizar o sorteio de frases ou atributos relacionados ao mercado imobiliário, em particular sobre imóveis específicos, utilizando o método de MaxDiff para efetuar o sorteio das frases. O MaxDiff é empregado aqui como uma técnica para determinar quais frases ou atributos são mais relevantes ou preferidos em um conjunto de opções disponíveis, proporcionando uma visão mais clara das preferências em contextos onde múltiplas versões ou descrições estão envolvidas.

**Motivação**
Minha principal motivação para criar este projeto é a busca constante por aprendizado. No meu ambiente de trabalho, frequentemente encontro situações onde esse tipo de atividade é relevante, e considero fundamental entender de maneira prática o funcionamento do sorteio de frases e o conceito subjacente ao MaxDiff. Este projeto busca solucionar problemas relacionados a sorteios aleatórios, oferecendo um método estruturado para entender e aplicar preferências dentro do contexto do mercado imobiliário. Essa prática não só melhora a compreensão teórica, mas também oferece insights sobre a aplicação prática destas técnicas em cenários do mundo real.

**2. O que é MaxDiff?**

**Definição**

MaxDiff, ou Maximum Difference Scaling, é uma técnica estatística usada para medir as preferências relativas entre diferentes itens ou atributos. Essa metodologia é comumente utilizada em pesquisas de mercado para entender quais atributos de um produto ou serviço são mais valorizados pelos consumidores. Ao apresentar conjuntos de opções e solicitar que os participantes selecionem os itens mais e menos preferidos, o MaxDiff reúne dados que refletem as preferências agregadas de forma mais eficaz do que métodos tradicionais.

**Aplicação no Projeto**
Neste projeto, o MaxDiff é aplicado para realizar o sorteio de frases ou atributos relacionados a imóveis específicos. Utilizamos quatro principais métricas para assegurar que o processo de seleção seja robusto e compreensível:

**Número Total de Frases Distintas:** Determina quantas frases ou atributos únicos serão incluídos no conjunto de opções analisadas. Essa métrica garante uma variedade suficiente para proporcionar uma avaliação abrangente das preferências.

**Cobertura (Repetição de Frases):** Definimos quantas vezes cada frase deve aparecer nos diferentes conjuntos (ou telas). Isso assegura que todas as frases tenham uma cobertura adequada e sejam avaliadas pelo mesmo número de vezes, garantindo a representatividade no levantamento de dados.

**Número de Frases por Conjunto (Tela):** Especifica quantas frases serão mostradas simultaneamente em cada conjunto apresentado aos entrevistados. Este número equilibra a complexidade cognitiva necessária para a seleção, sem sobrecarregar os participantes.

**Número de Entrevistados:** Estabelece quantos indivíduos participarão do estudo, contribuindo para a robustez estatística dos resultados. Um número maior de entrevistados pode aumentar a confiança nas conclusões extraídas dos dados.

Essas métricas formam a base para a implementação eficaz do MaxDiff no sorteio de frases, oferecendo insights sobre quais atributos são mais valorizados em um cenário imobiliário. A abordagem metodológica rigorosa permite que as empresas de mercado imobiliário analisem dados de preferência de maneira eficaz e fundamentada.

**3. Funcionamento do MaxDiff**

**Visão Geral do Algoritmo**

O MaxDiff, neste projeto, é implementado para efetuar uma seleção otimizada de frases ou atributos, garantindo que as preferências dos entrevistados sejam bem mapeadas e entendidas. Aqui está uma visão detalhada de como o algoritmo funciona e quais são suas partes principais:

**Processos e Funções:**
**Cálculo de Conjuntos Otimizados**

- **Função: calcular_conjuntos_otimizados**
- Esta função calcula o número mínimo de conjuntos necessários para apresentar todas as frases de forma que cada uma apareça em um conjunto um número alvo de vezes.
Parâmetros:
- **num_frases:** Quanto mais frases diferentes temos, maior o número de conjuntos necessário.
- **aparicoes_alvo:** Quantas vezes queremos que uma frase apareça, garantindo cobertura adequada.
- **itens_por_conjunto:** O número de frases que aparecem em cada tela apresentada aos entrevistados.
- Esta abordagem assegura que a variação e a cobertura sejam mantidas em um nível aceitável, distribuindo as frases de forma equilibrada.

**Geração de Conjuntos MaxDiff e Exportação**

- **Função: gerar_excel_para_entrevistados**
- Baseando-se nas combinações calculadas, esta função gera conjuntos MaxDiff para cada entrevistado e exporta os dados em um arquivo Excel, organizando as informações de forma clara e acessível.
- **Detalhes do Processo:**
- Para cada entrevistado, são gerados conjuntos de frases de acordo com as métricas calculadas previamente.
- As frases são organizadas aleatoriamente para evitar qualquer viés.
- Diversas combinações garantem que todos os atributos sejam cobertos uniformemente, evitando repetições inadequadas.

**Validação dos Dados de Entrada**

- **Função: validar_entrada**
- Antes de executar o sorteio, essa função assegura que o número de frases e a configuração dos conjuntos sejam suficientes para cobrir as necessidades da amostra.
- **Verificações Incluídas:**
- Garante que o número de frases não é menor que o necessário para preencher os conjuntos.
- Avalia se as configurações permitirão a repetição necessária das frases.

**Relatório de Validação**

- **Função: gerar_relatorio_validacao**
- Gera um relatório textual que resume os resultados da validação das entradas, indicando se há algum problema que precise ser corrigido antes da execução completa do algoritmo.

**Saída Esperada**
- O resultado final é um arquivo Excel contendo a organização completa dos conjuntos MaxDiff otimizados, com resumos claros sobre quantos conjuntos foram gerados e a distribuição das frases por conjunto.
- Outras saídas incluem um relatório de validação que ajuda a contextualizar se os dados configurados foram suficientes e corretamente distribuídos.
Este esquema geral do funcionamento do algoritmo MaxDiff no projeto fornece um entendimento sólido de como a técnica é aplicada para sortear frases dentro do contexto imobiliário. O código é projetado para maximizar a eficácia das decisões dos entrevistados, garantindo que as preferências sejam refletidas de forma justo e precisa.

**4. Sorteio de Frases**

**Objetivo do Sorteio de Frases**
O sorteio de frases é uma etapa crucial do projeto, projetada para garantir que todas as frases ou atributos relacionados aos imóveis sejam apresentados aos entrevistados de maneira equilibrada. A ideia é explorar as preferências dos entrevistados sem viés e de forma abrangente, maximizando a eficiência do levantamento de dados.

**Processo do Sorteio**

**Preparação e Estruturação**

Primeiramente, as frases relacionadas aos atributos dos imóveis são listadas e indexadas em um dicionário, onde cada frase é associada a um código único. Isso facilita o acompanhamento e a identificação das frases ao longo do processo.

**Algoritmo de Sorteio**

- **Seleção Aleatória:** Utilizando a função np.random.choice, o código seleciona aleatoriamente um número específico de frases para cada conjunto (ou tela) a ser apresentado. Esta seleção é feita sem reposição, garantindo que as frases em um mesmo conjunto sejam distintas.
- **Número de Conjuntos:** O total de conjuntos, calculado na etapa anterior, é gerado para cada entrevistado de forma que todas as frases sejam adequadamente cobertas, conforme o número alvo de aparições estabelecido.

**Distribuição Balanceada**

O sorteio é projetado para assegurar que todas as frases aparecem em diferentes combinações através dos conjuntos, minimizando a repetição e maximizando a diversidade apresentada aos entrevistados. Este equilíbrio é essencial para a veracidade dos dados coletados, evitando que determinadas palavras sejam sub-representadas.

**Criação e Armazenamento dos Dados**

Uma vez que as frases são sorteadas para cada conjunto e entrevistado, estas informações são organizadas em um DataFrame e exportadas para um arquivo Excel. Isso permite que os dados sejam facilmente analisados e revisados, além de permitir a visualização clara dos conjuntos apresentados.

**Benefícios do Sorteio**

- **Cobertura Completa:** O uso de sorteio aleatório garante que todas as frases têm igual oportunidade de serem selecionadas, cumprindo com os requisitos de aparição definidos.
- **Redução de Viés:** A aleatoriedade no sorteio proporciona um campo de avaliação mais justo e, ao mesmo tempo, aprofunda a análise comportamental dos entrevistados.
- **Facilidade de Implementação:** Automatizar o sorteio e a exportação dos dados em arquivos organizados simplifica o processo de coleta e análise de dados, permitindo um foco maior na interpretação dos resultados.

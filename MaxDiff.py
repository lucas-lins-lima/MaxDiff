import numpy as np
import pandas as pd
import math
import os

def calcular_conjuntos_otimizados(num_frases, aparicoes_alvo=3, itens_por_conjunto=7):
    """
    Calcula o número mínimo de conjuntos necessário dadas as restrições.

    :param num_frases: Número total de frases distintas.
    :param aparicoes_alvo: Quantas vezes você quer que cada frase apareça (para garantir cobertura).
    :param itens_por_conjunto: Número de frases por conjunto (tela).
    :return: Número mínimo de conjuntos necessários.
    """
    num_conjuntos = math.ceil(num_frases * aparicoes_alvo / itens_por_conjunto)
    return num_conjuntos, itens_por_conjunto

def gerar_excel_para_entrevistados(frases, num_entrevistados, aparicoes_alvo=3, nome_arquivo='sorteio_maxdiff.xlsx'):
    """
    Gera um arquivo Excel com conjuntos MaxDiff otimizados com base na melhor combinação calculada.
    """
    num_frases = len(frases)
    num_conjuntos, itens_por_conjunto = calcular_conjuntos_otimizados(num_frases, aparicoes_alvo)
    
    codigos_frases = {frase: f"F{indice+1}" for indice, frase in enumerate(frases)}
    
    todos_os_dados = {'Entrevistado': [], 'Conjunto': [], 'Codigo_Frase': [], 'Frase': []}
    for id_entrevistado in range(1, num_entrevistados + 1):
        conjuntos_maxdiff = [np.random.choice(frases, itens_por_conjunto, replace=False).tolist()
                             for _ in range(num_conjuntos)]
        
        for indice_conjunto, conjunto_maxdiff in enumerate(conjuntos_maxdiff):
            for frase in conjunto_maxdiff:
                todos_os_dados['Entrevistado'].append(f"Entrevistado {id_entrevistado}")
                todos_os_dados['Conjunto'].append(f"Conjunto {indice_conjunto + 1}")
                todos_os_dados['Codigo_Frase'].append(codigos_frases[frase])
                todos_os_dados['Frase'].append(frase)
    
    df_todos = pd.DataFrame(todos_os_dados)
    
    total_conjuntos = num_entrevistados * num_conjuntos
    
    with pd.ExcelWriter(nome_arquivo, engine='xlsxwriter') as escritor:
        df_todos.to_excel(escritor, index=False, sheet_name='Atribuições MaxDiff')
        
        dados_resumo = {
            'Total de Telas': [total_conjuntos],
            'Telas por Entrevistado': [num_conjuntos],
            'Frases por Tela': [itens_por_conjunto]
        }
        df_resumo = pd.DataFrame(dados_resumo)
        df_resumo.to_excel(escritor, index=False, sheet_name='Resumo')

    print(f"Arquivo Excel '{nome_arquivo}' foi criado com conjuntos MaxDiff otimizados.")

# Exemplo de uso baseado em frases
frases = [
    "Próximo a escolas, supermercados e transporte público",
    "Desfrute de incríveis vistas do horizonte da cidade",
    "Salas e quartos espaçosos para maior conforto",
    "Espaço ideal para confraternizações com churrasqueira",
    "Com piscina, academia e salão de festas",
    "Materiais de alta qualidade em todos os cômodos",
    "Completa com armários embutidos de primeira linha",
    "Grandes janelas para aproveitar a luz do dia",
    "Com closet e banheira de hidromassagem",
    "Um refúgio verde no meio da cidade",
    "Porteiro 24 horas e câmeras de vigilância",
    "Duas vagas cobertas e espaço para visitantes",
    "Espaço adequado para o seu amigo de quatro patas",
    "Atualizações modernas e elegantes",
    "A poucos passos das areias brancas",
    "Sustentabilidade em primeiro lugar",
    "Conforto térmico em todos os ambientes",
    "Ambiente silencioso e bem iluminado",
    "Estrutura robusta e bem conservada",
    "Economia na conta de luz com painéis solares",
    "Diversão garantida para toda a família",
    "Acesso direto e exclusivo ao seu andar",
    "Imóvel com arquitetura charmosa e única",
    "Para momentos agradáveis com amigos",
    "Facilite seu deslocamento diário",
    "Espaço destinado para lavanderia",
    "Otimização de espaço nos dormitórios",
    "Ótima opção para relaxar ao ar livre",
    "Incentive a alimentação saudável",
    "Arquitetura contemporânea e funcional"
]

# Parâmetros otimizados derivados
num_entrevistados = 300
gerar_excel_para_entrevistados(frases, num_entrevistados)


def validar_entrada(frases, aparicoes_alvo, itens_por_conjunto):
    """
    Valida as entradas fornecidas para o algoritmo MaxDiff.
    
    :param frases: Lista de frases únicas.
    :param aparicoes_alvo: Quantas vezes cada frase deve aparecer.
    :param itens_por_conjunto: Número de frases por conjunto.
    :return: Mensagem de validação.
    """
    num_frases = len(frases)
    if num_frases < itens_por_conjunto:
        return ("Erro: O número de frases únicas ({}) é menor que o número de frases por conjunto ({}).".format(num_frases, itens_por_conjunto))
    
    num_conjuntos_necessarios = math.ceil(num_frases * aparicoes_alvo / itens_por_conjunto)
    if num_conjuntos_necessarios < aparicoes_alvo:
        return ("Aviso: Poderá haver repetição insuficiente das frases. São necessários mais conjuntos ou ajustes no aparicoes_alvo.")

    return "Validação concluída: Nenhum problema encontrado."

def gerar_relatorio_validacao(mensagem_validacao, nome_arquivo="relatorio_validacao.txt"):
    """
    Gera um arquivo de texto contendo o resultado da validação.
    
    :param mensagem_validacao: Mensagem resultante da validação.
    :param nome_arquivo: Nome do arquivo para salvar o relatório de validação.
    """
    with open(nome_arquivo, 'w') as arquivo:
        arquivo.write(mensagem_validacao)
    print(f"Relatório de validação salvo como '{nome_arquivo}'.")

# Inclua as chamadas durante a execução do script principal.
frases = [
      "Próximo a escolas, supermercados e transporte público",
    "Desfrute de incríveis vistas do horizonte da cidade",
    "Salas e quartos espaçosos para maior conforto",
    "Espaço ideal para confraternizações com churrasqueira",
    "Com piscina, academia e salão de festas",
    "Materiais de alta qualidade em todos os cômodos",
    "Completa com armários embutidos de primeira linha",
    "Grandes janelas para aproveitar a luz do dia",
    "Com closet e banheira de hidromassagem",
    "Um refúgio verde no meio da cidade",
    "Porteiro 24 horas e câmeras de vigilância",
    "Duas vagas cobertas e espaço para visitantes",
    "Espaço adequado para o seu amigo de quatro patas",
    "Atualizações modernas e elegantes",
    "A poucos passos das areias brancas",
    "Sustentabilidade em primeiro lugar",
    "Conforto térmico em todos os ambientes",
    "Ambiente silencioso e bem iluminado",
    "Estrutura robusta e bem conservada",
    "Economia na conta de luz com painéis solares",
    "Diversão garantida para toda a família",
    "Acesso direto e exclusivo ao seu andar",
    "Imóvel com arquitetura charmosa e única",
    "Para momentos agradáveis com amigos",
    "Facilite seu deslocamento diário",
    "Espaço destinado para lavanderia",
    "Otimização de espaço nos dormitórios",
    "Ótima opção para relaxar ao ar livre",
    "Incentive a alimentação saudável",
    "Arquitetura contemporânea e funcional"
]

aparicoes_alvo = 3
itens_por_conjunto = 7

# Realiza a validação antes de chamar a função principal
mensagem_validacao = validar_entrada(frases, aparicoes_alvo, itens_por_conjunto)
print(mensagem_validacao)

# Gera relatório de validação
gerar_relatorio_validacao(mensagem_validacao)

# Caso a validação não retorne nenhum problema crítico, você pode prosseguir
if "Erro" not in mensagem_validacao:
    gerar_excel_para_entrevistados(frases, num_entrevistados)
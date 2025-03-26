# Instale o pacote 'openxlsx' se ainda não o tiver
if (!require("openxlsx")) install.packages("openxlsx")
library(openxlsx)

# Função para calcular o número necessário de conjuntos
calcular_conjuntos_otimizados <- function(num_frases, aparicoes_alvo = 3, itens_por_conjunto = 7) {
  num_conjuntos <- ceiling(num_frases * aparicoes_alvo / itens_por_conjunto)
  return(list(num_conjuntos = num_conjuntos, itens_por_conjunto = itens_por_conjunto))
}

# Função para gerar o Excel para entrevistados
gerar_excel_para_entrevistados <- function(frases, num_entrevistados, aparicoes_alvo = 3, nome_arquivo = 'sorteio_maxdiff.xlsx') {
  num_frases <- length(frases)
  otimizado <- calcular_conjuntos_otimizados(num_frases, aparicoes_alvo)
  num_conjuntos <- otimizado$num_conjuntos
  itens_por_conjunto <- otimizado$itens_por_conjunto
  
  codigos_frases <- paste0("F", seq_along(frases))
  todos_os_dados <- data.frame(Entrevistado = character(),
                               Conjunto = character(),
                               Codigo_Frase = character(),
                               Frase = character(),
                               stringsAsFactors = FALSE)
  
  for (id_entrevistado in seq_len(num_entrevistados)) {
    conjuntos_maxdiff <- replicate(num_conjuntos, sample(frases, itens_por_conjunto, replace = FALSE), simplify = FALSE)
    
    for (indice_conjunto in seq_along(conjuntos_maxdiff)) {
      conjunto_maxdiff <- conjuntos_maxdiff[[indice_conjunto]]
      for (frase in conjunto_maxdiff) {
        todos_os_dados <- rbind(todos_os_dados, data.frame(
          Entrevistado = paste("Entrevistado", id_entrevistado),
          Conjunto = paste("Conjunto", indice_conjunto),
          Codigo_Frase = codigos_frases[which(frases == frase)],
          Frase = frase,
          stringsAsFactors = FALSE
        ))
      }
    }
  }
  
  total_conjuntos <- num_entrevistados * num_conjuntos
  dados_resumo <- data.frame(
    "Total de Telas" = total_conjuntos,
    "Telas por Entrevistado" = num_conjuntos,
    "Frases por Tela" = itens_por_conjunto,
    stringsAsFactors = FALSE
  )
  
  # Criar o arquivo Excel
  wb <- createWorkbook()
  addWorksheet(wb, "Atribuições MaxDiff")
  writeData(wb, "Atribuições MaxDiff", todos_os_dados)
  
  addWorksheet(wb, "Resumo")
  writeData(wb, "Resumo", dados_resumo)
  
  saveWorkbook(wb, file = nome_arquivo, overwrite = TRUE)
  cat(sprintf("Arquivo Excel '%s' foi criado com conjuntos MaxDiff otimizados.\n", nome_arquivo))
}

# Função para validar a entrada
validar_entrada <- function(frases, aparicoes_alvo, itens_por_conjunto) {
  num_frases <- length(frases)
  if (num_frases < itens_por_conjunto) {
    return(sprintf("Erro: O número de frases únicas (%d) é menor que o número de frases por conjunto (%d).", 
                   num_frases, itens_por_conjunto))
  }
  
  num_conjuntos_necessarios <- ceiling(num_frases * aparicoes_alvo / itens_por_conjunto)
  if (num_conjuntos_necessarios < aparicoes_alvo) {
    return("Aviso: Poderá haver repetição insuficiente das frases. São necessários mais conjuntos ou ajustes no aparicoes_alvo.")
  }
  
  return("Validação concluída: Nenhum problema encontrado.")
}

# Função para gerar relatório de validação
gerar_relatorio_validacao <- function(mensagem_validacao, nome_arquivo = "relatorio_validacao.txt") {
  writeLines(mensagem_validacao, con = nome_arquivo)
  cat(sprintf("Relatório de validação salvo como '%s'.\n", nome_arquivo))
}

# Frases exemplo
frases <- c(
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
)

# Parametros
aparicoes_alvo <- 3
itens_por_conjunto <- 7

# Validação
mensagem_validacao <- validar_entrada(frases, aparicoes_alvo, itens_por_conjunto)
cat(mensagem_validacao, "\n")

# Gerar relatório de validação
gerar_relatorio_validacao(mensagem_validacao)

# Gerar Excel caso a validação esteja correta
num_entrevistados <- 300
if (!grepl("Erro", mensagem_validacao)) {
  gerar_excel_para_entrevistados(frases, num_entrevistados)
}
# GmailToDrive
Projeto utilizando Google Apps Script, linguagem baseada em Javascript, a fim de importar arquivos do Gmail para o Google Drive e converter para uma planilha no Google Docs.

## Problema
Em 2020, trabalhando como servidor público na Pró-Reitoria de Planejamento e Gestão da Universidade Federal de Lavras - UFLA, surgiu a necessidade de desenvolver um Dashboard para tornar a visualização dos dados orçamentários mais prática, fácil, automatizada e auxiliar na tomada de decisão pelo Pró-Reitor e Reitor nas decisões de planejamento e gestão da Universidade.

Como a UFLA já possuia na época a parceria do Google Workspace, visualizei a alternativa de utilizar o Data Studio como ferramenta para criação do Dashboard. O Data Studio permite a criação de gráficos iterativos utilizando como base dados as planilhas do Google Docs. Porém, os relatórios orçamentários da UFLA somente eram extraídos por meio do sistema do governo Tesouro Direto e a única função automática para emissão de relatórios era o envio por e-mail, a outra opção era baixando manualmente o arquivo para o computador. O que não atendia o interesse, uma vez que queríamos uma automatizar todo o processo com atualização diária.

Em meio a pesquisas encontrei o aplicativo Google Apps Script que fornece uma 

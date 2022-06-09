# GmailToDrive
Projeto utilizando Google Apps Script, linguagem baseada em Javascript, a fim de importar arquivos do Gmail para o Google Drive e converter para uma planilha no Google Docs.

## Problema
Em 2020, trabalhando como servidor público na Pró-Reitoria de Planejamento e Gestão da Universidade Federal de Lavras - UFLA, surgiu a necessidade de desenvolver um Dashboard para tornar a visualização dos dados orçamentários mais prática, fácil, automatizada e auxiliar na tomada de decisão pelo Pró-Reitor e Reitor nas decisões de planejamento e gestão da Universidade.

Como a UFLA já possuia na época a parceria do Google Workspace, visualizei a alternativa de utilizar o Data Studio como ferramenta para criação do Dashboard. O Data Studio permite a criação de gráficos iterativos utilizando como base dados as planilhas do Google Docs. Porém, os relatórios orçamentários da UFLA somente eram extraídos por meio do sistema do governo Tesouro Gerencial e a única função automática para emissão de relatórios era o envio por e-mail, a outra opção era baixando manualmente o arquivo para o computador. O que não atendia o interesse, uma vez que queríamos uma automatizar todo o processo com atualização diária do Dashboard.

Em meio a pesquisas encontrei o aplicativo Google Apps Script que fornece integração das ferramentas do Google por meio de tarefas automáticas.

## Resumo da Solução 
O script desenvolvido funciona resumidamente da seguinte forma:
1. Tesouro Gerencial envia o relatório em formato .txt para o Gmail;
2. O script importa o arquivo do e-mail para uma pasta do Google Drive e faz o backup do relatório do dia anterior, se houver;
3. Com os arquivos importados para o Drive, o script faz a conversao dos dados de cada arquivo para uma planilha única do Google Docs, para cada aba e formato de layout específico;
4. Com isso, o gráfico desenvolvido no Data Studio consegue ler os dados da planilha.

## Imagens da Solução
![image](https://user-images.githubusercontent.com/36649420/172835122-b2584faf-7d84-4b6d-9198-49efd1dfc9c2.png)

![image](https://user-images.githubusercontent.com/36649420/172835655-6eec012c-e75b-4f2c-9beb-58f5c85e2df0.png)

![image](https://user-images.githubusercontent.com/36649420/172835334-01e66c05-289c-4a97-8664-585a3fbcc0b0.png)

![image](https://user-images.githubusercontent.com/36649420/172835427-d40057ff-2f42-49f2-9326-d59eec73f39d.png)

![image](https://user-images.githubusercontent.com/36649420/172835519-2a5a3352-6c62-4f36-8f19-3786b2e15ee8.png)

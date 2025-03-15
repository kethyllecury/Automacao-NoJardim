# Automação de Relatório de Vendas - Projeto No Jardim

Este repositório contém um script de automação desenvolvido para facilitar a coleta, processamento e organização dos relatórios de vendas mensais do restaurante **No Jardim**. A solução utiliza o **Selenium** para a navegação automatizada no portal web da plataforma **Takeat**, juntamente com o **Pandas** para processamento e manipulação de dados em arquivos **CSV/Excel**. O objetivo principal é gerar relatórios de vendas de forma automatizada, combinando dados de diferentes fontes e organizando-os de acordo com o mês de venda.

## Tecnologias Utilizadas

- **Python 3.x**: Linguagem de programação utilizada para o desenvolvimento do script.
- **Selenium**: Biblioteca para automação de navegação no navegador, usada para interagir com o portal **TakeEat** e coletar dados de vendas.
- **Pandas**: Biblioteca para manipulação de dados em formato de tabelas (DataFrames), usada para editar, limpar e combinar planilhas.
- **Openpyxl**: Biblioteca usada para ler e escrever arquivos **Excel**.
- **Datetime**: Módulo utilizado para manipulação de datas e horários durante o processo de automação.
- **dotenv**: Biblioteca para gerenciar variáveis de ambiente de forma segura.
- **webdriver_manager**: Biblioteca para facilitar o gerenciamento automático do driver do navegador (Microsoft Edge).

## Funcionalidades

O script automatiza as seguintes etapas do processo de coleta e organização dos relatórios:

### 1. **Configuração do Driver**
O script configura o **Microsoft Edge** com preferências de download, garantindo que o arquivo exportado seja salvo na pasta correta, sem solicitações adicionais de download.

### 2. **Login Automático**
O script realiza o login automaticamente no portal **TakeEat** utilizando as credenciais armazenadas em um arquivo `.env`.

### 3. **Geração e Download do Relatório**
Após o login, o script navega até a seção de relatórios de vendas, seleciona o relatório do dia anterior e realiza o download do arquivo em **Excel**.

### 4. **Processamento de Dados**
O script carrega o arquivo de vendas gerado, adiciona uma nova coluna com a data do relatório e insere outras informações relevantes como o tipo de produto. A planilha é então salva com as alterações.

### 5. **Concatenação de Planilhas**
O script carrega uma planilha pré-existente com dados anteriores e concatena com o novo relatório. Ele garante que as colunas de ambas as planilhas sejam compatíveis e evita duplicações.

### 6. **Geração do Relatório Final**
Após a concatenação, o script aplica algumas regras de negócios para definir o tipo de produto (normal ou complemento) e realiza a conversão de datas para o formato **"mes-ano"** (exemplo: janeiro-2023).

### 7. **Organização e Salvamento**
O script organiza o arquivo final e o salva no caminho especificado. Caso existam arquivos antigos, eles são deletados para manter a pasta limpa.

### 8. **Deleção de Arquivos Temporários**
Após a execução, o script remove os arquivos temporários de download e os arquivos intermediários criados durante o processo.

### 9. **Criptografia**
O projeto utiliza um ambiente virtual **.venv** para isolar as dependências do projeto, garantindo um ambiente de desenvolvimento mais seguro e eficiente. Isso impede conflitos com outras bibliotecas e facilita a instalação das dependências de forma controlada.

## Conclusão

Este projeto proporciona uma solução automatizada para coletar, processar e organizar relatórios de vendas do portal **TakeEat**. Com a utilização de bibliotecas como Selenium, Pandas e Openpyxl, conseguimos criar um fluxo eficiente e seguro para automatizar tarefas repetitivas, economizando tempo e reduzindo erros manuais.

Através da configuração simples do arquivo `.env`, o script garante que o processo seja personalizado de acordo com suas credenciais e preferências. Além disso, a organização dos relatórios finais e a remoção de arquivos temporários ajudam a manter o ambiente limpo e funcional.

Com essa automação, as empresas podem melhorar a produtividade, concentrando-se em tarefas mais estratégicas enquanto o processo de coleta e organização de dados é realizado de forma eficiente e confiável.

Sinta-se à vontade para adaptar, melhorar ou expandir o script conforme suas necessidades!






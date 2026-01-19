📊 Automação de Consolidação de Banco de Horas
- Este projeto surgiu a partir de um gargalo identificado no processo de controle de banco de horas da empresa, realizado por meio de planilhas individuais preenchidas por aproximadamente 70 colaboradores. Mensalmente, a analista de RH precisava abrir manualmente cada arquivo para coletar os totais, tornando o processo lento e suscetível a erros.
- Com o objetivo de otimizar essa rotina, foi desenvolvida uma automação em Python que realiza a leitura automática da célula responsável pelo total mensal de horas.

⚙️ Como funciona
- Todas as planilhas são inseridas em uma pasta padrão chamada input
- O script percorre cada arquivo presente nessa pasta
- Para cada planilha, são extraídas as seguintes informações:
  - Nome do arquivo
  - Mês de referência
  - Total de horas
- Os dados são exibidos diretamente no terminal (CMD), facilitando a conferência e consolidação das informações

🎯 Benefícios
- Redução significativa de trabalho manual
- Agilidade na consolidação do banco de horas
- Menor chance de erros humanos
- Melhoria na produtividade do RH

🛠 Tecnologias utilizadas
- Python
- Manipulação de planilhas (Excel)

# Funções em ADVPL para o Totvs Protheus

Aqui tenho alguns códigos fontes de funções que utilizo no Totvs Protheus. Não coloco todos os fontes abertos, pois muitos tem regras de negócio exclusivos, mas funções, relatórios ou pontos de entrada que possam ser compartilhados deixarei por aqui. 

Tendo alguma contribuição ou sugestão de melhoria, fique a vontade! 😉

## Descrição dos Fontes

#### Frel0001
Relatório de Usuários do sistema em excel, com dados básicos do usuário, se está bloqueado e ultimo acesso, bom pra fazer uma inspeção se tem algum usuário que esqueceu de deletar.

#### Frel0002
Relatório para auxiliar no processo de geração de Impostos de Serviço (ISS, COFINS, PIS, IR), mostrando em uma tabela excel, os parametros, e os campos de cadastros que deveriram estar preenchidos, com ser valor atual e o valor esperado para gerar os impostos.

#### GD2excel
Biblioteca muito simples para auxiliar na geração de relatórios em excel do Marinaldo de Jesus.

#### Functions
Funções auxiliares. 
    UsrRetNome => Retorna o nome do usuário no sistema
    zCriaPar => Cria parâmetros no sistema via código. (Atilio)
    zPutSX1 => Cria Pergunte, para a versão Protheus12 que retirou o PutSX1. (Atilio)

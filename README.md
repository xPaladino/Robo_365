Inicio do Projeto em 02/02/2024

Identificado que o cliente recebe vários e-mails solicitando para lançar notas complementares ao sistema já utilizado. 
Os Fornecedores possuem diversas transportadoras, onde encaminham essas notas sem nenhum tipo de tratamento, sendo eles encaminhados por PDF, no Corpo do E-mail, Zip, Rar, Excel.
Pensado nisso foi criado esse projeto onde tenta recolher os diversos padrões e compara-los junto a uma base de dados já existente.

Exemplo:
Fornecedor A possue transportadoras A, B, C, D, E e F, onde cada um manda um PDF com uma informação na descrição, o que nos interessa é a Nota Fiscal.
Segue a NF 000 ; Segue a Nota Fiscal referente a data xxx, 000; Segue o número da NF: 000; e assim por diante.

Objetivo:
Esse código, cria um robo onde o mesmo conecta a um e-mail, que será responsável por lê-los, extrai-los e compara-los junto a um banco SQL, e por fim, armazenando em um EXCEL as informações.

Dificuldades:
Não há padrão entre os fornecedores;

# Informações sobre o formulário Gerador de Relatório</br>
</br>

# Compilado e Testado em:</br>
* Borland Delphi Studio 2006;</br>
* Embarcadero Delphi XE8.</br>
</br>

# Conexão de Banco de Dados</br>
O formulário trabalha com conexão MySQL utilizando componentes DBXpress.</br>
</br>
<strong>Componentes utilizados:</strong></br>
* TSQLConnection</br>
* TSQLDataSet</br>
* TDataSetProvider</br>
* TClientDataSet</br>
</br>

# Como incluir o formulário em seu projeto</br>
Adicione o "UfrmGeradorDeRelatorio.pas" em seu projeto e ele acrescentará também o ".dfm".</br>
Na cláusula <strong>"uses"</strong> da sua unit acrescente a unit "UfrmGeradorDeRelatorio".</br>
</br>

# Como inicializar o formulário e deixá-lo pronto para uso</br>
Ao menos uma chamada de configuração é indispensável para utilização do form. Esta chamada faz a conexão do formulário com o banco de dados.</br>
É necessário indicar ao formulário como se conectar com o banco de dados, passando o nome do driver, o nome do host, o nome da base de dados, usuário e senha.</br>
</br>
<strong>frmGeradorDeRelatorios.ConexaoBD('MySQL','dataserver','nome_do_banco_de_dados','usuario','senha');</strong></br>
</br>
</br>

# Modelos de formulários baseados em tags</br>
O gerador de relatórios precisa de um modelo de relatório pronto e salvo em uma pasta para que possa abrí-lo e gerar o relatório final.</br>
O gerador se baseia no modelo de substituição de tags encontradas no documento; realizando assim uma busca de uma tag no documento modelo e substituindo-a pelo conteúdo de um campo desejado.</br>
</br>

# Formato das TAGs</br>
As tags nos documentos modelos devem estar no formato <strong>&lt;tag&gt;</strong></br>
</br>

# Chamada do procedimento para impressão</br>
A chamada do procedimento para gerar o relatório se dá pelos seguintes procedimentos:</br>

## Chamada 1
* <strong>frmGeradorDeRelatorios.Relatorio(modelo,relatorio,titulorelatorio,sql);</strong>

Os parâmetros são:
* modelo: indica a localização e nome do arquivo modelo a ser utilizado. ex.: "c:\modelos\envelope.doc"</br>
* relatorio: indica a localização e nome do arquivo a ser criado. ex.: "c:\relatorios\envelope.doc"</br>
* titulorelatorio: indica um nome para o relatório. Somente se o modelo possuir a tag &lt;titulorelatorio&gt; este parâmetro irá substituir a tag, caso contrário este parâmetro será ignorado.</br>
* sql: este parâmetro deve conter a consulta SQL que retornará os dados desejados para o relatório. ex.: "select nome, tel, endereco from funcionario where aniversariante_do_mes(id_funcionario)"</br>
</br>

<strong>Obs.:</strong> Para esta chamada, o nome dos campos retornados pela consulta SQL devem ser os mesmos das tags do modelo.</br>
Ex.:</br>
tags do documento modelo ( &lt;nome&gt;, &lt;tel&gt;, &lt;endereco&gt; )</br>
campos retornados pela consulta ( nome, tel, endereco )</br>
</br>

## Chamada 2
* <strong>frmGeradorDeRelatorios.Relatorio(modelo,relatorio,titulorelatorio,sql,tags);</strong>

Os parâmetros são:
* modelo: indica a localização e nome do arquivo modelo a ser utilizado. ex.: "c:\modelos\envelope.doc"</br>
* relatorio: indica a localização e nome do arquivo a ser criado. ex.: "c:\relatorios\envelope.doc"</br>
* titulorelatorio: indica um nome para o relatório. Somente se o modelo possuir a tag &lt;titulorelatorio&gt; este parâmetro irá substituir a tag, caso contrário este parâmetro será ignorado.</br>
* sql: este parâmetro deve conter a consulta SQL que retornará os dados desejados para o relatório. ex.: "select nome, tel, endereco from funcionario where aniversariante_do_mes(id_funcionario)"</br>
* tags: este parâmetro deve ser um TStringList contendo uma lista das tags que devem ser procuradas no documento modelo. ex.: [&lt;tag1&gt;, &lt;tag2&gt;, &lt;tag3&gt;]
</br>

<strong>Obs.:</strong> Para esta chamada, o nome dos campos retornados pela consulta SQL não precisam ser os mesmos das tags do modelo, porém precisam ser a mesma quantidade. Se a lista de tags contiver 3 tags, a consulta SQL deve retornar 3 campos, sendo que o primeiro campo substituirá a primeira tag e assim sucessivamente.</br>
Ex.:</br>
tags do documento modelo ( &lt;tag1&gt;, &lt;tag2&gt;, &lt;tag3&gt; )</br>
campos retornados pela consulta ( nome, tel, endereco )</br>
</br>

## Chamada 3
* <strong>frmGeradorDeRelatorios.Relatorio(modelo,relatorio,titulorelatorio,sql,indice_tabela);</strong>

Os parâmetros são:
* modelo: indica a localização e nome do arquivo modelo a ser utilizado. ex.: "c:\modelos\envelope.doc"</br>
* relatorio: indica a localização e nome do arquivo a ser criado. ex.: "c:\relatorios\envelope.doc"</br>
* titulorelatorio: indica um nome para o relatório. Somente se o modelo possuir a tag &lt;titulorelatorio&gt; este parâmetro irá substituir a tag, caso contrário este parâmetro será ignorado.</br>
* sql: este parâmetro deve conter a consulta SQL que retornará os dados desejados para o relatório. ex.: "select nome, tel, endereco from funcionario where aniversariante_do_mes(id_funcionario)"</br>
* indice_tabela: indica qual tabela do modelo deverá ser usada. ex.: Caso o modelo possua 3 tabelas e a segunda deva ser utilizada, basta passar o número 2 como parâmetro.</br>
</br>

<strong>Obs.:</strong> Esta chamada é utilizada para preencher uma tabela vazia que não contenha tags; basta apenas que a quantidade de campos retornada pela consulta seja igual à quantidade de colunas da tabela indicada.</br>
Ex.:</br>
campos retornados pela consulta ( nome, tel, endereco )</br>
A tabela indicada deve estar vazia, possuir apenas 1 linha e 3 colunas.</br>
<table><tr><td></td><td></td><td></td></tr></table></br>
</br>

## Chamada 4
* <strong>frmGeradorDeRelatorios.Relatorio(modelo,relatorio,titulorelatorio,sql,indice_tabela,tags);</strong>

Os parâmetros são:
* modelo: indica a localização e nome do arquivo modelo a ser utilizado. ex.: "c:\modelos\envelope.doc"</br>
* relatorio: indica a localização e nome do arquivo a ser criado. ex.: "c:\relatorios\envelope.doc"</br>
* titulorelatorio: indica um nome para o relatório. Somente se o modelo possuir a tag &lt;titulorelatorio&gt; este parâmetro irá substituir a tag, caso contrário este parâmetro será ignorado.</br>
* sql: este parâmetro deve conter a consulta SQL que retornará os dados desejados para o relatório. ex.: "select nome, tel, endereco from funcionario where aniversariante_do_mes(id_funcionario)"</br>
* indice_tabela: indica qual tabela do modelo deverá ser usada. ex.: Caso o modelo possua 3 tabelas e a segunda deva ser utilizada, basta passar o número 2 como parâmetro.</br>
* tags: este parâmetro deve ser um TStringList contendo uma lista das tags que devem ser procuradas no documento modelo. ex.: [&lt;tag1&gt;, &lt;tag2&gt;, &lt;tag3&gt;]</br>
</br>

<strong>Obs.:</strong> Esta chamada é utilizada para preencher uma tabela que contenha tags; a quantidade de campos retornada pela consulta não precisa ser a mesma quantidade de colunas da tabela indicada. Esta chamada é indicada para situações onde uma coluna deva conter mais de uma tag.</br>
Ex.:</br>
tags do documento modelo ( &lt;tag1&gt;, &lt;tag2&gt;, &lt;tag3&gt; )</br>
campos retornados pela consulta ( nome, tel, endereco )</br>
A tabela indicada poderia conter, por exemplo, uma tabela com duas colunas, sendo que na primeira constaria o nome do funcionário e na segunda o telefone e endereço.</br>
<table><tr><td valign=top>&lt;tag1&gt;</td><td>&lt;tag2&gt;</br>&lt;tag3&gt;</td></tr></table></br>
</br>

# Tipos de relatórios que podem ser gerados</br>
O gerador de relatórios está preparado para gerar 4 tipos de relatórios, ou seja, relatórios com estruturas diferentes... vamos à elas:</br>
</br>
## Modelo 1 - ver acima *Chamadas 1 e 2</br>
Seria um relatório, de uma ou mais páginas, que contém as tags a serem substituídas pelas informações de <strong>um único registro</strong>.</br>
Caso tenham mais registros a serem impressos, o gerador adiciona uma quebra de página, e vai "repetindo" o conteúdo do modelo com as informações dos próximos registros.</br>

Ex.: Ficha Cadastral de um ou mais funcionários de uma empresa.</br>
</br>

## Modelo 2 - ver acima *Chamada 3</br>
Este modelo realiza o preenchimento de uma tabela vazia; bastando que exista uma tabela, de apenas uma linha, e a quantidade de colunas e de campos da consulta sejam a mesma.</br>
O gerador irá preencher a primeira linha da tabela e adicionar linhas à tabela na mesma quantidade de registros retornados pela consulta.</br>

Ex.: Relação de funcionários de uma empresa, contendo nome, telefone e endereço.</br>
Neste exemplo o modelo de relatório precisa ter uma tabela vazia de 1 linha e 3 colunas.</br>
<table><tr><td></td><td></td><td></td></tr></table></br>
</br>

## Modelo 3 - ver acima *Chamada 4</br>
Este modelo realiza o preenchimento de uma tabela com tags, onde a tabela deve conter apenas 1 linha com as tags e a quantidade de colunas não precisa corresponder à quantidade de campos retornados pela consulta.</br>
O gerador irá substituir as tags da linha, com os respectivos valores dos campos, e adicionar linhas à tabela na mesma quantidade de registros retornados pela consulta.</br>

Ex.: Relação de funcionários de uma empresa, contendo nome, telefone e endereço.</br>
Neste exemplo o modelo de relatório pode ter, por exemplo, uma tabela de 1 linha e 2 colunas com as tags a serem substituídas.</br>
<table><tr><td valign=top>&lt;tag1&gt;</td><td>&lt;tag2&gt;</br>&lt;tag3&gt;</td></tr></table></br>
</br>


# FAQ</br>
</br>
1. Por que o formulário Gerador de Relatório possui seu próprio componente TSQLConnection?</br>
R. A ideia do formulário possuir sua própria conexão é dar mais liberdade ao form evitando erros de acesso ao banco.</br>
Se o usuário logado no seu sistema não possui permissão de 'select' em uma tabela da qual ele irá imprimir um relatório, o form também não teria permissão de leitura e causaria erro de acesso ao banco.</br>
O ideal é criar um usuário em seu banco, exclusivo para este form e dar permissão de leitura em todas as tabelas do banco; ao inicializar o formulário, passe este usuário e senha como parâmetros.</br>
</br>

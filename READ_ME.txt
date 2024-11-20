_________________________________________________________________________________________________________

----------AUTOMAÇÃO DE ENVIO AUTOMÁTICO DE MENSAGENS VIA WHATSAPP----------
_________________________________________________________________________________________________________

## Sobre o projeto: 

Esse projeto tem a finalidade de facilitar a comunicação entre o colaborador
na ponta operação com o nossos setores indiretos.

O projeto consiste em uma página web streamlit integrada a um teste de software automatizado que manipula o whatsapp 
no navegardor Chrome.

Para utilizar o projeto você deverá informar o id da planilha no smartsheet e se vai algum documento após a mensagem.


_________________________________________________________________________________________________________

## Estrutura de dados:

Dentro do smartsheet as 4 primeiras colunas são padrãom, todas as demais podem ser usadas como valores
variaveis para compor a mensagem 

*Pontos de atenção: 
	*Coluna CPF: deve conter apenas números;
	*Coluna Status: NÃO deve ser preenchida manualmente;
	*Coluna Caminho_arquivo: Todo arquivo enviado deve estar dentro do diretório -> C:\Users\seu_user\gpssa.com.br\Dashboard_Brian_Silva - Documentos\05 - AUTOMAÇÕES\Envio automático de mensagens - Whatsapp\FILES
	*Coluna Caminho_arquivo: O caminho não deve estar entre áspas.

_________________________________________________________________________________________________________

## Requisitos

- **Python 3.7+**: Certifique-se de que você tem o Python instalado.
- **Streamlit**: Para criar a interface web.
- **WebDriver do Chrome**: Para automação com o navegador Chrome.
- **Bibliotecas adicionais**: `pandas`, `openpyxl`, `selenium`, entre outras necessárias para a automação e manipulação de dados.

1. **Configuração Inicial**:
 - Instale as dependências necessárias listadas no arquivo `requirements.txt` usando o comando:
   pip install -r requirements.txt

 - Configure o WebDriver do Chrome e certifique-se de que está acessível no seu PATH.

2. **Execução da Página Web**:
 - Navegue até o diretório do projeto e inicie o aplicativo Streamlit com o comando:
   ```bash
   streamlit run app.py
   ```
 - A página será aberta no navegador, onde você deverá informar o ID da planilha no Smartsheet e se vai enviar algum documento após a mensagem.




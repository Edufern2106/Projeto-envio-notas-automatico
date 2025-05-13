# Envio Automático de Arquivos por E-mail via Outlook

Este sistema automatiza o envio de documentos por e-mail, agrupando anexos por cliente e personalizando o assunto do e-mail conforme o nome do cliente e o mês de envio.

## Como funciona
- Escaneia uma pasta (e subpastas) em busca de arquivos no padrão:
  `Número da Nota - Nome do Cliente - Tipo de Documento.ext`
- Agrupa arquivos por cliente (nome extraído do arquivo).
- Usa um dicionário para mapear clientes para e-mails.
- Envia e-mail via Outlook para cada cliente, com anexos, assunto personalizado e corpo formal.
- Loga sucesso/erro por cliente.

## Requisitos
- Python 3.x
- Outlook instalado e configurado
- Instalar dependências:
  ```bash
  pip install -r requirements.txt
  ```

## Como usar
1. Ajuste o dicionário de clientes/e-mails no script `send_documents.py`.
2. Configure o caminho da pasta a ser escaneada.
3. Execute o script:
   ```bash
   python send_documents.py
   ```
4. Verifique o log no console.

## Como criar o executável (.exe)
1. Instale as dependências:
   ```bash
   pip install -r requirements.txt
   ```
2. Gere o executável com:
   ```bash
   pyinstaller --onefile send_documents.py
   ```
3. O executável estará na pasta `dist` como `send_documents.exe`.

## Como criar o botão no Excel para enviar e-mails
1. Abra o Excel e pressione `ALT+F11` para abrir o Editor VBA.
2. No menu, clique em `Inserir > Módulo`.
3. Cole o seguinte código (ajuste o caminho para o seu .exe):
   ```vba
   Sub EnviarEmailsPython()
       Dim Ret_Val
       Ret_Val = Shell("C:\\Users\\eduardo.fernandes\\CascadeProjects\\envio_arquivos_email_outlook\\dist\\send_documents.exe", 1)
   End Sub
   ```
4. Volte ao Excel, insira um botão (Desenvolvedor > Inserir > Botão) e vincule à macro `EnviarEmailsPython`.
5. Clique no botão para disparar o envio dos e-mails.

## Observações
- O Outlook será usado para enviar os e-mails, então a conta deve estar configurada no computador.
- O script pode ser adaptado para rodar como executável (usando pyinstaller, por exemplo).

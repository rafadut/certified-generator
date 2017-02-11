# GeradorDeCertificados


O gerador funciona assim:

1 - Realiza a leitura de uma planilha com nomes e emails dos alunos que receberão os certificados;
2 - Copia uma imagem modelo para montar os certificados;
3 - Gera um certificado para cada nome da lista dentro da pasta "C:\GeradorDeCertificados\gerados";
4 - Envia o certificado por email (opcional).


Para utilizá-lo:

1 - Copie a pasta "GeradorDeCertificados" para o seu "C:"
Vai ficar assim: "C:\GeradorDeCertificados"
Essa pasta contêm uma lista de nomes e uma imagem de exemplos.

2 - Instale o programa "AccessDatabaseEngine.exe", que está na pasta "AccessDatabaseEngine". Ele serve para que o componente consiga fazer a leitura da planilha;

3 - Execute o arquivo "GeradorDeCertificados.exe";

Após a geração dos certificados, o prompt vai perguntar se você deseja enviar os certificados por email, na versão atual é possível enviar a partir de um email Microsoft (@hotmail.com, @live.com, @outlook.com) ou Google (@gmail.com). Para enviar pelo gmail é preciso liberar o acesso na sua conta do Google (http://stackoverflow.com/questions/29465096/how-to-send-an-e-mail-with-c-sharp-through-gmail/29465275).

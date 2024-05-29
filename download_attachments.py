import imaplib
import email
import os
import re
import datetime

# Conectando ao Gmail com IMAP
objCon = imaplib.IMAP4_SSL("imap.gmail.com")

# Login e senha
login = "email"
senha = "senha"

objCon.login(login, senha)

objCon.list()
objCon.select(mailbox='inbox', readonly=True)
print("Login realizado com sucesso!")


# Filtros
data_inicio = datetime.datetime(2024, 3, 11).strftime('%d-%b-%Y')
data_final = datetime.datetime(2024, 3, 18).strftime('%d-%b-%Y')

respostas, idDosEmails = objCon.search(None, f'SINCE {data_inicio} BEFORE {data_final}')

# Salvar na pasta
if not os.path.exists("anexos"):
    os.makedirs("anexos")

# Loopando cada ID do email na caixa de entrada
for num in idDosEmails[0].split():

    try:
        # Decodificando o email e jogando em variáveis
        resultado, dados = objCon.fetch(num, '(RFC822)')
        texto_do_email = dados[0][1]
        texto_do_email = texto_do_email.decode('utf-8')
        texto_do_email = email.message_from_string(texto_do_email)

        # Loopando as partes do email
        for part in texto_do_email.walk():

            # Se tiver anexo, pegar nome do arquivo
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            # Pegando o nome do anexo
            fileName = part.get_filename()

            # Limpe o nome do arquivo substituindo caracteres inválidos por underscores
            if fileName is not None:
                cleaned_fileName = re.sub(r'[\/:*?"<>|]', '_', fileName)  # Substitua caracteres inválidos por underscores
                # Adicione um contador ao nome do arquivo para torná-lo único
                base_name, ext = os.path.splitext(cleaned_fileName)
                counter = 1
                while os.path.exists(f"carol/{base_name}_{counter}{ext}"):
                    counter += 1
                unique_fileName = f"carol/{base_name}_{counter}{ext}"

                # Substituindo caracteres de quebra de linha e retorno de carro no nome do arquivo
                unique_fileName = unique_fileName.replace('\r', '').replace('\n', '')

                arquivo = open(unique_fileName, 'wb')

                # Escrevendo binário
                arquivo.write(part.get_payload(decode=True))
                arquivo.close()
            else:
                print("Nenhum nome de arquivo encontrado para este anexo.")

    except Exception as e:
        print(f"Ocorreu um erro ao processar o e-mail: {e}")
        continue

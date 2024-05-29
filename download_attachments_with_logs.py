import imaplib
import email
import os
import re
import datetime
from colorama import Fore

# criando conexão ao Gmail com IMAP
objCon = imaplib.IMAP4_SSL("imap.gmail.com")

# login e senha
login = ""
senha = ""

objCon.login(login, senha)

objCon.list()
objCon.select(mailbox='inbox', readonly=True)
print("Logado com sucesso!")

# data selecionada
data_inicio = datetime.datetime(2024, 4, 1).strftime('%d-%b-%Y')
data_final = (datetime.datetime(2024, 5, 29) + datetime.timedelta(days=1)).strftime('%d-%b-%Y')

# respostar e id dos e-mails buscando pela data selecionada
respostas, idDosEmails = objCon.search(None, f'SINCE {data_inicio} BEFORE {data_final}')

# pasta onde será salvo
if not os.path.exists("attachments"):
    os.makedirs("attachments")

# iterando cada ID do email na caixa de entrada
for num in idDosEmails[0].split():
    try:
        # decodificando o email e jogando em variáveis
        resultado, dados = objCon.fetch(num, '(RFC822)')
        texto_do_email = dados[0][1]

        try:
            texto_do_email = email.message_from_bytes(texto_do_email)
        except Exception as e:
            print(Fore.RED + f"Erro ao decodificar e-mail ID {num}: {e}")
            continue

        # extraindo o remetente e o assunto para depuração
        remetente = texto_do_email.get('From')
        assunto = texto_do_email.get('Subject')
        print(Fore.GREEN + f"Processando e-mail ID {num} de {remetente} com assunto: {assunto}")

        # iterando dentro do email
        for part in texto_do_email.walk():
            try:
                # se tiver anexo, pegar nome do arquivo
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                # pegando o nome do anexo
                fileName = part.get_filename()

                # limpando o nome do arquivo e substituindo caracteres inválidos por underscores
                if fileName is not None:
                    cleaned_fileName = re.sub(r'[\/:*?"<>|]', '_', fileName)  # Substitua caracteres inválidos por underscores
                    # adicionando um contador ao nome do arquivo para torná-lo único
                    base_name, ext = os.path.splitext(cleaned_fileName)
                    counter = 1
                    while os.path.exists(f"attachments/{base_name}_{counter}{ext}"):
                        counter += 1
                    unique_fileName = f"attachments/{base_name}_{counter}{ext}"

                    # substituir caracteres de quebra de linha e retorno de carro no nome do arquivo
                    unique_fileName = unique_fileName.replace('\r', '').replace('\n', '')

                    with open(unique_fileName, 'wb') as arquivo:
                        # escrevendo binário
                        arquivo.write(part.get_payload(decode=True))
                    print(Fore.LIGHTGREEN_EX + f"Anexo salvo como {unique_fileName}")
                else:
                    print(Fore.RED + "Nenhum nome de arquivo encontrado para este anexo.")
            except Exception as e:
                print(Fore.RED + f"Erro ao processar anexo do e-mail ID {num}: {e}")

    except Exception as e:
        print(Fore.RED + f"Ocorreu um erro ao processar o e-mail ID {num}: {e}")
        continue

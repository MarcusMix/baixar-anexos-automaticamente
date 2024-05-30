# Baixar anexos automaticamente do e-mail.

Este repositório contém um script Python para baixar todos os anexos de um email.

## Instalação

Para usar o código, siga os passos abaixo:

1. Clone o repositório:

    ```bash
    git clone https://github.com/MarcusMix/baixar-anexos-automaticamente
    ```

2. Navegue até o diretório do projeto:

    ```bash
    cd baixar-anexos-automaticamente
    ```

3. Instale as dependências necessárias (requer `pip`):

    ```bash
    pip install imaplib
    pip install colorama
    ```

## Uso

1. Abra sua IDE e abra o terminal.
2. Execute o script para baixar os anexos.

    ```bash
    python download_attachments_with_logs.py
    ```

## Fluxo

O script segue o fluxo abaixo:

1. **Efetua login no e-mail**: O script acessa o e-mail, e efetua o login.
2. **Iteração de cada e-mail**: Itera cada e-mail procurando os anexos.
3. **Download dos anexos**: Se existir anexos, o script baixará todos e guardará na pasta indicada.

![Fluxo do Processo](https://i.imgur.com/0LnJx5e.png)


## Licença

Este projeto está licenciado sob a licença MIT. Veja o arquivo [LICENSE](https://choosealicense.com/licenses/mit/) para mais detalhes.

### Autor

Desenvolvido por MarcusMix

---

### Referências

- [Documentação do Python](https://docs.python.org/3/)
- [Pandas Library](https://pandas.pydata.org/)
- [openpyxl Library](https://openpyxl.readthedocs.io/en/stable/)
- [IMAP lib](https://openpyxl.readthedocs.io/en/stable/)

---


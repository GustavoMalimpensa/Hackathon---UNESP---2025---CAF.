<table>
  <tr>
    <td><h1> Assistente Virtual SABO</h1></td>
    <td><img src="./sabo-.png" alt="Logo do SABO" width="45" style="margin-left:10px; vertical-align:middle;"></td>
  </tr>
</table>


O programa é uma automação para envio de mensagens no WhatsApp utilizando Python. O código lê dados de clientes a partir de uma planilha Excel, gera mensagens personalizadas e as envia via WhatsApp Web.

## Funcionalidades

- Abre o navegador Chrome e acessa o WhatsApp Web.
- Lê dados dos clientes a partir de uma planilha Excel (`clientes.xlsx`).
- Gera mensagens personalizadas incluindo o nome do cliente.
- Codifica a mensagem em uma URL e a abre no navegador.
- Localiza e clica no botão de envio usando `pyautogui`.

## Tecnologias Utilizadas

- `webbrowser` - Para abrir o navegador e acessar URLs.
- `openpyxl` - Para ler dados da planilha Excel.
- `urllib.parse` - Para codificar mensagens em URLs.
- `time` - Para adicionar delays no processo de automação.
- `pyautogui` - Para localizar e clicar em elementos na tela.

## Como Usar

1. **Instale as dependências necessárias**:
    ```sh
    pip install openpyxl pyautogui
    ```

2. **Prepare sua planilha Excel**:
    - Crie um arquivo `clientes.xlsx` com os dados dos clientes.
    - A planilha deve ter pelo menos duas colunas: `Nome` e `Telefone`.
    - Certifique-se de que a planilha possua uma **aba (guia)** chamada **`basedados`**, pois o script acessa os dados por esse nome.  

3. **Após criar a planilha**:
    - Coloque o arquivo **`clientes.xlsx`** na **mesma pasta onde está o script principal**.

4. **Execute o script**:
    - Certifique-se de que o WhatsApp esteja logado no navegador.
    - Execute o script Python para iniciar o processo de envio de mensagens.

    ---

## 🧾 License

This project is licensed under the [MIT License](LICENSE) © 2025 Solution Tech.


# relatorioBT

RelatorioBT é uma automação em Python para geração e envio de relatórios diários e agendados ao final do expediente, focada no acompanhamento de vendas e pedidos pendentes de veículos (carros e caminhões). O robô faz integração com o Power BI, trata dados das planilhas via pandas e envia os resultados resumidos via WhatsApp para todos os contatos cadastrados, promovendo agilidade e precisão na comunicação dos resultados do dia.

---

## Sumário

- [Funcionalidades](#funcionalidades)
- [Fluxo de funcionamento](#fluxo-de-funcionamento)
- [Tecnologias utilizadas](#tecnologias-utilizadas)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como executar](#como-executar)
- [Estrutura dos arquivos](#estrutura-dos-arquivos)
- [Personalização](#personalização)
- [Contribuição](#contribuição)
- [Licença](#licença)

---

## Funcionalidades

1. **Coleta automática de dados do Power BI**
   - Acessa o portal do Power BI e baixa planilhas de vendas e pedidos pendentes de carros/caminhões.

2. **Processamento de dados**
   - Utiliza `pandas` para ler, tratar e consolidar informações das planilhas Excel.
   - Gera textos predefinidos, padronizando o relatório diário/mensal.

3. **Envio automatizado via WhatsApp**
   - Utiliza automação para acessar o WhatsApp Web.
   - Envia o relatório para todos os contatos cadastrados na planilha `ContatosGEST.xlsx`.

---

## Fluxo de funcionamento

1. **Login**
   - O robô realiza login no portal do Power BI utilizando credenciais armazenadas em variáveis de ambiente.

2. **Download das planilhas**
   - Baixa planilhas de vendas e pendentes referentes a carros e caminhões.

3. **Tratamento e geração do texto**
   - Processa abas específicas das planilhas, extrai totais e valores, e monta o texto do relatório.

4. **Envio dos relatórios**
   - Abre o WhatsApp Web, pesquisa cada contato na lista, e envia o relatório consolidado.

5. **Limpeza**
   - Após o envio, exclui os arquivos das planilhas baixadas para evitar acúmulo.

---

## Tecnologias utilizadas

- **Python 3**
- [BotCity Automation Framework](https://github.com/botcity-dev/botcity-framework)
- **Pandas**
- **OpenPyXL**
- **dotenv**
- **Maestro SDK (opcional)**

---

## Pré-requisitos

- Python 3.7+
- Google Chrome (para automação do WhatsApp Web)
- Driver de automação configurado (ver BotCity documentação)
- Conta válida no portal do Power BI
- WhatsApp Web ativo nos contatos desejados
- Acesso ao arquivo `ContatosGEST.xlsx` com os contatos de envio

---

## Instalação

1. Clone o repositório:

    ```bash
    git clone https://github.com/lucas-rcalves/relatorioBT.git
    cd relatorioBT
    ```

2. Instale as dependências:

    ```bash
    pip install -r requirements.txt
    ```

---

## Configuração

1. **Arquivo `.env`**  
   Crie um arquivo `.env` na raiz do projeto com suas credenciais:

    ```
    USUARIO=seu_usuario
    SENHA=sua_senha
    ```

2. **Contatos**
   - Edite a planilha `ContatosGEST.xlsx` para incluir os nomes dos contatos do WhatsApp que receberão o relatório (coluna chamada `contato`).

---

## Como executar

```bash
python relatorioBT.py
```

O robô irá executar todas as etapas automaticamente, desde o login até o envio dos relatórios.

---

## Estrutura dos arquivos

```text
relatorioBT/
├── arquivosEXCEL/              # Pasta temporária para planilhas baixadas
├── ContatosGEST.xlsx           # Planilha de contatos para envio
├── relatorioBT.py              # Script principal
├── .env                        # Variáveis de ambiente (não versionado)
├── README.md                   # Este arquivo
└── requirements.txt            # Dependências do projeto
```

---

## Personalização

- **Template do relatório:**  
  Os textos enviados pelo WhatsApp podem ser personalizados editando as strings formatadas no código.

- **Abas e colunas das planilhas:**  
  Se sua planilha tiver estrutura diferente, ajuste os nomes das abas e índices das colunas no código.

---

## Contribuição

Contribuições são bem-vindas!  
Abra uma issue ou um pull request com sugestões ou correções.

---

## Licença

Este projeto está sob a licença MIT.

---

## Contato

Dúvidas ou sugestões?  
Entre em contato pelo [GitHub](https://github.com/lucas-rcalves) ou via WhatsApp.

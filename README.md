# RelatórioBT - Automação de Relatórios de Vendas

## Descrição

O **RelatórioBT** é uma automação desenvolvida em Python que coleta dados de vendas do Power BI, processa as informações e envia relatórios diários personalizados via WhatsApp para uma lista de contatos cadastrados.

### Funcionalidades Principais

- **Extração Automatizada**: Conecta-se ao portal Power BI e baixa planilhas Excel com dados de vendas
- **Processamento de Dados**: Analisa as informações de vendas diárias e mensais para carros e caminhões
- **Relatórios Personalizados**: Gera relatórios formatados com métricas de vendas e pedidos pendentes
- **Envio Automático**: Distribui os relatórios via WhatsApp para todos os contatos cadastrados

## Estrutura do Projeto

```
RelatorioBT/
├── RelatorioBT/
│   ├── bot.py                 # Script principal da automação
│   ├── requirements.txt       # Dependências do projeto
│   ├── ContatosGEST.xlsx     # Lista de contatos para WhatsApp
│   ├── resources/            # Imagens para reconhecimento de elementos
│   │   ├── BaixarPlanilha.png
│   │   ├── BtExportar.png
│   │   ├── Lupinha.png
│   │   └── ...
│   ├── build.bat             # Script de build para Windows
│   ├── build.ps1             # Script de build PowerShell
│   └── build.sh              # Script de build para Linux/Mac
└── README.md                 # Este arquivo
```

## Pré-requisitos

- **Python 3.8+**
- **Google Chrome** (para automação web)
- **WhatsApp Web** configurado
- Acesso ao portal Power BI (portalanalysisbi.com)
- **Windows** (recomendado, devido aos caminhos específicos)

## Instalação

### 1. Clone o Repositório
```bash
git clone https://github.com/lucas-rcalves/RelatorioBT.git
cd RelatorioBT
```

### 2. Instale as Dependências
```bash
cd RelatorioBT
pip install -r requirements.txt
```

### 3. Configuração das Variáveis de Ambiente

Copie o arquivo de exemplo e configure suas credenciais:

```bash
cp .env.example .env
```

Edite o arquivo `.env` com suas credenciais do Power BI:

```env
USUARIO=seu_usuario_powerbi
SENHA=sua_senha_powerbi
```

### 4. Configure a Lista de Contatos

Edite o arquivo `ContatosGEST.xlsx` com os contatos que devem receber os relatórios:

| contato |
|---------|
| Nome do Contato 1 |
| Nome do Contato 2 |
| ... |

## Como Usar

### Execução Manual
```bash
cd RelatorioBT
python bot.py
```

### Execução via Build Scripts

**Windows (Batch):**
```cmd
build.bat
```

**Windows (PowerShell):**
```powershell
.\build.ps1
```

**Linux/Mac:**
```bash
chmod +x build.sh
./build.sh
```

## Fluxo de Funcionamento

1. **Login no Power BI**: A automação acessa o portal com as credenciais configuradas
2. **Coleta de Dados**:
   - Extrai planilhas de vendas de carros
   - Extrai planilhas de pedidos pendentes de carros
   - Muda para o portal de caminhões
   - Extrai planilhas de vendas de caminhões
   - Extrai planilhas de pedidos pendentes de caminhões
3. **Processamento**: Analisa os dados e gera relatório formatado
4. **Envio**: Distribui o relatório via WhatsApp para todos os contatos

## Formato do Relatório

O relatório gerado inclui:

```
*INFORMATIVO DE VENDAS VEÍCULOS*

Venda diária - DD/MM/AAAA:
● VDI: X
● VN: X
● VU: X
● Total: X

Vendas Mensal - MM/AAAA:
● VDI: X
● VN: X
● VU: X
● Total: X

Total de pedidos pendentes:
● Volks: X
● Renault: X
● GM: X
● Citroen: X
● Peugeot: X
● Ford: X
● Total: X
● Valor: R$ X.XXX,XX

*CAMINHÕES*
[Dados similares para caminhões]
```

## Configurações Avançadas

### Caminhos de Arquivo
Por padrão, o sistema usa os seguintes caminhos:
- **Downloads**: `C:\Users\adtsa\Downloads\`
- **Destino**: `C:\Users\adtsa\PycharmProjects\RelatorioBT\arquivosEXCEL\`

Para alterar esses caminhos, edite as variáveis no arquivo `bot.py`.

### Timeout e Esperas
Os tempos de espera podem ser ajustados nas funções conforme necessário:
- Login: 5 segundos
- Download: 10 segundos
- WhatsApp: 15 segundos

## Dependências

- `botcity-framework-core`: Framework principal para automação
- `botcity-maestro-sdk`: SDK para integração com BotCity Maestro
- `pandas`: Processamento de dados Excel
- `python-dotenv`: Gerenciamento de variáveis de ambiente
- `openpyxl`: Leitura de arquivos Excel

## Solução de Problemas

### Elemento não encontrado
- Verifique se as imagens na pasta `resources` estão atualizadas
- Ajuste o valor de `matching` nas funções `bot.find()`

### Erro de login
- Confirme as credenciais no arquivo `.env`
- Verifique se o portal Power BI está acessível

### Falha no WhatsApp
- Certifique-se de que o WhatsApp Web está logado
- Verifique os nomes dos contatos na planilha

### Arquivo não encontrado
- Confirme os caminhos de diretório no código
- Verifique permissões de escrita na pasta de destino

## Contribuição

1. Faça um fork do projeto
2. Crie uma branch para sua feature (`git checkout -b feature/nova-feature`)
3. Commit suas mudanças (`git commit -am 'Adiciona nova feature'`)
4. Push para a branch (`git push origin feature/nova-feature`)
5. Abra um Pull Request

## Licença

Este projeto é de uso interno. Entre em contato com o desenvolvedor para mais informações sobre licenciamento.

## Suporte

Para dúvidas ou problemas:
- Abra uma issue no GitHub
- Entre em contato com o desenvolvedor do projeto

---

**Nota**: Este projeto foi desenvolvido para automatizar processos internos específicos. Certifique-se de ter as permissões necessárias antes de usar em ambiente de produção.
# TextFormatter Pro — Add-in para Word for Mac

## Funções

| Função | Descrição |
|--------|-----------|
| **Formatar Parágrafos** | Recuo de 1,25cm + Espaçamento 1,5 |
| **Formatar Citações** | Recuo 1cm esq./dir. + Espaçamento Simples |
| **Aplicar Fonte** | Helvetica Neue 12pt em todo o documento |
| **Sinônimos** | Busca sinônimos em português via IA |
| **Análise de Linha** | Detecta espaço restante na linha (máx. 65 chars) |
| **Sugestão de Preenchimento** | Sugere substituições para completar a linha exatamente |

---

## Instalação — Word for Mac 16.89.1

### Pré-requisitos
- Word for Mac versão 16.89.1 ou superior  
- Node.js 18+ (para servidor local durante desenvolvimento)
- macOS 12+

---

### Opção 1 — Servidor Local (Desenvolvimento)

**1. Instale as dependências**
```bash
cd word-addin
npm install -g http-server
```

**2. Inicie o servidor HTTPS local**
```bash
# Gere certificado autoassinado (necessário para Office.js)
openssl req -x509 -newkey rsa:4096 -keyout key.pem -out cert.pem -days 365 -nodes \
  -subj "/CN=localhost"

# Suba o servidor
http-server . -p 3000 --ssl --cert cert.pem --key key.pem -c-1
```

**3. Confie no certificado (Mac)**
```bash
sudo security add-trusted-cert -d -r trustRoot -k /Library/Keychains/System.keychain cert.pem
```

**4. Registre o add-in no Word**
- Abra o Word
- Menu: **Inserir → Suplementos → Meus Suplementos → Carregar Suplemento**
- Selecione o arquivo `manifest.xml`

---

### Opção 2 — Pasta de Catálogo Compartilhado (Mais Simples)

**1. Crie uma pasta de catálogo**
```bash
mkdir -p ~/Office-Addins/TextFormatter
cp -r /caminho/para/word-addin/* ~/Office-Addins/TextFormatter/
```

**2. Edite o manifest.xml**
Altere todas as URLs `https://localhost:3000/` para o caminho local ou seu servidor:
```xml
<SourceLocation DefaultValue="file:///Users/SEU_USUARIO/Office-Addins/TextFormatter/taskpane.html"/>
```

**3. Configure o Word para usar a pasta**
- Word → **Preferências → Segurança** → Catálogos de Suplementos Confiáveis
- Adicione o caminho: `file:///Users/SEU_USUARIO/Office-Addins/TextFormatter`
- Marque **Mostrar no Menu**

**4. Carregue o suplemento**
- **Inserir → Suplementos → Meus Suplementos** → encontre "TextFormatter Pro"

---

### Opção 3 — Deploy em Servidor (Produção)

Hospede os arquivos em qualquer servidor HTTPS:
```
https://seusite.com/textformatter/
  ├── taskpane.html
  ├── commands.html
  └── assets/
      ├── icon-16.png
      ├── icon-32.png
      └── icon-80.png
```

Atualize as URLs no `manifest.xml` e distribua o manifest via:
- **Microsoft 365 Admin Center** (SharePoint App Catalog)
- **Inserir → Suplementos → Meus Suplementos → Carregar** (individual)

---

## Estrutura de Arquivos

```
word-addin/
├── manifest.xml          ← Manifesto do add-in
├── taskpane.html         ← Interface principal (tudo em um arquivo)
├── commands.html         ← Suporte a comandos de faixa de opções
├── README.md             ← Este arquivo
└── assets/
    ├── icon-16.png
    ├── icon-32.png
    └── icon-80.png
```

---

## Uso das Funções

### Formatar Parágrafos
Clique sem seleção para aplicar ao documento inteiro, ou selecione trechos específicos.

### Formatar Citações
**Selecione o bloco de citação** antes de clicar — o botão aplica recuo e espaçamento simples apenas na seleção.

### Sinônimos
- Selecione uma palavra no documento → ela aparece automaticamente no campo
- Clique **Buscar** → sinônimos aparecem como chips clicáveis
- Clique num sinônimo para substituir no texto

### Análise de Linha + Preenchimento
1. Selecione um parágrafo ou posicione o cursor
2. Clique **Analisar Linha Atual**
3. Se sobrar espaço (2–30 chars), aparece o botão **Sugerir Palavras para Preencher**
4. Sugestões aparecem como chips — clique para substituir automaticamente

---

## Notas Técnicas

- **65 caracteres** por linha é o valor configurado para Helvetica Neue 12pt com margem padrão ABNT
- **1,25cm** de recuo = 35,4pt (conversão utilizada na API do Office.js)
- **1cm** de recuo = 28,35pt
- Espaçamento 1,5 = `lineSpacingRule: multiple` + `lineSpacing: 18pt`
- A sugestão de sinônimos usa a API Claude (claude-sonnet-4-20250514) via `fetch`

---

## Requisitos de Rede

O add-in faz chamadas para:
- `https://appsforoffice.microsoft.com` — biblioteca Office.js
- `https://api.anthropic.com/v1/messages` — sugestões de sinônimos (IA)

> Para uso offline ou corporativo, substitua as chamadas à API Claude por um dicionário local ou proxy interno.

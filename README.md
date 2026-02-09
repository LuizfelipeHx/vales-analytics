# Vales Analytics Mobile - PWA

Aplicativo mobile (PWA) para análise de vales físicos com sincronização via OneDrive.

## 📁 Estrutura

```
vales-analytics-mobile/
├── index.html      # Interface do app
├── styles.css      # Estilos responsivos
├── app.js          # Lógica principal
├── manifest.json   # Config PWA
├── sw.js           # Service Worker
├── icons/          # Ícones do app
└── README.md       # Este arquivo
```

## 🚀 Como Publicar no GitHub Pages

### 1. Criar Repositório
1. Acesse [github.com](https://github.com)
2. Clique em **"New repository"**
3. Nome: `vales-analytics`
4. Deixe público
5. Clique em **"Create repository"**

### 2. Fazer Upload dos Arquivos
1. Na página do repositório, clique em **"uploading an existing file"**
2. Arraste todos os arquivos desta pasta
3. Clique em **"Commit changes"**

### 3. Ativar GitHub Pages
1. Vá em **Settings** → **Pages**
2. Em **Source**, selecione: `Deploy from a branch`
3. Em **Branch**, selecione: `main` e `/ (root)`
4. Clique em **Save**
5. Aguarde alguns minutos

### 4. Acessar o App
- URL: `https://SEU_USUARIO.github.io/vales-analytics`

## 📱 Como Usar no Celular

### Primeira Configuração
1. Acesse a URL do GitHub Pages no navegador do celular
2. Clique em **"Configurar"**
3. Cole o link do Excel compartilhado do OneDrive
4. Clique em **"Salvar e Carregar"**

### Instalar como App
**Android (Chrome):**
1. Acesse o site
2. Toque nos 3 pontos (⋮)
3. Selecione **"Adicionar à tela inicial"**
4. Confirme

**iPhone (Safari):**
1. Acesse o site
2. Toque no ícone de compartilhar (□↑)
3. Selecione **"Adicionar à Tela de Início"**

## 🔗 Como Pegar o Link do OneDrive

1. Acesse [onedrive.live.com](https://onedrive.live.com)
2. Localize o arquivo Excel
3. Clique com botão direito → **"Compartilhar"**
4. Clique em **"Copiar link"**
5. Cole no app mobile

## ❓ Solução de Problemas

| Problema | Solução |
|----------|---------|
| Dados não carregam | Verifique se o link do OneDrive está correto |
| Erro de permissão | O arquivo precisa estar compartilhado como "Qualquer pessoa com o link" |
| App não instala | Certifique-se de estar usando HTTPS |

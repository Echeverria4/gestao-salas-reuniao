# SalasPro — Gestão de Salas de Reunião

Sistema web corporativo para **gerenciar salas de reunião por andar** e **agendar horários**.
Feito sem build step (HTML + React via CDN), pronto para deploy no **Vercel** com banco no **Supabase**.

## ✨ Funcionalidades

- **Início (Dashboard)** — visão geral: total de andares, salas ativas, salas livres agora e reuniões do dia, próximas reuniões e status em tempo real de cada sala.
- **Disponibilidade** — filtra salas por andar e data, mostra reservas do dia e o que está livre; agenda com 1 clique.
- **Agendar** — formulário completo (sala, data, horário, responsável, participantes, etc.) com **bloqueio automático de conflitos de horário**.
- **Agendamentos** — lista de reservas (próximos / hoje / passados / todos) com opção de cancelar.
- **Administração** — (protegida por senha) cadastro de **andares** e **salas** ilimitados, com capacidade, equipamentos, cor e localização.

> Funciona em **modo demonstração** (dados em memória) enquanto o Supabase não estiver configurado — ótimo para testar antes do deploy.

## 📁 Arquivos

| Arquivo | Função |
|---|---|
| `index.html` | Página única que carrega tudo |
| `config.js` | Suas chaves do Supabase + nome da empresa + senha admin |
| `data.js` | Camada de dados (Supabase ou modo demo) |
| `app.jsx` | Toda a interface React (telas e componentes) |
| `styles.css` | Design system corporativo |
| `schema.sql` | Script para criar as tabelas no Supabase |
| `vercel.json` | Configuração de deploy no Vercel |

## 🗄️ Banco de dados (Supabase)

Tabelas criadas pelo `schema.sql`:

- **`floors`** — andares (nome, número/ordem, descrição).
- **`rooms`** — salas (andar, nome, capacidade, localização, equipamentos, cor, ativa).
- **`bookings`** — agendamentos (sala, assunto, responsável, e-mail, departamento, participantes, início, fim, status).

A tabela `bookings` tem uma *exclusion constraint* (`bookings_no_overlap`) que **impede duas reservas confirmadas sobrepostas na mesma sala** — diretamente no banco.

## 🚀 Como colocar no ar

### 1. Criar o banco no Supabase
1. Crie um projeto em [supabase.com](https://supabase.com).
2. Vá em **SQL Editor → New query**, cole todo o conteúdo de `schema.sql` e clique em **RUN**.
3. Em **Project Settings → API**, copie a **Project URL** e a **anon public key**.

### 2. Configurar o app
Edite `config.js`:
```js
window.APP_CONFIG = {
  SUPABASE_URL: "https://xxxx.supabase.co",
  SUPABASE_ANON_KEY: "eyJhbGci...",
  COMPANY_NAME: "Nome da sua empresa",
  ADMIN_PASSWORD: "uma-senha-forte",
};
```

### 3. Deploy no Vercel
**Opção A — via site (mais fácil):**
1. Suba esta pasta para um repositório no GitHub.
2. Em [vercel.com](https://vercel.com) → **Add New → Project** → importe o repositório.
3. Framework Preset: **Other** (não precisa build). Clique em **Deploy**.

**Opção B — via CLI:**
```bash
npm i -g vercel
cd Gestao_Salas_Reuniao
vercel        # segue o assistente
vercel --prod # publica em produção
```

Pronto! O site abre direto na tela inicial.

## 🔒 Segurança (importante para produção)

Para começar rápido, o `schema.sql` libera leitura/escrita pública via a chave `anon`.
**Antes de usar pra valer**, recomenda-se:
- Ativar **Supabase Auth** e trocar as políticas RLS para usar `auth.uid()`/papéis.
- Restringir a aba de Administração ao Auth (a senha em `config.js` é apenas uma trava de front-end).

## 🧪 Testar localmente
Como tudo é estático, basta servir a pasta:
```bash
npx serve .
# ou
python -m http.server 8080
```
Abra `http://localhost:8080`.

---
Senha padrão de admin no modo demo: **`admin123`** (troque em `config.js`).

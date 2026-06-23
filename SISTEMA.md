# Gestão de Salas de Reunião — Pluma Agroavícola

Sistema corporativo web para **reserva e gerenciamento de salas de reunião por andar**, desenvolvido sob medida para a Pluma Agroavícola.

---

## O que é

Um aplicativo web acessado pelo navegador (sem instalação) que centraliza o controle de todas as salas de reunião da empresa. Qualquer colaborador com o link pode consultar disponibilidade e fazer agendamentos. O administrador gerencia andares, salas e reservas pelo mesmo sistema.

---

## Para que serve

| Problema que resolve | Como resolve |
|---|---|
| Conflito de reservas ("duas equipes marcam a mesma sala") | Bloqueio automático no banco de dados — impossível reservar uma sala já ocupada no mesmo horário |
| Falta de visibilidade ("não sei qual sala está livre agora") | Dashboard em tempo real com status de todas as salas |
| Comunicação manual ("reserva por WhatsApp/e-mail") | Formulário de agendamento digital com confirmação imediata |
| Controle descentralizado ("cada andar gerencia sozinho") | Painel único com todos os andares e salas da empresa |

---

## Funcionalidades

### Usuário comum

- **Ver disponibilidade** — Seleciona uma data e vê quais salas estão livres ou ocupadas, hora por hora, em todos os andares
- **Agendar sala** — Preenche título da reunião, nome, departamento, horário de início/fim e número de participantes
- **Cancelar reserva** — Cancela agendamentos futuros que foram criados por ele
- **Consultar histórico** — Lista todos os agendamentos do sistema (confirmados, cancelados)

### Administrador

Acessa o modo admin com senha (`Pluma123@` por padrão, alterável no `config.js`).

- **Gerenciar andares** — Adiciona, edita e remove andares (ex.: "Térreo", "1º Andar", "2º Andar")
- **Gerenciar salas** — Para cada andar, cadastra salas com:
  - Nome e localização
  - Capacidade de pessoas
  - Equipamentos disponíveis (projetor, TV, videoconferência, quadro branco, etc.)
  - Cor de identificação visual
- **Cancelar qualquer reserva** — Pode cancelar agendamentos de qualquer colaborador
- **Criar andares e salas ilimitados** — Sem limite de cadastros

---

## Telas do sistema

| Tela | Rota | O que mostra |
|---|---|---|
| **Dashboard** | `/` | Resumo do dia: total de salas, salas ocupadas agora, próximas reuniões, agenda de hoje |
| **Disponibilidade** | `/disponibilidade` | Grade visual de todas as salas por horário em uma data escolhida |
| **Agendar** | `/agendar` | Formulário completo de agendamento de sala |
| **Reservas** | `/reservas` | Lista de todos os agendamentos com filtros e opção de cancelamento |
| **Admin** | `/admin` | Painel protegido por senha para cadastro de andares, salas e gestão geral |

---

## Tecnologia

| Componente | Tecnologia |
|---|---|
| Front-end | React 18 (via CDN, sem build step) |
| Linguagem | JSX transpilado no browser pelo Babel Standalone |
| Estilização | CSS puro com design system Pluma (verde `#1f574b` + dourado `#b6862f`) |
| Fontes | IBM Plex Sans (corpo) + Encode Sans Condensed (labels/tabelas) |
| Banco de dados | Supabase (PostgreSQL na nuvem) — projeto compartilhado com o Arena |
| Hospedagem | Vercel (deploy automático a cada `git push`) |
| Repositório | GitHub — `Echeverria4/Salas-de-Reuniao` |

### Tabelas no banco de dados

```
floors    → andares cadastrados (id, nome, número, descrição)
rooms     → salas por andar (id, andar, nome, capacidade, localização, equipamentos, cor)
bookings  → agendamentos (id, sala, título, organizador, departamento, horário início/fim, status)
```

A tabela `bookings` tem uma **constraint de exclusão** no PostgreSQL que impede fisicamente o cadastro de dois agendamentos simultâneos para a mesma sala — mesmo que dois usuários tentem ao mesmo tempo.

---

## Modo demonstração

Se o sistema for aberto sem conexão com o Supabase (credenciais ausentes ou inválidas), ele entra em **modo demo**: carrega dados fictícios em memória com 3 andares, 5 salas e agendamentos de hoje, permitindo explorar todas as telas sem banco de dados.

---

## Como acessar

| Ambiente | URL |
|---|---|
| Produção (Vercel) | Deploy automático via GitHub push |
| Local | Abrir `index.html` em qualquer servidor HTTP local |

Para rodar localmente com Python:
```bash
cd c:\PYTHON\Gestao_Salas_Reuniao
python -m http.server 8080
# Acessar: http://localhost:8080
```

---

## Como publicar alterações

```bash
cd c:\PYTHON\Gestao_Salas_Reuniao
git add .
git commit -m "descrição da alteração"
git push   # Vercel detecta e faz deploy automático em ~15 segundos
```

---

## Arquivos do projeto

```
Gestao_Salas_Reuniao/
├── index.html      # Ponto de entrada — carrega fontes, libs e scripts
├── config.js       # URL Supabase, chave anon, nome da empresa, senha admin
├── data.js         # Camada de dados: lê/grava no Supabase ou no modo demo
├── app.jsx         # App React: todas as telas, componentes e lógica de UI
├── styles.css      # Design system completo com variáveis de cor e tipografia
├── schema.sql      # Script SQL para criar as tabelas no Supabase
├── vercel.json     # Configuração de rotas para o Vercel
└── SISTEMA.md      # Este arquivo
```

---

## Configuração inicial (novo ambiente)

1. Criar projeto no [Supabase](https://supabase.com) (ou reutilizar o existente)
2. Executar `schema.sql` no SQL Editor do Supabase
3. Preencher `config.js` com a URL e a chave anon do projeto
4. Fazer `git push` para o repositório GitHub conectado ao Vercel
5. Acessar a URL gerada pelo Vercel

---

*Desenvolvido para Pluma Agroavícola — Planejamento Integrado*

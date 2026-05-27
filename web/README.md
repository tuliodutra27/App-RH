# 📊 App-RH — Analisador de Planilhas do RH

> Sistema web para comparação automática de planilhas semanais de colaboradores, com autenticação via Active Directory e relatórios em Excel.

![Python](https://img.shields.io/badge/Python-3.12-blue?logo=python)
![Flask](https://img.shields.io/badge/Flask-3.1-lightgrey?logo=flask)
![Docker](https://img.shields.io/badge/Docker-Compose-2496ED?logo=docker)
![Bootstrap](https://img.shields.io/badge/Bootstrap-5-purple?logo=bootstrap)

---

## 🎯 O que faz

A cada semana o RH envia uma planilha `.xlsx` com todos os colaboradores ativos. Este sistema:

- **Compara** a nova planilha com a anterior automaticamente
- **Detecta** admissões, desligamentos e alterações de cargo/departamento/gestor
- **Exibe** as diferenças em uma interface web com abas e filtro por nome
- **Gera** relatório em Excel formatado com cores para download
- **Controla** acesso via Active Directory (somente membros do grupo TI)

---

## ✨ Funcionalidades

| Funcionalidade | Descrição |
|---|---|
| 🔐 Login AD | Autenticação com usuário e senha de rede (LDAP/LDAPS) |
| 📤 Upload de planilha | Envio do `.xlsx` via browser |
| 🆕 Admissões | Colaboradores que aparecem na nova planilha e não estavam na anterior |
| ❌ Desligamentos | Colaboradores que saíram entre as duas planilhas |
| 🔄 Alterações | Mudanças de CARGO, DEPARTAMENTO ou GESTOR |
| 📜 Histórico | Dashboard com todas as comparações anteriores |
| 📥 Export Excel | Download do relatório com células coloridas por tipo |
| 🧪 Modo Dev | Login sem AD para testes (desabilitado em produção) |

---

## 🧱 Stack

- **Backend:** Python 3.12 + Flask 3.1 + Gunicorn
- **Planilhas:** pandas + openpyxl
- **Autenticação:** ldap3 (Active Directory)
- **Frontend:** Bootstrap 5 + HTML/CSS
- **Infraestrutura:** Docker + Docker Compose

---

## 📋 Pré-requisitos

### No servidor de destino
- Ubuntu 22.04 / 24.04 (ou qualquer Linux com Docker)
- Docker Engine 24+
- Docker Compose v2+
- Acesso de rede ao servidor AD na porta **389** (LDAP) ou **636** (LDAPS)
- Porta **5000** liberada no firewall (ou outra porta que você configurar)

### Instalar Docker no Ubuntu (se ainda não tiver)
```bash
curl -fsSL https://get.docker.com | sh
sudo usermod -aG docker $USER
newgrp docker
docker --version   # confirma instalação
```

---

## 🚀 Deploy — Passo a Passo

### 1. Clonar o repositório no servidor

```bash
cd /opt
sudo git clone https://github.com/tuliodutra27/App-RH.git
sudo chown -R $USER:$USER /opt/App-RH
cd /opt/App-RH/web
```

### 2. Criar e configurar o `.env`

```bash
cp .env.example .env
nano .env
```

Preencha com os dados reais do seu ambiente:

```env
# Gere com: python3 -c "import secrets; print(secrets.token_hex(32))"
SECRET_KEY=cole-aqui-a-chave-gerada

# Porta que será exposta (padrão: 5000)
APP_PORT=5000

# Active Directory
AD_SERVER=ldap://192.168.1.10       # IP do seu servidor AD
AD_PORT=389
AD_USE_SSL=false
AD_DOMAIN=aliseo.local
AD_BASE_DN=DC=aliseo,DC=local
AD_TI_GROUP_CN=TI                   # grupo AD com permissão de acesso

# Modo dev — MANTER false em produção!
DEV_MODE=false
DEV_USERS={}
```

> 💡 **Gerar SECRET_KEY segura:**
> ```bash
> python3 -c "import secrets; print(secrets.token_hex(32))"
> ```

### 3. Subir o container

```bash
docker compose up -d --build
```

A imagem será construída (~2 min na primeira vez) e o container iniciará automaticamente.

### 4. Verificar que está rodando

```bash
docker compose ps          # deve mostrar status "healthy"
docker compose logs -f     # acompanhar logs em tempo real (Ctrl+C para sair)
```

### 5. Acessar no browser

```
http://IP-DO-SERVIDOR:5000
```

---

## ⚙️ Variáveis de ambiente

| Variável | Obrigatória | Padrão | Descrição |
|---|---|---|---|
| `SECRET_KEY` | ✅ | — | Chave secreta Flask (gere aleatoriamente) |
| `APP_PORT` | ❌ | `5000` | Porta exposta pelo Docker |
| `AD_SERVER` | ✅ | — | Endereço LDAP do AD (`ldap://IP`) |
| `AD_PORT` | ❌ | `389` | Porta LDAP (636 para LDAPS) |
| `AD_USE_SSL` | ❌ | `false` | Usar LDAPS (true/false) |
| `AD_DOMAIN` | ✅ | — | Domínio AD (ex: `aliseo.local`) |
| `AD_BASE_DN` | ✅ | — | Base DN (ex: `DC=aliseo,DC=local`) |
| `AD_TI_GROUP_CN` | ❌ | `TI` | Grupo AD com acesso ao sistema |
| `DEV_MODE` | ❌ | `false` | Login sem AD para testes |
| `DEV_USERS` | ❌ | `{}` | Usuários de teste (JSON) |

---

## 📁 Estrutura de dados persistida

Os dados ficam em `web/data/` (volume Docker montado em `/data` no container):

```
data/
├── historico.json            ← índice de todas as comparações
├── snapshots/                ← cópias das planilhas enviadas (base de comparação)
│   └── snapshot_20250527_143000_efetivo-26.05.xlsx
├── relatorios/               ← resultados em JSON por comparação
│   └── resultado_20250527_143000.json
└── uploads/                  ← arquivos enviados pelos usuários
    └── 20250527_143000_efetivo-26.05.xlsx
```

> ⚠️ **Faça backup periódico da pasta `data/`**, especialmente `snapshots/` e `historico.json`.

---

## 📌 Regras de negócio

### Chave de identificação dos colaboradores

| Tipo | Campo-chave | Exemplo de chave interna |
|---|---|---|
| **CLT** | MATRÍCULA numérica | `CLT::610` |
| **PJ** | NOME (sem matrícula numérica) | `PJ::FULANO DA SILVA` |

### Campos monitorados para "alteração"
- **CARGO**
- **DEPARTAMENTO**
- **GESTOR**

Qualquer mudança nesses campos entre duas planilhas gera um registro na seção "Alterações".

### Fluxo de comparação
1. **Primeira planilha enviada** → criada como baseline (sem diferenças exibidas)
2. **Planilhas seguintes** → comparadas contra o snapshot anterior

---

## 🔄 Como atualizar o sistema

```bash
cd /opt/App-RH
git pull origin main
cd web
docker compose up -d --build
```

---

## 🛑 Comandos úteis

```bash
# Reiniciar container (ex: após editar .env)
docker compose restart

# Parar tudo
docker compose down

# Ver logs em tempo real
docker compose logs -f

# Ver logs das últimas 100 linhas
docker compose logs --tail=100

# Reconstruir imagem (após update do código)
docker compose up -d --build

# Verificar saúde do container
docker inspect --format='{{.State.Health.Status}}' analisador-rh
```

---

## 🧪 Modo desenvolvimento (sem AD)

Para testar sem precisar de Active Directory:

```env
# no .env
DEV_MODE=true
DEV_USERS={"admin":"admin123","teste":"teste123"}
```

```bash
docker compose restart
```

> ⚠️ **NUNCA deixe `DEV_MODE=true` em produção!**

---

## 🔒 Boas práticas de segurança

- Use uma `SECRET_KEY` longa e aleatória (mínimo 32 bytes)
- Nunca versione o arquivo `.env` (já está no `.gitignore`)
- Considere usar LDAPS (`AD_USE_SSL=true`, porta 636) para criptografar o tráfego AD
- Coloque um **Nginx como proxy reverso** na frente se for expor externamente (adiciona HTTPS)
- Faça backup regular da pasta `data/`

---

## ❓ Troubleshooting

### Container não sobe / fica reiniciando
```bash
docker compose logs   # ver o erro
```

### Erro de autenticação AD
- Verifique se o servidor AD está acessível: `ping 192.168.1.10`
- Teste a porta LDAP: `nc -zv 192.168.1.10 389`
- Confirme que `AD_DOMAIN` e `AD_BASE_DN` estão corretos
- O usuário precisa ser membro do grupo definido em `AD_TI_GROUP_CN`

### Planilha não processa / erro de colunas
- Abra a planilha e verifique se não há linhas em branco **no topo** (antes do cabeçalho)
- Colunas esperadas (após normalização para maiúsculas): `MATRÍCULA, NOME, CARGO, DEPARTAMENTO, GESTOR, ADMISSÃO`

### Porta 5000 já em uso
- Mude `APP_PORT=5001` no `.env` e faça `docker compose up -d`

---

## 👨‍💻 Desenvolvido por

**Setor de TI — AliseoSA**  
Tulio Pereira · [`tuliodutra27@gmail.com`](mailto:tuliodutra27@gmail.com)

---

## 📄 Licença

Uso interno — AliseoSA. Todos os direitos reservados.

# Analisador RH — AliseoSA (Web)

Sistema web para comparação de planilhas do RH, com login via Active Directory.

## Requisitos

- Docker Engine 24+
- Docker Compose v2+
- Acesso de rede ao servidor AD (porta 389 ou 636)

## Configuração inicial

### 1. Clone / copie os arquivos para o servidor

```bash
# Na máquina onde o Docker está instalado
cd /opt   # ou onde quiser instalar
```

### 2. Configure o .env

```bash
cp .env.example .env
nano .env          # edite com os dados do seu AD
```

Campos obrigatórios:
| Variável | Descrição | Exemplo |
|---|---|---|
| `SECRET_KEY` | Chave secreta Flask (gere uma aleatória) | `python -c "import secrets; print(secrets.token_hex(32))"` |
| `AD_SERVER` | Endereço do servidor AD | `ldap://192.168.1.10` |
| `AD_DOMAIN` | Domínio AD | `aliseo.local` |
| `AD_BASE_DN` | Base DN para pesquisa | `DC=aliseo,DC=local` |
| `AD_TI_GROUP_CN` | CN do grupo com acesso | `TI` |

### 3. Suba o container

```bash
docker compose up -d
```

Acesse: **http://IP-DO-SERVIDOR:5000**

---

## Uso

1. **Login**: Entre com seu usuário de rede (mesmo do Windows) — sem o domínio
2. **Dashboard**: Veja o histórico de comparações e os números da última planilha
3. **Enviar planilha**: Clique em "Enviar nova planilha" e selecione o arquivo `.xlsx`
   - **Primeira vez**: cria a baseline (referência para futuras comparações)
   - **Próximas vezes**: compara com a planilha anterior e mostra admissões, desligamentos e alterações
4. **Resultado**: Veja as diferenças organizadas em abas com filtro por nome
5. **Download**: Baixe o relatório em Excel com formatação colorida

---

## Estrutura dos dados

```
data/
├── historico.json          # índice de todas as comparações
├── snapshots/              # cópias das planilhas enviadas (base de comparação)
│   └── snapshot_YYYYMMDD_HHMMSS_DD-MM.xlsx
├── relatorios/             # resultados em JSON
│   └── resultado_YYYYMMDD_HHMMSS.json
└── uploads/                # arquivos enviados pelo usuário
    └── YYYYMMDD_HHMMSS_efetivo-DD.MM.xlsx
```

> **Backup**: faça backup periódico da pasta `data/`, especialmente `snapshots/` e `historico.json`.

---

## Comandos úteis

```bash
# Ver logs em tempo real
docker compose logs -f

# Reiniciar após alteração no .env
docker compose restart

# Reconstruir imagem após atualização do código
docker compose up -d --build

# Parar
docker compose down
```

---

## Teste sem AD (modo desenvolvimento)

Útil para validar a aplicação antes de conectar ao AD:

```env
# no .env
DEV_MODE=true
DEV_USERS={"admin":"senha123"}
```

> ⚠️ **NUNCA deixe DEV_MODE=true em produção!**

---

## Coluna monitoradas para "alteração"

- **CARGO**
- **DEPARTAMENTO**
- **GESTOR**

Mudança em qualquer um desses campos entre duas planilhas gera uma linha na seção "Alterações".

## Identificação de colaboradores

- **CLT**: chave = `MATRÍCULA` (ex: `CLT::610`)
- **PJ**: chave = `NOME` (ex: `PJ::FULANO DA SILVA`) — não têm matrícula numérica

---

## Suporte

Setor de TI — AliseoSA

# Deploy - geradorderelatorios.com.br

## Configuracao recomendada

Para colocar no ar com seguranca, use:

- Aplicacao Python/Flask com Gunicorn.
- Banco PostgreSQL em producao.
- Pasta persistente ou storage externo para DOCX, ZIP e imagens.
- HTTPS habilitado.
- Variavel `RL_METAIS_SECRET` com uma chave forte.

## Comando de start

```bash
gunicorn web_app:app
```

## Dependencias de producao

```bash
pip install -r requirements-production.txt
```

## Variaveis de ambiente

```bash
RL_METAIS_SECRET=troque-por-uma-chave-grande-e-secreta
DATABASE_URL=postgresql://usuario:senha@host:5432/banco
RL_METAIS_OUTPUT_DIR=/data/output
RL_METAIS_UPLOAD_DIR=/data/uploads
```

Se usar SQLite em servidor com disco persistente:

```bash
RL_METAIS_DATA_DIR=/data
DATABASE_URL=
```

## DNS do dominio

No painel onde voce comprou `geradorderelatorios.com.br`, configure:

- `www` como `CNAME` apontando para o host da aplicacao.
- raiz `@` como `A` ou `ALIAS/ANAME`, conforme o provedor de hospedagem informar.

Depois configure o dominio customizado na hospedagem:

```text
www.geradorderelatorios.com.br
geradorderelatorios.com.br
```

## Primeiro acesso

Apos publicar, acesse:

```text
https://www.geradorderelatorios.com.br/setup
```

Crie a empresa e o usuario administrador.

## Observacao importante

Hospedagens com filesystem temporario apagam o SQLite e os arquivos gerados a cada redeploy/restart. Para uso real, prefira PostgreSQL e storage persistente.

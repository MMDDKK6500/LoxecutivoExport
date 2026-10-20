# Loxecutivo Export

Microserviço responsável por buscar um evento no Loxecutivo Core e gerar o relatório Excel aprovado para o projeto.

## Configuração

```bash
cp .env.example .env
npm install
npm start
```

Variáveis:

- `PORT`: porta do serviço (padrão `8083`)
- `CORE_URL`: URL interna do Loxecutivo Core (padrão `http://127.0.0.1:8081`)
- `REQUEST_TIMEOUT_MS`: tempo limite das chamadas ao Core

## API

### Saúde

```http
GET /health
```

### Exportar um evento

```http
GET /eventos/{id}/excel
Authorization: Bearer <jwt>
```

O serviço encaminha o JWT para o Core. O Core continua responsável por validar a sessão através do serviço de autenticação.

O arquivo contém as abas `Resumo`, `Viagens`, `Passageiros`, `Motoristas`, `Veículos` e `Endereços`. As viagens são agrupadas por motorista e ordenadas por data e hora. A aba de passageiros contém uma coleção (`P-001`, `P-002`, etc.) e sua quantidade para cada viagem.

## Nginx

Adicione este bloco ao servidor HTTPS, junto das outras rotas de API:

```nginx
location /export/ {
    proxy_pass http://localhost:8083/;
    proxy_set_header Host loxecutivo.mmddkk.dev.br;
    proxy_set_header X-Real-IP $remote_addr;
    proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    proxy_set_header X-Forwarded-Proto $scheme;
}
```

Depois disso, o frontend pode chamar `/export/eventos/{id}/excel` e baixar a resposta como um arquivo `.xlsx`.

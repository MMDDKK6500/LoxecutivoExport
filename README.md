# Loxecutivo Export

Microserviço responsável por buscar dados no Loxecutivo Core e gerar relatórios Excel e PDF por evento ou da operação inteira.

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
- `CORS_ORIGINS`: origens permitidas pelo CORS, separadas por vírgulas (padrão `https://loxecutivo.mmddkk.dev.br,http://localhost:5173`)

Em produção, o frontend e as APIs ficam na mesma origem HTTPS (`https://loxecutivo.mmddkk.dev.br`) por meio do NGINX. O CORS do serviço aceita somente essa origem e a origem local do Vite para desenvolvimento; configure `CORS_ORIGINS` caso use outra origem.

## API

### Saúde

```http
GET /health
```

### Exportar um evento para Excel

```http
GET /eventos/{id}/excel
Authorization: Bearer <jwt>
```

### Exportar um evento para PDF

```http
GET /eventos/{id}/pdf
Authorization: Bearer <jwt>
```

### Consultar o relatório geral da empresa

```http
GET /relatorio-geral?de=2026-01-01&ate=2026-12-31
Authorization: Bearer <jwt>
```

As datas são opcionais e inclusivas, no formato `AAAA-MM-DD`. Sem filtro, o relatório considera todo o histórico disponível.

### Exportar o relatório geral para Excel ou PDF

```http
GET /relatorio-geral/excel?de=2026-01-01&ate=2026-12-31
GET /relatorio-geral/pdf?de=2026-01-01&ate=2026-12-31
Authorization: Bearer <jwt>
```

O serviço encaminha o JWT para o Core. O Core continua responsável por validar a sessão através do serviço de autenticação.

O arquivo contém as abas `Resumo`, `Viagens`, `Passageiros`, `Motoristas`, `Veículos` e `Endereços`. As viagens são agrupadas por motorista e ordenadas por data e hora. A aba de passageiros contém uma coleção (`P-001`, `P-002`, etc.) e sua quantidade para cada viagem.

O PDF por evento usa o mesmo conteúdo e agrupamento do Excel em um relatório A4 com resumo, viagens, coleções de passageiros, motoristas, veículos e endereços.

O relatório geral apresenta viagens por status, passageiros transportados em viagens realizadas, média de passageiros por viagem realizada, viagens por motorista, veículo e evento, distribuição por dia e hora, e rotas mais utilizadas. As abas incluem também a lista detalhada de viagens e os registros cadastrados sem uso no período para ajudar a identificar recursos ociosos. O cálculo não estima receita, custos, quilometragem, duração ou ocupação, pois esses dados ainda não são registrados pelo sistema.

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

Depois disso, o frontend pode chamar `/export/eventos/{id}/excel` ou `/export/relatorio-geral/excel` para baixar um `.xlsx`, e as rotas `/pdf` correspondentes para baixar um `.pdf`.

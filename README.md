# BTG Pactual — Debêntures Distressed Monitor

## Como usar

1. Coloque os arquivos Excel na pasta:
   - `anbima v0.xlsx` — data base (ANBIMA MSD anterior)
   - `anbima data.xlsx` — data atual (ANBIMA MSD mais recente)
   - `119452.xlsx` — BTG Weekly (classificação setorial)

2. Execute `start.bat` (duplo clique) ou:
   ```
   py app.py
   ```

3. Acesse: **http://localhost:5000**

## Funcionalidades

- KPI cards: total de papéis, distressed, em stress, sem cotação
- Filtros: indexador, setor, delta spread, % PU Par, busca livre
- Tabs rápidas: Todos / Top Aberturas / Distressed / Eventos
- Tabela ordenável com delta pills, barras de PU Par, badge de risco
- **Clique em qualquer papel** para ver gráfico de variação, PD implícita e detalhes completos
- Gráficos de overview: top 10 aberturas, distribuição PU Par, delta por indexador

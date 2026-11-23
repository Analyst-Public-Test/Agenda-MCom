# Agenda MCom

Gera `dados/agenda-mcom.json` com a agenda do Secretário e dos 3 diretores da SERAD, **de 2026 em diante** (2023–2025 ficam fixos no HTML).

| Fonte | O que traz | Como é lida |
|---|---|---|
| e-Agendas (CGU) | compromissos | página pública de cada autoridade (sem captcha) |
| Portal da Transparência | viagens (PCDP, processo, motivo) | zip anual `.../download-de-dados/viagens/ANO` → `ANO_Viagem.csv` |

- **Automático:** todo dia às 8h e 14h (Brasília), ou manual em *Actions → Agenda MCom → Run workflow*. Em 2027 não precisa mexer em nada: o ano novo entra sozinho.
- **Se uma fonte falhar:** o que já estava no JSON é mantido.
- **Plano B:** coloque `.xlsx`/`.csv` do e-Agendas em `dados/entrada/` e o zip/csv de viagens em `dados/entrada/viagens/` (o nome do arquivo de viagens deve conter o ano).
- **Autoridades:** `config/autoridades.json` (trocar de diretor = trocar uma linha).
- **Correção manual de classificação ou lugar:** `config/ajustes.json` (nunca é sobrescrito).
- **Revisar:** `pendencias_de_lugar` no JSON lista compromissos presenciais fora de Brasília que não foram identificados.

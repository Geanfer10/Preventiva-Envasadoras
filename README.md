# Painel de Preventiva · Envasadoras (Goiás Verde · PCM)

Painel publicado no GitHub Pages que cruza:
- **Planilha de tagueamento** (`uploads/`): abas Base, Preventiva e Execuções
- **BASE_PCM** do repositório [MTBF-MTTR](https://github.com/geanfer10/MTBF-MTTR): falhas corretivas
- **Programação de Produção Diária** do PCP (`programacao/`): PDF recebido por e-mail

## Como os dados chegam (sem trabalho manual)
1. O PCP envia o e-mail "Programação de Produção dd.mm.aaaa" (Clênio, Claudionor ou Cleiton em CC).
2. O fluxo **Programação → Painel Preventiva** (Power Automate) salva o PDF em
   `OneDrive - GOIAS VERDE ALIMENTOS LTDA\Painel Preventiva\programacao`.
3. O **Robô PCM** (tarefa agendada no PC do PCM, a cada 15 min) copia os PDFs novos e a planilha
   `X:\PCM\21 - TAGMENTO EQUIPAMENTOS\PLANO PREVENTIVO\BASE_SISTEMA_DE_TAGMENTO_v2.xlsx` (quando mudar) e faz o push.
   O robô fica fora deste repositório (pasta `RoboPCM`) e também atende o MTBF-MTTR.
4. O GitHub Actions (`.github/workflows/atualizar-painel.yml`) gera o `index.html`.
   Ele também roda de hora em hora (06h–22h) para pegar a BASE_PCM atualizada do MTBF-MTTR.

## Arquivos
| Caminho | O que é |
|---|---|
| `scripts/build.py` | Lê planilha + BASE_PCM + PDFs e gera o `index.html` |
| `scripts/parse_prog.py` | Leitor do PDF da programação (horários, CIP, setup, janelas) |
| `scripts/template.html` | Layout do painel |
| `data/programacao/*.json` | Histórico da programação lida, dia a dia |

## Regras importantes
- **Plano** = uma linha na aba Preventiva com TAG, RG (igual à aba Base), ATIVIDADE e FREQ.
- **Execução** = uma linha na aba Execuções, escolhendo o PLANO na lista. A última execução sai daí.
- Plano sem nenhuma execução registrada aparece como **vencido** ("sem registro").
- Linhas da programação são mapeadas para as máquinas em `parse_prog.py` (`maquinas_da_linha`).
  Máquina que não aparece no PDF = sem programação no dia.

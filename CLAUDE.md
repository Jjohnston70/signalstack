# CLAUDE.md — ML-SIGNAL-STACK

Agent-facing context for this repo. Applies to any LLM coding agent touching this code.

## What this is

SARIMA-based forecasting pipeline for small-operator KPIs (sales, ops tempo,
cash flow, pipeline, team activity). Adjusted-AIC model selection,
convergence filtering, MAPE guardrails, Word-report packaging.

All data in this repository is synthetic. There is no client PII or financial
information in any template, sample, or output artifact.

## Hard rules

1. **No PII ever.** This repo is public. Do not commit real company data,
   real customer names, real vendor names, real financial figures, or
   anything that resembles production data.
2. **Keep outputs reproducible.** `run_pipeline.py` must be runnable end-to-end
   against the bundled sample data without external services.
3. **No hardcoded local paths.** Every path is relative to the repo root or
   configured in `config.py`.
4. **Preserve the interface.** `config.py`, `run_pipeline.py`, and the five
   module outputs (`sales`, `ops_pulse`, `cash_flow_compass`, `pipeline_pulse`,
   `team_tempo`) form the public contract. Do not rename without updating
   the user manual and generated reports.
5. **Small surgical diffs.** Match existing style. Do not refactor unrelated
   code. Flag risks before acting.

## How it works (short)

```
run_pipeline.py
  └── src/data_loader.py       # read xlsx templates from data/raw
  └── src/preprocessor.py      # weekly aggregation, stationarity checks
  └── src/model.py             # SARIMA grid search with adjusted-AIC
  └── src/evaluator.py         # MAPE / RMSE / convergence guardrails
  └── src/visualizer.py        # matplotlib + docx report
  └── src/accuracy_log.py      # append to accuracy_log.csv
export_to_csv.py               # module-scoped CSV exports
generate_report.py             # Word report per module
package_output.py              # HTML + zipped deliverable
```

## When editing

- State assumptions before implementing if the task is ambiguous.
- Run `python run_pipeline.py` end-to-end after any src/ change.
- If you add a dependency, add it to `requirements.txt`.
- If you change a module name, update SIGNALSTACK_USER_MANUAL.md.

"""Global Market Research & Company Database pipeline.

Given a MARKET string, discovers real companies relevant to that market and
emits verified batches of up to 20 rows in the fixed 21-column schema defined
in ``schema.py``. Never fabricates a field: a company is dropped from the
batch rather than exported with a guessed/placeholder value.

Entry point: ``batch_runner.run_batch()``. See ``run_market_research.py`` at
the repo root for the CLI.
"""

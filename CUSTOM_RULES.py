"""Optional project-specific hook.

Keep ordinary mappings, formulas, KPIs and charts in PROJECT.json.  The AI
client may replace this function only when a row rule cannot be expressed in
that configuration.
"""


def transform_rows(rows, config):
    return rows

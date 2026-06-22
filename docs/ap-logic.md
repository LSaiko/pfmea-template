# Action Priority (AP) Logic

The generator computes Action Priority with `get_action_priority(s, o, d)` in
[`foremode.py`](../foremode.py). It is **transparent and auditable** — when an FDA or
notified-body auditor asks *"why is this Medium?"* the answer is one readable
branch, not a black-box model. This is a deliberate differentiator vs. AI FMEA
tools whose ratings are non-reproducible.

## The rules (detection-weighted for patient safety)

```
S 9-10 + D 7-10            -> H   (undetected critical failure, any O)
S 9-10 + D 4-6 + O >= 4    -> H
S 9-10 + D 4-6 + O <  4    -> M
S 9-10 + D 1-3             -> M   (strong detection mitigates)
S 7-8  + D 7-10 + O >= 4   -> H
S 7-8  + D 7-10 + O <  4   -> M
S 7-8  + D 4-6  + O >= 4   -> M
S 7-8  + D 4-6  + O <  4   -> L
S 7-8  + D 1-3             -> L
S <=6  + D >=7 + O >= 6    -> M
S <=6  (otherwise)         -> L
```

S, O, D must be integers 1–10; out-of-range values raise `ValueError` (a risk
tool must not silently misclassify).

## Important scope note

This is a **documented simplification**, not the full AIAG-VDA 2019 handbook AP
table (which is a ~1000-cell S×O×D lookup). It captures the handbook's core
principle — detection gaps on high-severity failures dominate — and is suitable
for medical-device pFMEA triage and portfolio/illustrative use. For a regulated
submission that must cite the handbook table verbatim, replace this function with
the licensed AIAG-VDA table; everything else in the pipeline is unchanged because
AP is the single source of truth feeding Sheet 5 *and* the Sheet 7 metrics.

## Why metrics are computed, not stored

Sheet 7 counts (initial H/M/L, after-action H/M/L) are derived from the same
`get_action_priority` call that fills Sheet 5. They cannot drift out of sync.
The previous per-scenario scripts hand-stored these counts and were already
internally inconsistent (e.g. Sheet 7 claimed "2 initial High" while Sheet 5
computed only 1). Computing eliminates that whole class of error.

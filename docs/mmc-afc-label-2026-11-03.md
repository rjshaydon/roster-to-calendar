# MMC AFC — 3 November 2026

User confirmed AFC means AM Float Consultant and should display Float AM for
SMS. The standard parser rules lacked this code. Added the SMS-specific default
with normal MMC AM hours (08:00–17:30); explicit roster hours take precedence.
Cached MMC SMS AFC events normalize to `MMC: Float AM` without changing their
hours, date, grade or raw text. Calendar filtering and Who assignment grouping
use the correction. Other sources/grades/codes are unchanged. No clinical data
rewrite or test import is needed. Existing explicit admin parser rules retain
normal precedence. Focused parser/cache/isolation tests, MMC Orientation and DDH
Clinical Support regressions, syntax and whitespace checks passed.

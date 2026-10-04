-- Compact term policy only: no roster facts, personal data or raw files change.
UPDATE facility_term_visibility
SET visible_from = date(term_start, 'start of month', '-1 month')
WHERE visible_from <> date(term_start, 'start of month', '-1 month');

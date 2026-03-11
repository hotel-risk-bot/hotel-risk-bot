# Proposal Builder App — Audit Findings & Improvements

> Last reviewed: 2026-03-11

## Critical Issues

### 1. ~~Workers Compensation key inconsistency (`workers_comp` vs `workers_compensation`)~~ ✅ RESOLVED
- **File**: `proposal_generator.py` and `web_app.py`
- **Status**: Fixed — `proposal_generator.py` line 3529 now checks both keys via `_wc_key` logic. `web_app.py` lines 1233-1240 normalizes `workers_compensation` → `workers_comp` before calling generator.

### 2. ~~Expiring premiums data not properly passed to generator~~ ✅ RESOLVED
- **File**: `templates/proposal_web.html` (syncVisualToData) and `web_app.py`
- **Status**: Fixed — `web_app.py` lines 1229-1231 maps `expiring_premiums_data` → `expiring_premiums`. `syncVisualToData()` populates both keys for backward compatibility.

### 3. ~~Missing `enviro_pack` coverage type in UI dropdown~~ ✅ RESOLVED
- **File**: `templates/proposal_web.html`
- **Status**: Fixed — `enviro_pack` is present in the coverage dropdown (line 690) and label mapping (line 1080).

### 4. ~~`workers_comp` duplicate key handling in premium summary~~ ✅ RESOLVED
- **File**: `proposal_generator.py` line 773+
- **Status**: Fixed — Added deduplication logic to `all_keys` list. Now checks which key has actual data and removes the unused variant to prevent duplicate rows.

## UI/UX Improvements

### 5. No expiring carrier/details input — OPEN (Enhancement)
- Users can enter expiring premiums but not expiring carrier names
- The generator supports `expiring_details` with carrier info for comparison mode
- **Fix**: Add expiring carrier input fields next to expiring premium inputs

### 6. ~~No deductible add button~~ ✅ RESOLVED
- **Status**: Fixed — `addCovDed()` function and "+ Add Deductible" button exist (line 1247, 1344).

### 7. ~~No sublimit/additional coverage add button~~ ✅ RESOLVED
- **Status**: Fixed — `addAddlCov()` function and "+ Add Sublimit" button exist (line 1284, 1375).

### 8. ~~No subjectivity add button~~ ✅ RESOLVED
- **Status**: Fixed — `addSubj()` function and "+ Add Subjectivity" button exist (line 1301, 1388).

### 9. Additional coverages field name mismatch — MITIGATED
- The `renderCoverageFields` reads `ac.description || ac.coverage || ac.name` with fallback logic
- The `updateAddlCov` writes to `description`
- The generator reads `ac.get("description")`
- **Status**: The fallback chain in the UI handles legacy data, and new entries use `description` consistently. Generator reads `description`. No action needed unless extraction produces `coverage` or `name` fields — in which case the UI will display them correctly and save as `description` on edit.

## Extraction Improvements

### 10. Truncation at 200K chars may lose data — OPEN (Low Priority)
- Large multi-property SOVs can exceed 200K chars
- **Fix**: Increase to 300K or implement smarter truncation (prioritize quote text over SOV)

### 11. Pass 2-4 send the ENTIRE combined_text each time — OPEN (Low Priority)
- Each focused pass sends the full document text, which is wasteful for gpt-4.1-mini
- Could be optimized to send only relevant sections
- **Note**: Works correctly, just costs more tokens

## DOCX Generation Improvements

### 12. ~~`additional_coverages` field name inconsistency in sublimits~~ ✅ RESOLVED
- **Status**: See Finding #9 — fallback chain in UI handles this. Generator consistently reads `description`.

### 13. ~~Missing `enviro_pack` in coverage section generation~~ ✅ RESOLVED
- **Status**: Fixed — `proposal_generator.py` lines 3657-3658 have `if "enviro_pack" in coverages: generate_coverage_section(...)`.

### 14. ~~Confirmation to Bind section missing effective date~~ ✅ RESOLVED
- **Status**: Fixed — `generate_confirmation_to_bind()` at lines 2936-2938 prominently displays the effective date in Electric Blue bold text.

## Additional Fix (2026-03-11)
- **`web_app.py`**: Expanded `coverage_display` dict in `_build_review_summary()` to include all supported coverage types (enviro_pack, terrorism, equipment_breakdown, liquor_liability, etc.) — previously only had 14 of 30+ supported types.

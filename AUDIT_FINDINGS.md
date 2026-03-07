# Proposal Builder App — Audit Findings & Improvements

## Critical Issues

### 1. Workers Compensation key inconsistency (`workers_comp` vs `workers_compensation`)
- **File**: `proposal_generator.py` and `web_app.py`
- The extraction prompt and UI use `workers_compensation`, but `generate_proposal()` checks for `workers_comp` (line 3449)
- The `coverage_names` dict in multiple sections has BOTH keys, but the coverage section generator only checks `workers_comp`
- **Fix**: Add `workers_compensation` check alongside `workers_comp` in `generate_proposal()`

### 2. Expiring premiums data not properly passed to generator
- **File**: `templates/proposal_web.html` (syncVisualToData)
- `syncVisualToData()` stores expiring premiums in `expiring_premiums_data` but the generator reads from `expiring_premiums`
- **Fix**: Ensure the generator receives the correct key, or sync properly

### 3. Missing `enviro_pack` coverage type in UI dropdown
- **File**: `templates/proposal_web.html`
- The generator supports `enviro_pack` but the UI dropdown doesn't include it
- **Fix**: Add to dropdown

### 4. Missing `workers_comp` key handling in premium summary
- **File**: `proposal_generator.py` line 802+
- The `all_keys` list in `generate_premium_summary` includes both `workers_comp` and `workers_compensation`, but the data typically only has one — can cause duplicate rows or missed rows

## UI/UX Improvements

### 5. No expiring carrier/details input
- Users can enter expiring premiums but not expiring carrier names
- The generator supports `expiring_details` with carrier info for comparison mode
- **Fix**: Add expiring carrier input fields

### 6. No deductible add button
- Users can add limits but there's no "Add Deductible" button
- **Fix**: Add `addCovDed()` function and button

### 7. No sublimit/additional coverage add button
- Users can view but not add additional coverages/sublimits
- **Fix**: Add `addAddlCov()` function and button

### 8. No subjectivity add button
- Users can view but not add subjectivities
- **Fix**: Add `addSubj()` function and button

### 9. Additional coverages field name mismatch
- The `renderCoverageFields` reads `ac.coverage || ac.name` but `updateAddlCov` writes to `coverage`
- The generator reads `ac.description` — field name mismatch
- **Fix**: Align field names to use `description` consistently

## Extraction Improvements

### 10. Truncation at 200K chars may lose data
- Large multi-property SOVs can exceed 200K chars
- **Fix**: Increase to 300K or implement smarter truncation (prioritize quote text over SOV)

### 11. Pass 2-4 send the ENTIRE combined_text each time
- Each focused pass sends the full document text, which is wasteful for gpt-4.1-mini
- Could be optimized to send only relevant sections
- **Note**: Low priority — works correctly, just costs more tokens

## DOCX Generation Improvements

### 12. `additional_coverages` field name inconsistency in sublimits
- The `renderCoverageFields` UI reads `ac.coverage` but the generator reads `ac.get("description")`
- **Fix**: Ensure consistent field naming

### 13. Missing `enviro_pack` in coverage section generation
- `generate_proposal()` doesn't have an `if "enviro_pack" in coverages` block
- **Fix**: Add enviro_pack to the coverage section generation

### 14. Confirmation to Bind section missing effective date
- The bind confirmation section should include the effective date prominently

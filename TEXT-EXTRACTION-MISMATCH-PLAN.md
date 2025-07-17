# Text Extraction Mismatch Resolution Plan

## Issue Summary
The text extraction for `samples/other.doc` produces incorrect output with text concatenation and duplication issues.

**Expected Output:**
```
Big line
Bold simple line
Italic line
```

**Actual Output:**
```
Big line
Bold simple lineItalic lineBig line
Italic line
```

## Root Cause Analysis

### Current Understanding
1. **CHPX Processing Issue**: Character Property Exceptions (CHPX) are not correctly mapping to their corresponding text ranges
2. **FC to CP Mapping Problem**: File Character (FC) positions are not accurately translating to Character Position (CP) ranges
3. **Paragraph Boundary Confusion**: The system is not properly respecting paragraph boundaries during text extraction
4. **Fallback Logic Overreach**: Current fallback mechanisms are causing text duplication

### Technical Details
- Document uses fast-save format (`_doc.FIB.cQuickSaves > 0`)
- Piece table has complex multi-piece structure
- Current fallback logic in `DocumentMapping.cs` line 634+ is too broad
- Text extraction spans multiple CHPXs incorrectly

## Systematic Resolution Plan

### Phase 1: Baseline Analysis
**Goal**: Understand the current state and establish clean baseline

#### 1.1 Document Current State
- [x] Revert all previous changes to `DocumentMapping.cs`
- [x] Test baseline extraction for `other.doc` to confirm original issue
- [x] Verify regression file `47950_normal.doc` works correctly in baseline
- [x] Document exact character counts and positions in expected vs actual output

**BASELINE ANALYSIS RESULTS:**

**For `other.doc`:**
- **Expected Output** (from `other.expected.txt`): 
  ```
  Big line
  Bold simple line
  Italic line
  ```
  - Characters: 35 total (including line breaks)
  - Lines: 3 content lines + 1 empty line at end
  - Line 1: "Big line" (8 chars)
  - Line 2: "Bold simple line" (16 chars)  
  - Line 3: "Italic line" (11 chars)

- **Actual Baseline Output** (from `other.txt`):
  ```
  Big line
  ```
  - Characters: 10 total (including line breaks)
  - Lines: 1 content line + 2 empty lines
  - Line 1: Empty
  - Line 2: "Big line" (8 chars)
  - Line 3: Empty

**Issue Confirmed**: Only first text segment extracted, missing "Bold simple line" and "Italic line"

**For `47950_normal.doc` (regression check):**
- **Expected Output**: "This is a sample Word document"
- **Actual Baseline Output**: "This is a sample Word document"
- **Status**: ✅ WORKING CORRECTLY (no regression)

#### 1.2 Deep Document Structure Analysis
- [x] Examine `other.doc` piece table structure
- [x] Map CHPX ranges to their expected text content
- [x] Identify FC positions for each text segment
- [x] Document paragraph boundaries and their CP positions

**DIAGNOSTIC OUTPUT ANALYSIS:**

**Document Structure Analysis for `other.doc`:**
```
PARAGRAPH: CP(0-9) -> FC(3106-1024)
CHPX[0]: FC(3106-1024) -> 0 chars: ''

PARAGRAPH: CP(9-26) -> FC(1024-3124)
CHPX[0]: FC(1024-1029) -> 0 chars: ''
CHPX[1]: FC(1029-1035) -> 0 chars: ''
CHPX[2]: FC(1035-1046) -> 0 chars: ''
CHPX[3]: FC(1046-3112) -> 3 chars: 'Big'
CHPX[4]: FC(3112-3124) -> 6 chars: ' line\'

PARAGRAPH: CP(26-38) -> FC(3124-1046)
CHPX[0]: FC(3124-1046) -> 0 chars: ''
```

**KEY FINDINGS:**
1. **Paragraph 1**: CP(0-9) contains only empty CHPX - should contain "Big line"
2. **Paragraph 2**: CP(9-26) contains the text "Big line" but should contain "Bold simple line"  
3. **Paragraph 3**: CP(26-38) contains empty CHPX - should contain "Italic line"

**CRITICAL ISSUE IDENTIFIED:**
- Multiple CHPXs in paragraph 2 return 0 characters (FC ranges 1024-1029, 1029-1035, 1035-1046)
- Only CHPX[3] and CHPX[4] extract text: "Big" + " line"
- The FC ranges appear to be inconsistent (FC 3106-1024 is negative range)
- Missing text segments suggest piece table navigation issues

### Phase 2: Targeted Diagnostics
**Goal**: Pinpoint exact failure points in text extraction

#### 2.1 CHPX Analysis
- [ ] Add detailed logging to `writeParagraph()` method
- [ ] Track each CHPX's FC range and extracted character count
- [ ] Identify which CHPX extractions are failing (returning 0 chars)
- [ ] Map failing ranges to expected text content

#### 2.2 Piece Table Investigation
- [ ] Examine piece table entries for `other.doc`
- [ ] Verify FC to CP mappings for problematic ranges
- [ ] Check for overlapping or missing piece table entries
- [ ] Document piece encoding types (Unicode vs ANSI)

#### 2.3 Paragraph Boundary Analysis
- [ ] Identify paragraph end markers in the text stream
- [ ] Map paragraph boundaries to CP positions
- [ ] Verify PAPX (Paragraph Property Exceptions) ranges
- [ ] Check for paragraph structure corruption

### Phase 3: Precision Fixes
**Goal**: Implement targeted fixes without breaking existing functionality

#### 3.1 CHPX Range Validation
- [ ] Implement bounds checking for FC ranges
- [ ] Add validation for CHPX to text mapping
- [ ] Create safe fallback for invalid FC ranges
- [ ] Ensure no character duplication across CHPXs

#### 3.2 Enhanced FC to CP Mapping
- [ ] Improve piece table navigation for complex documents
- [ ] Add robust handling for non-contiguous FC ranges
- [ ] Implement precise character range extraction
- [ ] Handle encoding transitions correctly

#### 3.3 Paragraph Boundary Preservation
- [ ] Ensure paragraph markers are correctly identified
- [ ] Prevent text bleeding between paragraphs
- [ ] Maintain proper line break structure
- [ ] Handle paragraph formatting inheritance

### Phase 4: Implementation Strategy
**Goal**: Systematic implementation with testing at each step

#### 4.1 Minimal Viable Fix
- [ ] Implement smallest possible change to fix empty CHPX extraction
- [ ] Focus only on the specific failing FC ranges
- [ ] Test with `other.doc` after each micro-change
- [ ] Ensure no regression in `47950_normal.doc`

#### 4.2 Iterative Refinement
- [ ] Test fix with multiple sample documents
- [ ] Verify no new text duplication issues
- [ ] Check paragraph boundary preservation
- [ ] Validate character position accuracy

#### 4.3 Edge Case Handling
- [ ] Test with other complex documents (`complex_sections.doc`)
- [ ] Handle Unicode vs ANSI encoding transitions
- [ ] Manage overlapping CHPX ranges
- [ ] Address piece table fragmentation

### Phase 5: Validation and Testing
**Goal**: Comprehensive validation of the fix

#### 5.1 Core Functionality Tests
- [ ] Verify `other.doc` produces exact expected output
- [ ] Test regression file `47950_normal.doc` remains correct
- [ ] Run basic sample files (`simple.doc`, `simple-table.doc`, `bug65255.doc`)
- [ ] Execute unit tests suite

#### 5.2 Extended Validation
- [ ] Test with additional complex documents
- [ ] Verify no performance regression
- [ ] Check memory usage patterns
- [ ] Validate with corrupted/edge-case documents

#### 5.3 Integration Testing
- [ ] Run integration test suite
- [ ] Compare results with baseline before changes
- [ ] Document any acceptable behavior changes
- [ ] Verify no new failures introduced

## Success Criteria

### Primary Success Metrics
1. **Exact Match**: `other.doc` output matches `other.expected.txt` exactly
2. **No Regression**: All existing working files continue to work
3. **Clean Structure**: No text duplication or concatenation artifacts
4. **Proper Formatting**: Paragraph boundaries preserved correctly

### Secondary Success Metrics
1. **Performance**: No significant performance degradation
2. **Robustness**: Handles edge cases gracefully
3. **Maintainability**: Code changes are minimal and well-documented
4. **Test Coverage**: All changes covered by existing or new tests

## Implementation Notes

### Key Files to Modify
- `Text/TextMapping/DocumentMapping.cs` - Main text extraction logic
- `Text/TextMapping/MainDocumentMapping.cs` - Document processing flow
- Potentially: `Doc/DocFileFormat/PieceTable.cs` - Piece table handling

### Debugging Tools
- Add conditional logging for fast-saved documents
- Use character position tracking for validation
- Implement CHPX range verification
- Add piece table integrity checks

### Safety Measures
- Maintain backup of working baseline
- Test each change incrementally
- Use feature flags for new logic where possible
- Document all character position calculations

## Risk Assessment

### High Risk Areas
1. **Text Duplication**: Previous attempts caused text to appear multiple times
2. **Paragraph Structure**: Breaking paragraph boundaries affects readability
3. **Encoding Issues**: Unicode/ANSI transitions can cause corruption
4. **Performance**: Complex piece table navigation may slow processing

### Mitigation Strategies
1. **Incremental Testing**: Test after each small change
2. **Regression Prevention**: Maintain comprehensive test suite
3. **Rollback Plan**: Keep clean baseline for quick reversion
4. **Validation**: Use multiple sample documents for verification

## Timeline Estimate
- **Phase 1**: 2-3 implementation cycles
- **Phase 2**: 3-4 implementation cycles  
- **Phase 3**: 4-5 implementation cycles
- **Phase 4**: 3-4 implementation cycles
- **Phase 5**: 2-3 implementation cycles

Total: 14-19 implementation cycles

## Next Steps
1. Begin with Phase 1.1: Revert changes and establish baseline
2. Document current behavior with detailed logging
3. Analyze `other.doc` structure systematically
4. Implement minimal targeted fix
5. Validate thoroughly before considering complete
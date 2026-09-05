# Generator batch mode — implementation summary

## Locked rules

| Item | Rule |
|------|------|
| Table 1 Current rows | This run only — **no persist** — cleared on next batch |
| Table 2 Batches | **Persist** until user deletes |
| Per run | Max **50 new** rows |
| Skip | Keys already in completed batches (`doneKeys`) |
| Live UI | One row at a time in Table 1 |
| Batch file | Merged DOCX → Downloads |
| Delete batch | Removes record + frees keys (± storage file) |

## Files changed / added

| File | Role |
|------|------|
| `generator.py` | `list_data_rows`, `generate_one_row`, `merge_docx_list` |
| `index.html` | Two tables + batch generate flow |
| `GeneratorBridge.java.txt` | Java AppBridge methods to merge |
| `analyser.*` | Unchanged pattern (Analyser page) |

## Java methods to add

1. `listExcelRows()`
2. `generateOneRow(excelRowNum)`
3. `mergeAndPublishBatch(pathsJson)`

Also need existing: `publishToDownloads`, `openAnalysedFile`, `deleteStorageFile`.

## Test plan

1. Upload Excel (e.g. 120 rows) + template  
2. Generate → Table 1 shows up to 50 · Table 2 Batch 1  
3. Generate again → next 50 · Batch 2  
4. Kill app, reopen → Table 1 empty, Table 2 still there  
5. Delete Batch 1 → those keys can generate again  

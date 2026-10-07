# Workflow Studio

A local Streamlit application that reads the supplied Excel workflow catalog, routes requests with a LangGraph state graph, and executes ten business workflows against CSV/XLSX uploads or included sample data. Every workflow can run without an API key. OpenAI is an optional content-generation service for WF004 and WF007; deterministic, input-grounded templates are used when it is not configured or available.

## Run It

Requires Python 3.10 or newer.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
Copy-Item .env.example .env
streamlit run app.py
```

Open the local URL printed by Streamlit. To enable OpenAI content generation, add `OPENAI_API_KEY` and optionally `OPENAI_MODEL` to `.env`; then explicitly enable the OpenAI checkbox in WF004 or WF007. API use is limited to those workflows and is opt-in per run. Without a key, those workflows return reproducible templates. Never commit `.env`.

Run the regression tests:

```powershell
python -m unittest discover -s tests -v
```

## Project Contents

- `app.py`: Streamlit UI, request entry, optional file uploads, workflow inputs, and downloadable JSON results.
- `workflow_agent/engine.py`: Excel catalog reader, request router, LangGraph graph, tool adapters, business rules, and error states.
- `data/AI_Agent_Workflow_Assessment.xlsx`: original assignment workbook, including all ten workflow specifications and ten test prompts.
- `data/*.csv`: synthetic sample inputs for each data-backed workflow. These contain no customer or company data.
- `tests/test_workflows.py`: routing, workflow, edge-case, and WF011 extensibility checks.
- `.env.example`, `requirements.txt`, `.gitignore`: configuration and setup.

## Architecture

```text
Request + structured inputs
            |
            v
     LangGraph StateGraph
       route_request  <--- Excel Workflows sheet (loaded at runtime)
            |
       execute_workflow
            |
   workflow-specific tools
   CSV/XLSX | rules | lookup | optional OpenAI
            |
     JSON result + tool/API trace
```

The graph has two explicit nodes. The routing node scores request phrases against the workbook workflow names/triggers and a small phrase map for the supplied test prompts. A single best match proceeds to execution. No match or a tie returns a clarification result instead of guessing. The execution node dispatches to one deterministic handler and returns structured data, status, tool names, and (when applicable) an external API call trace.

This deliberately uses rules for selection and business decisions rather than asking an LLM to choose or calculate. That makes workflow selection, thresholds, and exception paths reproducible and testable. The OpenAI Responses API is isolated to copy generation; its prompt forbids adding facts not in the user-provided attributes. Missing API/configuration or an API error falls back to a local template.

## Implemented Workflows

| ID | Workflow | Main behavior and checks |
| --- | --- | --- |
| WF001 | Inventory Restock Check | Flags only `current_stock < minimum_stock`; calculates the shortfall as reorder quantity. |
| WF002 | Product Price Validation | Matches on SKU; reports missing matches; flags absolute price differences strictly above 10% by default. |
| WF003 | Vendor File Processing | Reads CSV/XLSX; normalizes headers; rejects rows missing SKU or product name; returns cleaned and invalid rows. |
| WF004 | Product Description Generator | Requires product name; generates four content fields; marks missing attributes and avoids inventing them. Optional OpenAI. |
| WF005 | Customer Order Status | Looks up by order ID or email; returns status, items, shipment, and tracking data; handles missing orders. |
| WF006 | Duplicate Product Detection | Identical non-empty SKU is definite; product-name similarity above the configurable threshold is possible. |
| WF007 | Marketing Campaign Brief | Requires goal and both dates; returns objective, audience, message, channels, timeline, and checklist. Optional OpenAI. |
| WF008 | SEO Keyword Classification | Removes duplicate keywords; classifies all four intents; maps each to a recommended page and priority. |
| WF009 | Employee Task Assignment | Ranks employees by required-skill match and available capacity; returns reasoning or escalates if none qualify. |
| WF010 | Workflow Performance Report | Calculates run count, failure rate, mean duration, common errors, slow steps, threshold breaches, and recommendations. |

## Inputs and Decisions

Use **Auto-detect from request** to let the router select a workflow, or select a workflow ID directly to inspect it and provide its structured inputs. For tabular data, the app uses the corresponding `data/*.csv` sample unless an upload is supplied. CSV and `.xlsx` uploads are supported.

`WF002` calculates difference as `abs(vendor_price - internal_price) / internal_price * 100`; exactly 10% is not flagged. `WF001` uses each row's minimum unless the user explicitly overrides it. `WF006`'s name score is a string similarity heuristic, not a guarantee that two records represent the same item. `WF008` uses transparent keyword rules; review classifications before publishing. `WF010` defaults to a slow-step threshold of 30 seconds; this is an example operating threshold, adjustable in the UI.

The current employee ranking chooses a worker who has at least one requested skill and positive remaining capacity. This is a recommendation, not an automatic assignment to a live HR/task system. Order lookup and all other data tools use local fixture/upload data; no production database is connected.

## Workbook and Adding Workflow 11

The workbook's `Workflows` sheet is loaded at runtime; it is not duplicated as hard-coded workflow prose. Append a row with the same eight columns (`Workflow_ID`, `Workflow_Name`, `Trigger`, `Inputs`, `Steps`, `Decision_Logic`, `Tools_Required`, `Expected_Output`) and the UI will list it without a UI rewrite. Unrecognized IDs can already return a generic plan-only result using those workbook details. To automate the new business action, add a specialized branch in `_run_workflow` and its input controls in `collect_inputs`, then add tests. The test suite appends a sample WF011 to a temporary workbook and verifies catalog discovery and plan generation.

## Example Requests and Results

These ten prompts are copied from the `Test_Questions` sheet and also appear in the sidebar sample-request button:

| Workflow | Example request |
| --- | --- |
| WF001 | Which products need restocking? |
| WF002 | Find products where vendor price differs by more than 10%. |
| WF003 | Process this vendor spreadsheet and show invalid rows. |
| WF004 | Generate SEO content for this product. |
| WF005 | Where is order ORD-1001? |
| WF006 | Find likely duplicate products in the catalog. |
| WF007 | Create a campaign brief for the new collection. |
| WF008 | Classify these keywords and map them to pages. |
| WF009 | Assign this urgent task to the best available developer. |
| WF010 | Which workflows are failing most often? |

With the included fixtures, expected results include:

- WF001: two restock items; `MUG-101` needs 8 units and `TOTE-330` needs 6. The 8-unit `NOTE-412` is not flagged because it equals its threshold.
- WF002: three matched products and two exceptions (`MUG-101` at 16.67%, `TOTE-330` at 11.11%).
- WF003: 4 input rows, 2 valid and 2 invalid (one missing SKU, one missing product name).
- WF004: the request alone has no product name, so the workflow asks for it; with only a name supplied, the generated result explicitly lists missing attributes.
- WF005: `ORD-1001` is shipped/in transit and has tracking number `1Z999AA10123456784`; an unknown ID asks for another identifier.
- WF006: exact SKU duplicates are definite; similar product names are possible and include a similarity score.
- WF007: the sample request asks for a brief but omits campaign goal and dates, so the workflow requests them before generation.
- WF008: 5 unique keywords remain after de-duplication and include all four intent classes.
- WF009: Maya Chen is recommended for a Python/backend task; a request for an unavailable skill escalates.
- WF010: 10 runs, 2 failures (20%), and the 33-second WF006 step is flagged against the 30-second sample limit.

Full result JSON is downloadable from each run in the app.

## Loom Walkthrough Outline

Suggested 4-6 minute live demo: start with the workflow catalog in the sidebar and show the original workbook in `data/`; explain that the LangGraph router reads its definitions and returns clarification when a request is ambiguous. Run WF001 to show automatic selection, fixture loading, the threshold comparison, output, and tool trace. Run WF005 with `Where is order ORD-1001?` to show request extraction, lookup, shipment details, and then an unknown ID to demonstrate error handling. Run WF003 to show the valid/invalid row split. Run WF004 without a product name to demonstrate the missing-data guard, then with a name to show content generation and the OpenAI opt-in/fallback. Run WF007 without dates to demonstrate its condition. Finish with WF009's skill/capacity recommendation and escalation case, WF010's report, and the README's WF011 extension point (a workbook row registers immediately; a specialized handler adds automation). Mention that WF002, WF006, and WF008 decisions are deterministic and covered by the test suite.

## Evaluation Notes

The app is local-first and uses synthetic sample data. It does not connect to real order, vendor, employee, or workflow-log systems. Validate uploaded business data before relying on results. The included tests exercise each listed business rule and the ten workbook-provided routing prompts.
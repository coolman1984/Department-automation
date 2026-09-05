# Ultimate Excel Automation V10
## Unified Approved Skill, Product Constitution, and Implementation Master Plan

Status: Planning authority for review  
Date: 2026-08-18  
Language: English  
Execution state: PLAN ONLY — repository implementation is not authorized by this document alone

---

# 0. Purpose and authority

This document defines one coherent direction for the reusable Excel automation product.

It reconciles:

- the original reusable-template vision;
- the detailed Excel automation architecture;
- the zero-based V9 re-audit;
- the current Perfect-Project-Template repository;
- the approved local-first and optional-server direction;
- optional Data Hub, RPA, database, API, and Gauss integration;
- the requirement that a non-technical employee can receive a ready project and use an AI coding agent to adapt it with minimal changes.

After user approval, this document becomes the planning authority for later documentation and repository changes.

Until that approval and a separate implementation instruction:

- do not merge branches;
- do not rewrite repository code;
- do not seal or publish a runtime;
- do not delete working compatibility paths;
- do not claim employee readiness.

When older plans, addenda, prompts, diagrams, repository documents, or code comments conflict with this document, the conflict must be recorded and resolved deliberately. No agent may silently choose the most convenient interpretation.

---

# 1. Product promise

The product is a complete, reusable Excel automation application template.

A non-technical employee receives a copy containing the technical foundation already built:

- protected Excel extraction;
- data staging;
- database;
- history;
- quality controls;
- reconciliation;
- analytics;
- dashboards;
- exports;
- offline runtime;
- logs;
- recovery;
- tests;
- AI-agent guidance.

The employee then provides:

- the Excel files used by the business process;
- what each source means;
- what one row means;
- business keys;
- correction and update behavior;
- relationships between sources;
- trusted totals;
- KPIs;
- exceptions;
- desired decisions and outputs;
- required business approvals.

The AI coding agent:

1. understands the files and business meaning;
2. reuses the existing application;
3. changes configuration first;
4. adds isolated project-specific logic only where needed;
5. changes shared engine code only when genuinely necessary;
6. tests the result;
7. leaves one understandable, working, unique project.

Routine future operation must not require the coding agent unless business logic, source meaning, or source structure changes.

---

# 2. Approved mental model

~~~text
READY REUSABLE APPLICATION TEMPLATE
        technical foundation already built
                       |
              copied for one project
                       |
       employee provides files + business meaning
                       |
             AI ADAPTATION AGENT
                       |
       REUSE → CONFIGURE → UNIQUE LOGIC
                       |
      SMALL CORE CHANGE ONLY IF REQUIRED
                       |
            READY UNIQUE PROJECT
                       |
      ROUTINE RUNS WITHOUT CODING AGENT
~~~

The template is already the application.

The employee project is a normal project, not an opaque plug-in pack and not a second product platform.

The project may still use clear extension points, contracts, and reusable modules. Those are engineering tools, not a reason to turn ordinary adaptation into a complex distribution framework.

---

# 3. Approved decisions

| ID | Decision | Approved rule |
|---|---|---|
| D-001 | Primary runtime | Local Windows runtime is the primary mode for protected Excel files |
| D-002 | Server runtime | Company-server execution is an optional deployment mode, not a universal prerequisite |
| D-003 | Core changes | Core changes are rare, justified, minimal, reusable, and tested; they are not absolutely forbidden |
| D-004 | Adaptation agent | The coding agent adapts the project during setup or later change work |
| D-005 | Routine operation | Daily, weekly, and monthly runs do not require the coding agent |
| D-006 | Data Hub and RPA | Optional acquisition adapters; not mandatory architecture for every department |
| D-007 | Gauss | Optional reasoning, explanation, and review connector; never the trusted calculation engine |
| D-008 | Trusted calculation | Deterministic SQL first; isolated Python through one explicit interface when justified |
| D-009 | Browser role | Presentation and interaction only; never the authority for trusted calculations |
| D-010 | Reuse target | Reuse 70–80% or more of the technical foundation where practical; measure evidence, do not invent a flattering number |
| D-011 | Rebuild rule | Avoid unnecessary rebuilds; rebuild the smallest required package when code or dependencies genuinely require it |
| D-012 | Project intelligence | Keep a small map-first context system with one source per concept |
| D-013 | Employee questions | Ask business questions only; technical and security architecture belongs to the template and IT |
| D-014 | Product scope | Finish the reusable local product before optional enterprise connectors dominate the roadmap |

---

# 4. What 70–80% reuse means

The target applies to the reusable technical foundation, not to unknown business meaning.

Normally reused:

- extraction framework;
- source file protection;
- chunking and lineage;
- raw staging;
- database connection and transactions;
- history modes;
- quality engine;
- reconciliation framework;
- quarantine;
- event and progress system;
- dashboard framework;
- export framework;
- local web runtime;
- security boundary;
- packaging;
- logs and run manifests;
- generic tests;
- map and context tools.

Normally adapted:

- source roles;
- file and sheet matching;
- field mappings;
- types;
- keys;
- source relationships;
- update behavior;
- quality thresholds;
- trusted controls;
- business calculations;
- KPIs;
- charts;
- insights;
- output definitions;
- project-specific tests.

The reuse report should show:

- shared files changed;
- configuration files changed;
- project business-rule files added or changed;
- reusable modules used;
- generic tests reused;
- project tests added;
- new technical architecture decisions;
- selected context files and bytes;
- runtime rebuild required or avoided;
- reason for every shared-engine change.

A percentage may be shown as a rough summary, but it is never the main acceptance proof.

---

# 5. Two separate lifecycles

## 5.1 Adaptation lifecycle

This happens once during project creation and again only when the business changes.

~~~text
Receive template copy
→ collect representative source files
→ interview the business user
→ profile files safely
→ confirm source meaning, keys, relationships, controls and outputs
→ search reusable capabilities
→ configure the existing project
→ add unique SQL/Python only where necessary
→ make a small core improvement only if required
→ test
→ obtain business approval
→ package the unique project
~~~

## 5.2 Routine operating lifecycle

This happens daily, weekly, or monthly without a coding agent.

~~~text
Acquire approved source files
→ verify identity and completeness
→ extract
→ validate
→ standardize
→ load staging
→ update trusted history transactionally
→ calculate trusted KPIs
→ publish dashboard and reports
→ optionally request Gauss review
→ record logs, quality, lineage and run manifest
~~~

The coding agent returns only when:

- a source schema changes;
- a new source is introduced;
- business meaning changes;
- a KPI changes;
- a new output is requested;
- an unsupported edge case is proven;
- a defect requires repair.

---

# 6. Decision ownership

## 6.1 The employee or business owner decides

- business purpose;
- source meaning;
- row meaning;
- business key;
- source precedence;
- correction behavior;
- deletion meaning;
- trusted totals;
- KPI meaning;
- rounding, sign, period, unit and currency rules;
- exception meaning;
- target decisions;
- output audience;
- business approval.

## 6.2 IT and Security decide

- approved storage locations;
- retention;
- access rights;
- protected-data handling;
- external AI disclosure;
- company-server availability;
- database and API credentials;
- service accounts;
- scheduling policy;
- network and proxy rules;
- approved deployment mode.

## 6.3 The reusable template decides by default

- database architecture;
- local API architecture;
- folder conventions;
- transaction strategy;
- error and run-state model;
- logging;
- packaging;
- quality framework;
- history framework;
- browser framework;
- test framework;
- recovery framework.

## 6.4 The AI adaptation agent decides

- how to map approved meaning into existing configuration;
- which reusable capability to use;
- the smallest correct implementation;
- which tests prove it;
- whether an apparent new requirement can be solved without core changes;
- whether a minimal runtime rebuild is required.

The agent must never invent business meaning or approve its own business interpretation.

---

# 7. Canonical architecture

## 7.1 Shared processing architecture

~~~text
SOURCE ACQUISITION
manual upload / watched folder / Data Hub / RPA / database / API
                         |
                         v
AUTHORIZED EXTRACTION OR CONNECTOR
                         |
                         v
RAW STAGING + LINEAGE
                         |
                         v
QUALITY + RECONCILIATION + QUARANTINE
                         |
                         v
TYPED CLEAN DATA
                         |
                         v
TRANSACTIONAL HISTORY
                         |
                         v
TRUSTED SQL / APPROVED PYTHON RULE
                         |
                         v
DASHBOARD PACKAGE + REPORTS + EVIDENCE
                         |
                         v
OPTIONAL GAUSS REVIEW
~~~

Acquisition, processing, presentation, and optional AI review are separate responsibilities.

## 7.2 Local mode — primary

Use when:

- Excel files are DRM-protected;
- authorized desktop Excel is required;
- processing must remain on the employee PC;
- no company server is available;
- the network is closed;
- offline operation is required.

The local application:

- runs under the logged-in standard user;
- uses authorized Excel desktop automation;
- binds its local API only to loopback;
- stores trusted data locally in the approved location;
- uses no runtime internet or package download;
- publishes the local dashboard and approved exports.

## 7.3 Company-server mode — optional

Use only when IT confirms:

- the server can legally and technically access the source;
- protected files do not require an unavailable interactive user session;
- credentials and scheduling are approved;
- storage and retention are approved;
- the server runtime contains all required dependencies;
- monitoring and recovery exist.

Server mode reuses the same logical pipeline and contracts.

It changes only:

- acquisition adapter;
- scheduler;
- storage connection;
- deployment packaging;
- operational monitoring;
- approved authentication.

It must not create a second calculation definition.

## 7.4 Protected-file boundary

If a protected workbook requires the authorized employee Excel session, a generic unattended server must not be claimed as capable of processing it.

Approved patterns are:

1. local processing on the authorized employee PC;
2. an approved Windows automation worker with an authorized interactive session;
3. Data Hub or source-system delivery of an approved unprotected data extract;
4. local extraction followed by approved transfer of minimized structured data to the server.

The exact pattern is an IT and Security decision.

---

# 8. Source acquisition adapters

Every acquisition method must produce the same source manifest:

- project ID;
- source ID;
- original filename or source reference;
- acquisition method;
- acquisition timestamp;
- content hash;
- expected period;
- size;
- security classification;
- lineage reference.

Supported adapters:

## 8.1 Manual upload

The employee selects or drops files into the local web application.

Use as the first working mode because it is simple, visible, and easy to prove.

## 8.2 Watched folder

The application detects new approved files in a configured local or network folder.

It must:

- avoid partial-copy processing;
- wait for file stability;
- hash the file;
- prevent duplicate runs;
- never rename or alter the source.

## 8.3 Data Hub

Data Hub is an optional source-delivery channel.

It may:

- publish files;
- expose an API;
- provide a database table;
- provide a governed folder.

It does not replace quality, history, calculations, or business controls.

## 8.4 RPA

RPA may download or move approved source files into the intake area.

It must not:

- perform trusted calculations;
- silently edit source data;
- bypass source identity;
- hide failed downloads;
- mark a project successful.

## 8.5 Database or API

Database and API connectors must:

- use approved credentials outside source code;
- use typed contracts;
- support checkpointing or paging;
- record query/version identity;
- reconcile extracted counts and totals;
- fail closed on incomplete data.

---

# 9. Excel extraction

For protected corporate workbooks:

~~~text
authorized logged-in user
→ Microsoft Excel desktop
→ controlled automation
→ rectangular Value2 reads
→ local raw staging
~~~

Mandatory rules:

- preserve the source file byte-for-byte;
- record source hash before and after;
- use a dedicated Excel instance where practical;
- open read-only;
- never use cell-by-cell extraction for large ranges;
- detect tables, used ranges, headers and totals deliberately;
- read only required sheets and columns after discovery;
- use adaptive blocks based on total cells and memory;
- close workbooks and Excel safely;
- never silently fall back from protected production extraction to a fixture adapter.

Discovery priority:

1. approved table or named range;
2. approved configured range;
3. approved header and data area;
4. carefully detected candidate structure requiring confirmation.

Value conversion must explicitly handle:

- dates and serial dates;
- decimal and thousands separators;
- percentages;
- currency symbols;
- blanks and whitespace;
- identifiers with leading zeroes;
- large identifiers;
- errors;
- formulas and cached values;
- hidden rows or columns when business-relevant;
- merged cells;
- totals rows.

---

# 10. Multi-source project contract

A project may contain:

- one or many files;
- one or many sheets;
- transactions;
- snapshots;
- master data;
- targets;
- corrections;
- reference data;
- one or many outputs.

Each source declares:

- source ID;
- business role;
- required or optional;
- file patterns;
- allowed extensions;
- sheet/table/range discovery;
- row meaning;
- typed columns;
- aliases;
- business key;
- event or snapshot date;
- update mode;
- lookback;
- deletion rule;
- quality rules;
- trusted totals;
- expected period;
- freshness rule.

Relationships declare:

- relationship ID;
- left source;
- right source;
- left keys;
- right keys;
- cardinality;
- join type;
- whether every row must match;
- source precedence;
- approval state.

Relationships are never inferred solely because two columns have similar names.

---

# 11. Canonical deterministic pipeline

Every run follows this order:

1. discover and snapshot inputs;
2. validate project and source readiness;
3. hash and register source identities;
4. open approved extraction adapters;
5. extract into raw staging with lineage;
6. validate structure;
7. normalize fields while preserving raw values;
8. apply row and dataset quality rules;
9. quarantine rejected rows;
10. reconcile population and control totals;
11. create typed clean data;
12. validate relationships;
13. update trusted history inside one safe project transaction;
14. run trusted calculations;
15. create evidence-backed insights;
16. build dashboard and report outputs in temporary locations;
17. verify numerical, structural and visual results;
18. atomically publish;
19. preserve the last-good result on failure;
20. write logs, exceptions, metrics and run manifest.

The same approved inputs, configuration, and code version must produce the same trusted business result.

---

# 12. History and correction behavior

The reusable engine supports:

- append;
- upsert;
- snapshot;
- replace period.

Each source chooses independently.

The project must define:

- business key;
- row-content identity;
- lookback window;
- correction behavior;
- deletion behavior;
- period completeness;
- late-arrival behavior;
- backdated-file behavior.

Mandatory properties:

- idempotent reruns;
- no duplicate business history;
- late corrections applied predictably;
- exact duplicate detection;
- partial periods cannot replace complete periods silently;
- failure rolls back all trusted source changes;
- archive rebuild can reproduce trusted state when policy permits.

---

# 13. Quality, reconciliation, and quarantine

Canonical quality outcomes:

~~~text
PASS     trusted checks passed
WARNING  visible anomaly that does not invalidate trusted publication
BLOCK    trusted data must not be committed or published
~~~

A BLOCK outcome causes the run to fail. Do not create a second quality word such as FAIL.

Required control families:

- file identity;
- required source readiness;
- required columns;
- type validity;
- candidate and approved keys;
- duplicates;
- nulls;
- date boundaries;
- period completeness;
- sign and range rules;
- category validity;
- relationship cardinality;
- unmatched master/reference keys;
- freshness;
- row population equation;
- trusted control totals;
- cross-source reconciliation.

Population equation:

~~~text
source rows
= accepted rows
+ rejected rows
+ intentionally filtered rows
~~~

No row may disappear silently.

Quarantine records:

- project;
- run;
- source;
- file;
- sheet;
- row identity;
- reason code;
- severity;
- safe support detail;
- raw value only when policy allows.

---

# 14. Trusted business logic

## 14.1 Default rule order

~~~text
configuration
→ existing reusable SQL pattern
→ project-owned SQL
→ approved project-owned Python rule
→ minimal reusable engine change
~~~

## 14.2 SQL

Use deterministic, versioned SQL for:

- joins;
- filters;
- aggregations;
- ratios;
- variances;
- ranks;
- aging;
- Pareto;
- period comparison;
- most KPI calculations.

The trusted formula exists once.

## 14.3 Python extension contract

The reusable runtime must provide one explicit project-rule runner.

Project Python rules are allowed only when:

- SQL is materially unclear, unsafe, or unsuitable;
- the rule uses the approved interface;
- inputs and outputs are declared;
- dependencies are already bundled or trigger a controlled rebuild;
- tests exist;
- the rule cannot access network or arbitrary files;
- the rule cannot bypass quality or transaction controls.

Each Python rule declares:

- rule ID;
- version;
- input tables or datasets;
- input columns;
- output schema;
- deterministic behavior;
- allowed dependencies;
- error behavior;
- tests.

The runtime loads only rules explicitly listed in project configuration.

If the current packaged runtime cannot safely load a required rule:

- do not pretend no rebuild is needed;
- rebuild the smallest required application package;
- lock the dependency;
- run regression tests;
- record the reason.

## 14.4 Forbidden duplication

Never implement the same trusted formula independently in:

- SQL;
- Python;
- browser JavaScript;
- Gauss prompt;
- Excel formula.

One implementation is authoritative. Other surfaces display or explain its result.

---

# 15. Project-specific data schema and migrations

Initial project tables are created from approved source configuration:

- raw tables;
- clean typed tables;
- trusted history tables;
- required keys and indexes;
- lineage fields.

Schema changes after first operation use project-owned migrations:

~~~text
projects/<project_id>/migrations/
~~~

Each migration declares:

- project ID;
- schema version from and to;
- migration ID;
- forward operation;
- validation;
- rollback or recovery approach;
- affected tables;
- required history rebuild;
- tests.

Rules:

- project migrations change only the project schema;
- shared engine migrations remain separate;
- migrations run transactionally where supported;
- schema drift never silently changes trusted tables;
- a failed migration preserves the last trusted database;
- configuration and schema versions stay aligned.

---

# 16. Adaptation surface

The normal employee project remains understandable:

~~~text
projects/<project_id>/
    project.toml
    sources.toml
    relationships.toml
    quality.toml
    metrics.toml
    dashboard.toml
    output.toml
    setup_answers.json
    business_rules/
        metrics.sql
        insights.sql
        python/
    migrations/
    presentation/
    tests/
~~~

Not every project needs every optional file.

The generator is a convenience tool that creates the project-owned surface in a discovery state. It is not a separate plug-in architecture and it must not copy assumptions from a reference department.

References demonstrate patterns. They do not supply unknown business meaning.

---

# 17. Core Change Guard

The Core Change Guard is a warning, justification, and proof gate.

It is not an immutable ownership firewall.

Required decision sequence:

~~~text
Can configuration solve it?
→ Can an existing reusable capability solve it?
→ Can isolated project logic solve it?
→ Is a shared-engine change genuinely required?
~~~

If a shared-engine change is required:

1. explain the limitation;
2. identify affected shared files;
3. make the smallest reusable change;
4. avoid project values in shared code;
5. add regression tests;
6. run existing reference projects;
7. record runtime rebuild impact;
8. record the change in the adaptation report.

Target for ordinary projects: zero shared-engine changes.

This target is not an absolute prohibition.

If the change occurs in an employee copy:

- finish the employee project safely;
- record the reusable lesson;
- optionally submit it for later master-template review;
- do not require a central promotion platform to complete the employee project.

If the change occurs in the master-template repository:

- generalize it after proof;
- update reusable documentation and tests;
- issue a new approved release when ready.

---

# 18. Rebuild and packaging decision

| Change type | Runtime action |
|---|---|
| Configuration, mappings, SQL, dashboard configuration | Reuse existing runtime |
| Python rule supported by the existing external-rule interface and existing bundled dependencies | Reuse existing runtime |
| New dependency, incompatible Python extension, packaged static code, core code | Rebuild the minimum required package |
| Shared runtime or security boundary change | Full affected regression and release verification |

The goal is to avoid unnecessary rebuilds, not to ban every rebuild.

The employee must never be asked to:

- install Python;
- run pip;
- install Node;
- run npm;
- use Git;
- open a terminal;
- compile the application;
- download packages at runtime.

Developer and AI-agent rebuild tools may use the approved offline build kit.

---

# 19. Dashboard, reports, and useful insights

The reusable presentation framework supports:

- KPI cards;
- trends;
- comparisons;
- bar and ranking charts;
- Pareto;
- heatmap;
- details table;
- quality and reconciliation;
- run history;
- evidence-backed insights;
- exports;
- print;
- light and dark themes;
- English and Arabic;
- right-to-left layout;
- keyboard use;
- reduced motion.

Dashboard configuration declares:

- title;
- filters;
- KPI binding;
- chart type;
- dimensions;
- measures;
- target or comparison;
- units and formats;
- drill-through;
- table columns;
- insight binding;
- export behavior.

Trusted numbers come from the processing layer.

The browser:

- renders;
- filters an approved dashboard package;
- drills into provided detail;
- exports the current approved context.

The browser does not create a second trusted calculation engine.

---

# 20. Evidence-backed insights

Every insight must link to verified evidence:

- metric ID;
- SQL or approved rule version;
- period;
- current value;
- comparison;
- absolute change;
- percentage change;
- relevant dimension;
- supporting rows or query reference;
- confidence;
- quality state.

Insights may state:

- what changed;
- where it changed;
- how large the change is;
- which contributor is largest;
- which exception needs investigation.

Insights must not invent:

- root cause;
- owner;
- financial impact;
- future forecast;
- recommended action without evidence.

Use language such as investigation priority when the cause is not proven.

---

# 21. Gauss connector

Gauss is an optional reasoning and review layer after deterministic processing.

Approved responsibilities:

- explain verified results;
- summarize evidence;
- identify unusual patterns;
- suggest investigation questions;
- review whether the output addresses the requested business decision;
- help draft executive narrative.

Forbidden responsibilities:

- calculate the trusted KPI;
- change the trusted database;
- bypass failed quality;
- mark a blocked run as passed;
- receive protected raw rows without explicit IT approval;
- invent missing business meaning;
- confirm technical success without machine evidence.

Default request package:

- approved KPI values;
- comparisons;
- evidence objects;
- quality summary;
- safe aggregates;
- business question;
- no raw protected rows.

Gauss failure normally produces:

~~~text
Trusted report available
AI review unavailable
~~~

It must not invalidate correct deterministic calculations unless the business workflow explicitly requires a separate review approval.

Authentication, endpoint, model, limits, allowed fields, logging, and retention must be confirmed from the official company contract before implementation.

---

# 22. Local web application

The local web application is the employee operating surface.

Normal experience:

~~~text
Open application
→ choose automation project
→ see required sources
→ add files or use configured intake
→ resolve only ambiguous source assignment
→ process
→ see real progress
→ review quality
→ use dashboard and reports
→ export
~~~

The application must not expose:

- database technology choices;
- local port choices;
- Python packages;
- server frameworks;
- folder architecture;
- command lines;
- raw stack traces.

Progress must come from real durable events, not decorative animation.

Failed runs preserve:

- source files;
- trusted history;
- last-good dashboard;
- failure evidence;
- recovery instructions.

---

# 23. Local security boundary

Local mode requirements:

- standard-user execution;
- no administrator rights;
- no Windows Service;
- no firewall changes;
- no URL reservation;
- no machine-wide installation;
- loopback-only local API;
- safe selected high port;
- wrong Host and Origin rejected;
- per-launch secret or equivalent local session protection;
- no wildcard cross-origin access;
- no LAN or public binding;
- no runtime internet;
- no telemetry;
- no CDN;
- no credentials in source or logs.

The browser renderer never accesses:

- Excel automation directly;
- the database directly;
- arbitrary local files;
- trusted business calculations.

It communicates through the typed local application boundary.

---

# 24. Offline delivery

Offline means:

~~~text
NO RUNTIME DOWNLOADS
~~~

It does not mean dependency-free.

The release includes:

- private runtime;
- required Python libraries;
- required native binaries;
- approved browser/chart assets;
- configuration schemas;
- database migrations;
- launchers;
- tests and verification tools;
- checksums;
- software bill of materials;
- licenses and notices;
- repair or full-replacement path;
- quick start;
- backup and recovery instructions.

Normal employee operation must work with:

- no system Python;
- no Node;
- no package manager;
- no Git;
- no editor;
- no terminal;
- no internet;
- no administrator rights.

---

# 25. Company-server deployment

Server deployment is a later optional profile.

It must not block completion of the local product.

Required server proof:

- approved operating system and runtime;
- source accessibility;
- credential handling;
- scheduler;
- concurrency control;
- project isolation;
- database isolation;
- logs and monitoring;
- retry policy;
- last-good publication;
- backup;
- recovery;
- data retention;
- security review;
- end-to-end reconciliation.

If SQL Server is enabled:

- local or staging truth must be stable first;
- write through controlled staging and transaction;
- reconcile row counts and totals;
- queue on outage;
- never re-extract Excel merely because central synchronization failed.

The same metric definition and project configuration must serve local and server profiles.

---

# 26. Minimal AI project intelligence

The project keeps only five user-visible AI intelligence artifacts:

1. PROJECT_SKILL.md  
   The operating rules and task router.

2. .ai/PROJECT_MAP.md  
   Important files, responsibilities, dependencies, tests, and change guidance.

3. .ai/CURRENT_STATE.md  
   Proven state, open blockers, active work, and exact next action.

4. .ai/CONTEXT_PACK.md  
   Generated task-ranked context for the current task.

5. .ai/ADAPTATION_REPORT.md  
   Business decisions, assumptions, changes, reuse evidence, lessons, and remaining approvals.

Rules:

- one source of truth per concept;
- generated files are marked generated;
- derived summaries are not edited manually;
- entry files point to PROJECT_SKILL.md;
- agents use the map before broad reading;
- agents still read every file they edit and every affected contract/test;
- stale maps block completion;
- large raw data never enters the context pack;
- safe representative samples are allowed only when policy permits.

The capability registry, if kept, is the machine source for reusable capability definitions. The project map references it and must not duplicate full definitions.

Internal caches and hashes may exist, but they are not separate human-maintained truth files and agents are not required to read them.

---

# 27. Privacy-conscious source profiling

Default profile:

- workbook and sheet structure;
- header candidates;
- column names;
- inferred types;
- row and column counts;
- null counts;
- distinct counts;
- candidate keys;
- date ranges when allowed;
- formula/macro/query/pivot presence;
- protection and external-link indicators.

When policy permits, add a tiny representative safe sample:

- masked;
- redacted;
- limited;
- selected to reveal formats and codes;
- never a bulk raw-data dump.

Samples are useful for:

- coded categories;
- date formats;
- mixed identifiers;
- totals rows;
- status values;
- unusual blanks;
- locale conventions.

The policy owner, not an employee checkbox alone, controls disclosure.

---

# 28. Current repository reality

Repository:

~~~text
coolman1984/Perfect-Project-Template
~~~

Observed on 2026-08-18:

- main contains the earlier executable template foundation;
- agent/universal-excel-automation-engine is substantially ahead of main;
- the repository default branch is not main;
- the agent branch contains useful project-centric multi-source implementation;
- the agent branch also contains the over-rigid sealed-core and upgrade direction;
- current-state documents, acceptance gates, and implementation evidence are not fully synchronized;
- Data Hub, RPA, Gauss, and company-server runtime are not implemented as complete product paths;
- the local application remains the strongest proven direction;
- employee distribution is not yet approved.

Until reconciliation:

- freeze merges;
- treat branch claims as evidence to verify, not final truth;
- preserve working code;
- do not delete legacy paths merely for conceptual cleanliness;
- choose one canonical branch only after the documentation and gate audit.

---

# 29. What to keep from the repository

Keep and strengthen:

- local loopback application;
- Excel extraction port and authorized COM adapter;
- raw, clean and history separation;
- DuckDB and Parquet recovery design;
- project-centric multi-source contracts;
- independent source keys and history modes;
- explicit relationships;
- transactional multi-source pipeline;
- quality PASS/WARNING/BLOCK;
- quarantine and control totals;
- configuration-driven dashboard;
- source profiling;
- map-first context tools;
- project generator in discovery mode;
- reference projects;
- failure and adversarial tests;
- offline launchers and packaging direction.

---

# 30. What to simplify or retire

Simplify or remove from the mandatory employee path:

- immutable local core firewall;
- mandatory central promotion workflow;
- mandatory upgrade packages for every employee adaptation;
- separate formal session modes that add ceremony;
- repeated capability truth across several files;
- excessive AI context artifacts;
- absolute no-rebuild language;
- empty project-pack abstractions that do not solve a proven need;
- governance gates that test governance machinery instead of the employee outcome.

Keep integrity hashes and optional upgrade tooling only as supporting release capabilities. They must not control the basic architecture or block a legitimate project-specific improvement.

---

# 31. What to add or complete

Add or complete:

- explicit project Python rule interface;
- project-owned schema migrations;
- consistent current-state and acceptance evidence;
- polished project-first setup flow;
- safe representative sampling policy;
- Data Hub adapter contract;
- RPA intake contract;
- Gauss review contract;
- optional company-server deployment profile;
- scheduled routine operation;
- one independent finance adaptation proof;
- real corporate Windows, Excel, and protected-file evidence;
- clean offline employee package;
- non-technical operator handoff proof.

---

# 32. Documentation-first implementation plan

No repository code changes begin until this plan is approved.

## Phase P0 — Freeze and inventory

Actions:

- freeze merges;
- list branches and owners;
- record current branch heads;
- inventory planning and authority documents;
- identify generated versus maintained documents;
- record known evidence.

Exit gate:

- no ambiguity about the material being reconciled.

## Phase P1 — One authority

Actions:

- approve this V10 document;
- make it the single planning authority;
- mark prior V8/V9 addenda as superseded where they conflict;
- preserve them as audit history;
- write one concise decision register.

Exit gate:

- every major product decision has one approved answer.

## Phase P2 — Repository crosswalk

Classify every important repository feature:

~~~text
KEEP
KEEP BUT SIMPLIFY
COMPLETE
MOVE TO OPTIONAL PHASE
RETIRE
UNKNOWN — NEED EVIDENCE
~~~

Exit gate:

- no code is scheduled without a reason tied to an approved decision.

## Phase P3 — Document reconciliation

Later update:

- README;
- PROJECT_SKILL;
- agent entry files;
- current state;
- project map;
- architecture document;
- security document;
- factory/adaptation guidance;
- acceptance gates;
- implementation baseline wording;
- diagram and infographic.

Exit gate:

- every document describes the same product and lifecycle.

## Phase P4 — Semantic contradiction audit

Check for contradictions involving:

- local versus server;
- adaptation versus routine operation;
- core warning versus core prohibition;
- SQL versus Python;
- rebuild versus no rebuild;
- project versus report;
- metadata-only versus safe samples;
- Gauss versus calculation engine;
- Data Hub versus mandatory source;
- employee decision versus IT decision;
- code status versus gate status.

Exit gate:

- no unresolved critical contradiction.

## Phase P5 — User approval

Deliver:

- revised document set;
- keep/change/retire matrix;
- updated implementation order;
- exact repository changes proposed;
- no code changes yet.

Exit gate:

- explicit user approval to begin repository implementation.

---

# 33. Repository implementation roadmap

This roadmap begins only after Phase P5 approval.

## Phase I0 — Branch and truth cleanup

- select one canonical implementation branch;
- preserve branch history;
- remove temporary discovery files;
- update current state from evidence;
- reconcile acceptance gates;
- run existing tests before behavior changes.

## Phase I1 — Simplify adaptation governance

- convert the core guard to warning + justification;
- make project completion possible with a minimal justified core change;
- demote central upgrade/promotion machinery from mandatory employee flow;
- simplify AI artifacts;
- preserve useful integrity checks.

## Phase I2 — Complete the project contract

- finalize source, relationship, quality, metric, output, and project-migration contracts;
- add the explicit Python rule runner;
- prove config-only and Python-extension adaptations;
- define rebuild detection.

## Phase I3 — Prove multi-source foundation

- run the Supply Chain reference end to end;
- validate independent histories;
- validate relationships;
- inject downstream failure;
- prove complete rollback;
- prove identical rerun;
- prove archive rebuild.

## Phase I4 — Prove independent adaptation

Create Finance Purchase Price Variance using:

- purchase transactions;
- standard cost master;
- vendor master;
- budget or target.

Measure:

- shared files changed;
- project configuration;
- project SQL/Python;
- tests reused;
- new tests;
- runtime rebuild;
- context size;
- operator effort.

If major shared rebuilding is required, improve the reusable abstraction and repeat.

## Phase I5 — Finish local employee experience

- project discovery;
- business questionnaire;
- source-role assignment;
- guided upload;
- durable progress;
- quality explanation;
- dashboard;
- history;
- export;
- friendly recovery.

## Phase I6 — Package and prove local offline product

- bundle runtime and assets;
- run clean offline Windows proof;
- run standard-user proof;
- run protected Excel proof;
- run two stable runs;
- run operator handoff.

This is the first employee pilot target.

## Phase I7 — Optional acquisition connectors

Implement only approved needs:

- watched folder;
- Data Hub;
- RPA;
- database;
- API.

Each adapter must pass the same source-manifest and reconciliation contract.

## Phase I8 — Optional Gauss review

- verify official API contract;
- implement data-minimized request;
- separate review output from trusted output;
- handle unavailability;
- test privacy and prompt injection boundaries.

## Phase I9 — Optional company-server mode

- confirm source eligibility;
- deploy same project contracts;
- add scheduler and operations;
- prove data isolation and reconciliation;
- preserve local mode.

---

# 34. Reference projects

The reusable engine requires four different proofs:

## Reference A — Production Quality

Proves:

- protected Excel-shaped extraction;
- defects and production controls;
- corrections;
- quality and dashboard.

## Reference B — Maintenance

Proves:

- a different department;
- downtime and duration logic;
- reuse without copying quality assumptions.

## Reference C — Supply Chain

Proves:

- transaction source;
- snapshot source;
- master source;
- different keys;
- different history modes;
- explicit relationships;
- cross-source calculations;
- atomic rollback.

## Reference D — Finance Purchase Price Variance

Proves independent adaptation after the template is stable.

It must be created through the normal employee adaptation workflow, not hand-built through hidden architecture changes.

---

# 35. Test strategy

## 35.1 Business-rule tests

Test:

- normal case;
- zero;
- negative;
- missing;
- duplicate;
- conflicting duplicate;
- rounding;
- currency;
- sign;
- period boundary;
- zero denominator;
- late correction.

## 35.2 Source tests

Test:

- missing file;
- wrong file;
- ambiguous filename;
- wrong sheet;
- renamed column;
- missing column;
- extra column;
- different header row;
- empty file;
- partially copied file;
- Arabic and Unicode path;
- long path;
- uppercase/lowercase extension.

## 35.3 Relationship tests

Test:

- missing key;
- duplicate master key;
- orphan transaction;
- wrong cardinality;
- optional unmatched row;
- required unmatched row;
- source precedence.

## 35.4 History tests

Test:

- exact rerun;
- late correction;
- snapshot change;
- disappeared record;
- backdated file;
- partial period;
- replace-period safety;
- interruption;
- multi-source failure after partial work.

## 35.5 Security tests

Test:

- path traversal;
- forged project ID;
- forged source ID;
- cross-project upload reuse;
- wrong Host;
- wrong Origin;
- launch-secret misuse;
- LAN bind;
- runtime network attempt;
- raw values entering AI context;
- credentials in logs or source.

## 35.6 Browser tests

Test:

- no external network;
- no console error;
- charts render;
- filters reconcile;
- drill-through;
- reset;
- theme;
- Arabic and right-to-left;
- print;
- keyboard;
- reduced motion;
- no trusted KPI arithmetic.

## 35.7 Packaging tests

Test:

- no system Python;
- no Node;
- no Git;
- no internet;
- no administrator rights;
- missing component fails closed;
- repair or replacement works;
- checksums and licenses exist.

---

# 36. Acceptance gates

A project is not production-ready until all applicable gates pass:

1. business meaning approved;
2. security and storage approved;
3. source files preserved;
4. protected extraction proven in the target environment;
5. source readiness proven;
6. population reconciles;
7. control totals reconcile;
8. rejected rows are visible;
9. relationships validate;
10. history is idempotent and correction-aware;
11. failures roll back;
12. trusted calculations pass golden tests;
13. dashboard equals trusted results;
14. insights link to evidence;
15. browser works offline;
16. package runs under standard user;
17. last-good output survives failure;
18. archive recovery works when enabled;
19. two stable runs pass;
20. non-technical operator completes two runs and one recovery;
21. map and current state match the repository;
22. no critical plan/code/gate contradiction remains.

Optional features have separate gates:

- Data Hub;
- RPA;
- Gauss;
- SQL Server;
- company-server scheduling.

A local pilot is not blocked merely because an optional feature is not applicable.

---

# 37. Evidence maturity

Every important claim uses:

~~~text
REQUIRED
→ CONTRACT_DEFINED
→ MACHINE_VERIFIABLE
→ REFERENCE_PROVEN
→ ENVIRONMENT_PROVEN
~~~

Definitions:

- REQUIRED: the need is recorded.
- CONTRACT_DEFINED: the behavior and interface are defined.
- MACHINE_VERIFIABLE: automated checks verify the contract.
- REFERENCE_PROVEN: a realistic reference passes end to end.
- ENVIRONMENT_PROVEN: the real target corporate environment passes.

Do not call a feature supported because:

- a document mentions it;
- a file exists;
- a command exists;
- a test is skipped;
- a mockup looks complete.

---

# 38. Error and recovery principles

Errors must tell a non-technical user:

- what failed;
- whether trusted data changed;
- what remains safe;
- what the user can do;
- when IT or the coding agent is required.

Error classes:

- user-correctable source issue;
- retryable technical issue;
- blocking data-quality issue;
- security or policy block;
- project-change issue requiring the coding agent;
- environment proof required.

Never:

- silently downgrade functionality;
- publish partial trusted output;
- overwrite last-good output;
- hide skipped files;
- discard failed rows;
- expose confidential values in logs.

---

# 39. Change control

Every material change records:

- requested outcome;
- business owner;
- affected project;
- affected sources;
- affected contract;
- classification;
- shared files changed;
- data migration;
- runtime rebuild;
- tests;
- approvals;
- rollback;
- map update.

Classification:

~~~text
REUSE_AS_IS
CONFIGURE
PROJECT_SPECIFIC_LOGIC
SMALL_SHARED_IMPROVEMENT
OPTIONAL_CONNECTOR
~~~

Do not force every novel requirement into a central capability-promotion lifecycle before the employee project can be completed.

---

# 40. Canonical source-of-truth order

After approval:

1. this V10 planning authority;
2. approved business decisions;
3. approved IT and Security policy;
4. machine contracts and schemas;
5. tested implementation;
6. current-state and acceptance evidence;
7. compatibility and historical documents.

When code and documentation disagree:

- record the conflict;
- determine the approved rule;
- update every affected source;
- add a test where possible;
- do not silently describe code as correct merely because it exists.

---

# 41. Agent operating protocol

## 41.1 Start

~~~text
Read PROJECT_SKILL.md
→ read CURRENT_STATE
→ verify PROJECT_MAP freshness
→ generate task CONTEXT_PACK
→ inspect only routed files
→ confirm business/security blockers
→ classify requirements
~~~

## 41.2 Implement

~~~text
reproduce/profile evidence
→ write or identify failing proof
→ configure first
→ add project logic if needed
→ make smallest shared change only when justified
→ run focused tests
→ run affected integration/golden/failure/browser tests
~~~

## 41.3 Finish

~~~text
verify outputs
→ update CURRENT_STATE
→ update ADAPTATION_REPORT
→ refresh PROJECT_MAP
→ verify plan/contracts/gates consistency
→ report exact proof and remaining decisions
~~~

The agent must not:

- scan the whole repository without reason;
- avoid reading a file it will edit;
- invent business answers;
- rewrite working architecture from preference;
- lower a test to obtain green status;
- claim protected-file proof from fixtures;
- implement optional future scope before the requested product works.

---

# 42. Non-technical business interview

Ask:

1. What work are you automating?
2. Which files are used?
3. What does each file contain?
4. What does one row mean?
5. Which values identify the same business record?
6. Can old records change?
7. How should removed records be treated?
8. How are the sources related?
9. Which source wins when values disagree?
10. Which totals prove the result is correct?
11. Which KPIs and exceptions matter?
12. What decision should the output help make?
13. Who approves the business meaning?
14. How often should it run?

Do not ask the employee to choose:

- DuckDB;
- SQL Server architecture;
- FastAPI;
- port;
- Python package;
- folder structure;
- encryption design;
- retention design;
- network design;
- packaging technology.

---

# 43. Final definition of done for the reusable template

The reusable template is ready for controlled employee pilot only when:

- the approved V10 direction is reflected consistently in documents and code;
- one canonical branch exists;
- project-centric multi-source execution is proven;
- Production Quality, Maintenance, and Supply Chain references pass;
- Finance adaptation demonstrates low-change reuse;
- core change is guarded but not impossible;
- project SQL and Python extension paths are proven;
- project schema migration is proven;
- local protected Excel extraction is proven on the corporate machine;
- local offline package works under a standard user;
- the employee web flow works end to end;
- trusted history and last-good output survive failure;
- two stable runs pass;
- a non-technical operator completes two runs and one recovery;
- map, current state, tests, and gates agree;
- optional server, Data Hub, RPA, Gauss, and SQL Server features are clearly marked implemented, conditional, or not applicable.

---

# 44. Final operating principle

> Build the reusable technical foundation once. Let the employee provide business meaning. Let the AI agent adapt the smallest necessary surface. Keep calculations deterministic, history safe, quality visible, and operation simple. Use local protected-file processing as the reliable foundation. Add Data Hub, RPA, Gauss, and company-server automation only as approved connectors around the same trusted engine.

---

# 45. Immediate next action

1. Review and approve this V10 planning document.
2. Produce a document-by-document reconciliation matrix.
3. Update planning and repository documentation only.
4. Run a semantic consistency audit.
5. Return the proposed repository implementation changes for approval.
6. Stop before code until the user explicitly authorizes implementation.


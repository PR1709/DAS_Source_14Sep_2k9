# 07 Migration Stack And Roadmap

## Recommended Target Stack

### Frontend

Use Next.js + React + TypeScript.

Why:

- operations-heavy application with many forms, grids, reports, and dashboards
- large ecosystem and hiring base
- can support server-rendered pages where useful and client-heavy transactional screens where needed
- TypeScript aligns with backend DTOs and generated API clients

### UI Components

Use a serious data grid: AG Grid Enterprise or MUI X Data Grid.

Why:

- the legacy app is grid-first
- must support keyboard data entry, pinned columns, sorting, filtering, exports, editable rows, and dense layouts

Use Radix/shadcn-style primitives or MUI for standard controls.

### State Management

Use TanStack Query for server state.

Use React Hook Form + Zod for forms.

Avoid broad global state. Most state should be server-backed and transaction-scoped.

### Backend

Use NestJS + TypeScript.

Why:

- strong modular architecture
- dependency injection
- DTO validation
- OpenAPI support
- background worker integration
- TypeScript sharing with frontend

Alternative:

- Spring Boot is also enterprise-grade, especially if the team prefers Java.
- For this specific project, TypeScript full-stack reduces migration surface and DTO duplication.

### API

REST first with OpenAPI.

Why:

- workflows are transactional resources: products, customers, invoices, stock movements, schemes, van loads
- easier for integration and testing
- GraphQL is not needed initially

Use SSE/WebSocket only for:

- job progress
- live sync status
- dashboard refresh

### Database

PostgreSQL.

Why:

- strong relational integrity
- transactional correctness for invoice/stock/ledger posting
- views/materialized views for reports
- good JSONB support for integration payloads and audit metadata
- mature tooling

### Background Jobs

BullMQ + Redis.

Jobs:

- DBF import
- report generation
- Tally export
- SMS sending
- HHT compatibility import/export
- stock reconciliation
- scheme simulation/recalculation

### Cache

Redis.

Use for:

- lookup caching
- job status
- rate limiting
- session-related ephemeral state

Do not use cache as source of truth.

### Auth

OIDC/OAuth2.

Options:

- Keycloak for self-hosted enterprise
- Auth0 / Entra ID if cloud identity is preferred

Need:

- RBAC
- branch/location restrictions
- audit logs
- possibly maker-checker controls for financial/stock postings

### Storage

S3-compatible object storage.

Use for:

- generated PDFs
- uploaded import files
- archived Tally payloads/responses
- backups
- report exports

### Deployment

Start with Docker Compose for local/dev and single-site pilot.

Use Kubernetes when:

- multi-location deployment
- HA requirements
- independent scaling of API/workers
- managed observability

### Testing

Required:

- unit tests for quantity conversion, pricing, schemes, tax, route rounding
- golden-master tests comparing legacy examples
- integration tests for invoice posting
- migration reconciliation tests
- Playwright tests for critical user flows
- contract tests for Tally/SMS/HHT boundaries

### CI/CD

GitHub Actions or GitLab CI:

- lint
- typecheck
- unit tests
- integration tests
- DB migration check
- Playwright smoke tests
- container build
- deployment promotion

### Observability

Use:

- OpenTelemetry
- Prometheus/Grafana
- Sentry
- structured JSON logs
- audit log table

Track:

- invoice post failures
- stock negative attempts
- scheme calculation errors
- Tally export failures
- HHT/mobile sync failures
- report job failures

## Roadmap

### Phase 1 - Discovery And Canonical Architecture

Complexity: High

Deliverables:

- full module specs for Sales/Scheme, Inventory, Van/HHT, Finance/Tally, Purchase, Master Data
- canonical ERD
- transaction type dictionary
- report catalog
- golden-master fixtures
- migration acceptance criteria

Risks:

- hidden logic in form event handlers
- ambiguous transaction type meanings
- duplicate table semantics

Mitigations:

- source-cited extraction docs
- user/domain-owner validation
- test fixtures from real historical data

### Phase 2 - Core Infrastructure

Complexity: Medium-High

Deliverables:

- monorepo
- Postgres schema
- auth/RBAC/audit
- import staging tables
- base API
- base UI shell
- background worker
- observability baseline

Risks:

- schema churn
- over-normalization too early

Mitigations:

- keep raw import schema separate
- preserve legacy IDs
- migrate module by module

### Phase 3 - Feature Migration

Complexity: Very High

Order:

1. Master data read/write
2. Product/batch/rate management
3. Read-only reports over migrated data
4. Purchase and stock receipt
5. Sales invoice without schemes
6. Scheme engine
7. Stock ledger and reconciliation
8. Empties/deposit handling
9. Van sales and HHT/mobile
10. Finance posting and Tally sync
11. Full reporting replacement

### Phase 4 - Testing And Validation

Complexity: High

Deliverables:

- row count reconciliation
- account balance reconciliation
- stock balance reconciliation
- invoice total comparison
- scheme result comparison
- UAT sign-off

### Phase 5 - Deployment

Complexity: Medium-High

Approach:

1. read-only pilot
2. dual-run selected workflows
3. cut over master data
4. cut over new transactions
5. keep legacy read-only for reference
6. retire DBF writes after reconciliation sign-off

## Architecture Improvements To Intentionally Make

- explicit domain services
- immutable stock movement ledger
- double-entry ledger entries for finance
- versioned scheme rules
- audit log for every posting
- integration outbox pattern
- normalized schema
- typed APIs
- modern report views
- offline-first field-force replacement

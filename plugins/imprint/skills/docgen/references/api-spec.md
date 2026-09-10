# API specification

An API specification is the contract between a service and the clients that call it. Its
source of truth is a machine-readable **OpenAPI 3.1** description; the document explains the
contract to people and must never disagree with it.

Standards: OpenAPI Specification 3.1 (with JSON Schema 2020-12), HTTP semantics RFC 9110,
Problem Details for HTTP APIs RFC 9457, date-time format RFC 3339, OAuth 2.0 RFC 6749 and
bearer tokens RFC 6750, link-based pagination RFC 8288, Semantic Versioning 2.0.0. Template:
`templates/api-spec.md`.

## Structure

1. Overview: what the API is for, who calls it, base URL for each environment
2. Conventions: the rules every endpoint follows (below), stated once
3. Authentication and authorisation: how to obtain a token, scopes, token lifetime
4. Resources and endpoints: one subsection per resource
5. Errors: the error format and the list of error codes
6. Versioning and deprecation: how versions change and how long an old one is supported
7. Change log

For each endpoint:

| Part | Contents |
| :-- | :-- |
| Summary | one sentence: what it does |
| Request | method and path, path and query parameters, headers, body schema |
| Responses | each status code, when it happens, its body |
| Example | one request and its response, complete and copy-pasteable |
| Rules | idempotency, rate limit, permissions needed |

Field tables:

| Field | Type | Required | Description |
| :-- | :-- | :-: | :-- |
| `amount` | string (decimal) | Yes | Order amount in the account currency, two decimals |

## Conventions to decide and state

- **Methods by meaning (RFC 9110)**: GET reads and never changes state; POST creates or
  runs an action; PUT replaces; PATCH changes part; DELETE removes. GET, PUT and DELETE are
  idempotent.
- **Status codes by meaning**: 200 with a body, 201 created (with `Location`), 204 no body,
  400 malformed request, 401 not authenticated, 403 not allowed, 404 not found, 409 conflict,
  422 valid syntax but failed validation, 429 rate limited, 5xx server fault.
- **Errors as Problem Details (RFC 9457)**: `type`, `title`, `status`, `detail`, `instance`,
  plus named extension members such as a list of invalid fields. One format for every error.
- **Names**: plural nouns for collections (`/orders`), one casing for fields everywhere
  (camelCase or snake_case), no verbs in paths except for explicit actions.
- **Dates and times**: RFC 3339 in UTC, `2026-10-01T07:30:00Z`.
- **Money**: a decimal string plus a currency code (ISO 4217), never a binary float.
- **Pagination**: cursor-based for large or changing collections; say which, the page size
  limit, and how the next page is found (a cursor field or a `Link` header, RFC 8288).
- **Idempotency**: for POST requests that move money or create records, accept an
  idempotency key header and say how long a key is remembered.
- **Versioning**: a major version in the path (`/v1`) or a header; breaking changes only in a
  new major version; additive changes are not breaking.
- **Security**: TLS only; tokens in the `Authorization: Bearer` header, never in the URL;
  say which scope each endpoint needs.

## Rules

- Generate reference tables from the OpenAPI file where possible, so the two cannot drift.
- Every example is real: it runs against the sandbox and returns what is shown.
- Mark deprecated fields and endpoints, with the version and date they will be removed.

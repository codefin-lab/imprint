---
# An API specification. OpenAPI 3.1 is the source of truth; this document explains it.
# Guidance: skills/docgen/references/api-spec.md. Build:
#   python3 build_docx.py templates/api-spec.md --pdf
document_kind: API Specification
PROJECT_NAME: "Example Orders API v1"
CLIENT_LEGAL_NAME: "Example Client Co., Ltd."
VERSION: "0.1"
DD MMM YYYY: "1 Oct 2026"
Draft / For Review / Final: "Draft"
---

# DOCUMENT CONTROL

| Version | Date | Author | Description of change |
| :-- | :-- | :-- | :-- |
| 0.1 | 2026-10-01 | API owner | First draft |

# TABLE OF CONTENTS

<!-- toc -->

# 1. OVERVIEW

The Orders API lets partner applications place and track orders. It is described by `orders-v1.openapi.yaml` (OpenAPI 3.1); where this document and that file differ, the file is correct.

| Environment | Base URL |
| :-- | :-- |
| Sandbox | `https://sandbox.api.example.com/v1` |
| Production | `https://api.example.com/v1` |

# 2. CONVENTIONS

| Topic | Rule |
| :-- | :-- |
| Transport | HTTPS only (TLS 1.2 or later) |
| Format | JSON, UTF-8; field names in camelCase |
| Dates and times | RFC 3339 in UTC, for example `2026-10-01T07:30:00Z` |
| Money | Decimal string with an ISO 4217 currency code |
| Errors | Problem Details, RFC 9457, media type `application/problem+json` |
| Pagination | Cursor-based: pass `cursor` from the previous page; at most 100 items per page |
| Idempotency | POST requests accept an `Idempotency-Key` header, remembered for 24 hours |
| Rate limit | 100 requests per minute per client; above it, status 429 with `Retry-After` |

# 3. AUTHENTICATION AND AUTHORISATION

Clients use the OAuth 2.0 client credentials grant (RFC 6749) and send the token as a bearer token (RFC 6750) in the `Authorization` header. Tokens last 1 hour.

| Scope | Allows |
| :-- | :-- |
| `orders:read` | Reading orders |
| `orders:write` | Placing and cancelling orders |

# 4. ENDPOINTS

| Method | Path | Summary | Scope |
| :-- | :-- | :-- | :-- |
| POST | `/orders` | Place an order | `orders:write` |
| GET | `/orders/{orderId}` | Get one order | `orders:read` |
| GET | `/orders` | List orders | `orders:read` |

## Place an order

`POST /orders` creates an order and returns it with status `pending`.

| Field | Type | Required | Description |
| :-- | :-- | :-: | :-- |
| `productId` | string | Yes | The product to order |
| `quantity` | integer | Yes | Number of units, 1 or more |
| `amount` | string (decimal) | Yes | Total amount, two decimals |
| `currency` | string | Yes | ISO 4217 code, for example `THB` |

Example request:

```http
POST /v1/orders HTTP/1.1
Host: sandbox.api.example.com
Authorization: Bearer <token>
Content-Type: application/json
Idempotency-Key: 5f1c2d9e-7a4b-4c61-9a2e-2f0b8e1d3c77

{"productId": "P-100", "quantity": 2, "amount": "1500.00", "currency": "THB"}
```

Example response:

```json
{
  "orderId": "O-20261001-0001",
  "status": "pending",
  "createdAt": "2026-10-01T07:30:00Z"
}
```

| Status | When | Body |
| --: | :-- | :-- |
| 201 | The order was created | The order, with a `Location` header |
| 400 | The request is not valid JSON | Problem Details |
| 401 | The token is missing or expired | Problem Details |
| 409 | The idempotency key was used with a different body | Problem Details |
| 422 | A field failed validation | Problem Details with `invalidParams` |

# 5. ERRORS

Every error uses the Problem Details format (RFC 9457):

```json
{
  "type": "https://api.example.com/problems/validation",
  "title": "Request failed validation",
  "status": 422,
  "detail": "quantity must be 1 or more",
  "invalidParams": [{"name": "quantity", "reason": "must be 1 or more"}]
}
```

| Type | Status | Meaning |
| :-- | --: | :-- |
| `validation` | 422 | A field failed validation |
| `insufficient-balance` | 409 | The account cannot cover the amount |

# 6. VERSIONING AND DEPRECATION

The major version is in the path. Adding a field or an endpoint is not a breaking change; removing or renaming one is, and happens only in a new major version. A deprecated version keeps working for at least 12 months after its successor is released.

# 7. CHANGE LOG

| Version | Date | Change |
| :-- | :-- | :-- |
| 1.0 | 2026-10-01 | First release |

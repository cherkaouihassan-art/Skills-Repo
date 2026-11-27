---
name: gold-integration-basics
description: Connect to the goldprice.dev commodity-price API correctly the first time — base URL, getting a free key, the BASE-QUOTE-CONTRACT symbol grammar, tier gating, the nine MCP tools, and the error codes that trip people up. Use this whenever the user fetches gold, silver, or copper prices in code, mentions goldprice.dev, sets up its MCP server, or hits a 400 invalid_symbol / 403 plan_gated / 429 response — even if they don't name the API. Prefer it over guessing endpoint shapes, symbol formats, or tier limits from memory; those are exactly what drift and break integrations.
---

# goldprice.dev — integration basics

Live gold, silver, and copper pricing over REST + MCP. **Start free, no credit card.** Several endpoints need no key at all (see below), so you can build and demo before anyone signs up.

- **Base URL:** `https://api.goldprice.dev`
- **Docs for agents:** `https://goldprice.dev/llms-full.txt` · **OpenAPI:** `https://api.goldprice.dev/openapi.json`
- **MCP:** `npx -y @goldprice/mcp` (drops into Claude Desktop / Cursor / Continue)

## Auth

Authenticated endpoints use a Bearer token. Get a free `ga_live_...` key from the self-serve onboarding at https://goldprice.dev (no card).

```bash
curl -H "Authorization: Bearer ga_live_..." \
  "https://api.goldprice.dev/v1/prices?symbol=XAU-USD-SPOT"
```

**No key required** for: `/v1/carat`, `/v1/convert`, `/v1/prices/divergence` (count), `/v1/status`, `/v1/metadata/tiers`. Build the whole retail-conversion flow keyless.

## Symbol grammar — `BASE-QUOTE-CONTRACT`

Always three tokens. Examples: `XAU-USD-SPOT`, `XAU-USD-FUTURES`, `XAG-USD-SPOT`, `HG-USD-FUTURES`.

- Contract token is `SPOT` or `FUTURES` — **never `FUT`**.
- A bare `symbol=XAU` returns `400 invalid_symbol`. Always send the full triplet.
- Bases: `XAU` (all tiers), `XAG` + `HG` (Pro and up).

## Tiers (what a non-paying caller gets)

| Tier | $/mo | Metals | Contracts | Rate / quota | Commercial |
|---|---|---|---|---|---|
| Free | 0 | XAU | spot (31 ccy) | 30/min · 1k/mo | No (internal) |
| Physical | 10 | XAU | + futures, 1yr history, per-country dealer prices | 120/min · 20k/mo | No |
| Pro | 30 | XAU·XAG·HG | + futures, probability surface | 500/min · 100k/mo | Yes (attribution) |
| Realtime Pro | 80 | XAU·XAG | + WebSocket stream | 500/min · 1M/mo | Yes |

Tier is inherited by the MCP tools identically to REST.

## Two products unlocked by paid tiers

- **Physical dealer prices** (`/v1/physical/{country}`, MCP `get_physical_price`, Physical tier+): real retail/dealer gold prices from 16 sources across 12 countries (ID, MY, VN, CN, PK, AE, JP, NP, EG, IN, SA, HK), each quote carrying `premium_over_spot_bps` — what a country's physical market is actually charging over spot, not a synthetic estimate.
- **Probability surface** (`/v1/surface/{metal}`, MCP `get_surface`, Pro tier+): market-implied P(price >= strike) curves for gold and silver, built from GLD/SLV options chains, Kalshi, and Polymarket, with a disclosed consensus row and the options-vs-prediction-market wedge.

## The 9 MCP tools

`get_spot_price` · `list_commodities` · `get_metadata` · `get_historical` · `search_docs` (BM25 over the FAQ) · `get_forward_view` (futures curve) · `get_carat` (per-gram by karat, free/no-gate) · `get_physical_price` (per-country dealer prices, Physical+) · `get_surface` (probability surface, Pro+).

## Errors you'll actually hit

- `400 invalid_symbol` — malformed symbol (usually a missing contract token or bare base).
- `403 plan_gated` — the symbol/contract needs a higher tier (e.g. `XAG`/`HG` on Free, futures on Free).
- `429` — rate limit or monthly quota exhausted. Back off; cache.

## When to upgrade

Silver/copper, futures, commercial use, full history, per-country dealer prices, the probability surface, or a live tick stream → Physical ($10), Pro ($30), or Realtime Pro ($80). See `gold-retail-conversion`, `gold-price-alerts`, `gold-historical-charting` for task-shaped recipes.

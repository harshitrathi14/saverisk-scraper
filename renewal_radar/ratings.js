// renewal_radar/ratings.js — lender classification + NACL-competible filter.
//
// IMPORTANT (honesty): Saverisk does NOT expose lender credit ratings on this account (every
// rating endpoint returns 0 rows, even for NACL). So exact agency ratings are NOT scraped here.
// This module ships a CONSERVATIVE seed: the unambiguous AAA-tier majors (large PSU/private banks
// and the few AAA NBFCs) are hard-marked NOT competible (they fund cheaper than NACL, so NACL
// can't win on rate). Everything else is left UNKNOWN and treated as a *candidate* lead.
//
// To activate true tier filtering (AA- / A / BBB+ ...), Northern Arc should populate
// `lender_ratings.json` ({ "<lowercased lender name>": "<tier>" }) from its internal counterparty
// ratings — the same authoritative source used for the D/E gate. Tiers in TIER_ORDER below.
const fs = require('fs');
const path = require('path');

// Lower tier rank = stronger credit. NACL-competible = rank >= NACL's rank (similar/weaker credit).
const TIER_ORDER = ['AAA', 'AA+', 'AA', 'AA-', 'A+', 'A', 'A-', 'BBB+', 'BBB', 'BBB-', 'BB+', 'BB', 'BB-', 'B', 'UNKNOWN'];
const rankOf = (tier) => { const i = TIER_ORDER.indexOf(tier); return i < 0 ? TIER_ORDER.indexOf('UNKNOWN') : i; };

// Northern Arc's own benchmark tier (public long-term rating sits in the AA-/A+ band).
// Competible cutoff: incumbent lender rated at or weaker than this band -> NACL can plausibly match.
const NACL_TIER = 'AA-';
const COMPETIBLE_FROM_RANK = rankOf(NACL_TIER); // lenders with rank >= this are competible

const NACL_RE = /northern arc/i;
const TRUSTEE_RE = /catalyst|beacon|vistra|trustee|trusteeship|axis trustee|idbi trustee|mitcon|vardhman/i;

// Conservative seed: clearly AAA / very-high-grade lenders that fund cheaper than NACL.
// (Large banks + AAA NBFCs.) These are excluded from leads. NOT exhaustive — refine via JSON.
const SEED_AAA = [
  'state bank of india', 'hdfc bank limited', 'icici bank limited', 'axis bank limited',
  'kotak mahindra bank limited', 'bank of baroda', 'canara bank', 'union bank of india',
  'punjab national bank', 'indusind bank ltd.', 'idfc first bank limited', 'the federal bank ltd',
  'yes bank limited', 'bank of india', 'indian bank', 'bank of maharashtra', 'idbi bank limited',
  'uco bank', 'indian overseas bank', 'national housing bank', 'small industries development bank of india',
  'barclays bank plc.', 'the hongkong and shanghai banking corporation limited', 'sbm bank (india) limited',
  'standard chartered bank', 'bajaj finance limited', 'tata capital limited', 'aditya birla finance limited',
  'sundaram finance limited', 'mahindra and mahindra financial services limited', 'l&t finance limited',
  'kotak mahindra prime limited', 'hdb financial services limited',
];

function loadOverrides() {
  for (const p of [path.join(__dirname, 'lender_ratings.json'), path.join(__dirname, '..', 'lender_ratings.json')]) {
    try { return JSON.parse(fs.readFileSync(p, 'utf8')); } catch {}
  }
  return {};
}
const OVERRIDES = loadOverrides(); // { "<lower name>": "AA-" , ... }

function norm(name) { return String(name || '').trim().toLowerCase().replace(/\s+/g, ' '); }

// returns { tier, rank, source } for a lender name
function tierOf(name) {
  const n = norm(name);
  if (OVERRIDES[n]) return { tier: OVERRIDES[n], rank: rankOf(OVERRIDES[n]), source: 'override' };
  if (SEED_AAA.includes(n)) return { tier: 'AAA', rank: rankOf('AAA'), source: 'seed' };
  return { tier: 'UNKNOWN', rank: rankOf('UNKNOWN'), source: 'unknown' };
}

function isNACL(name) { return NACL_RE.test(name || ''); }
function isTrustee(name) { return TRUSTEE_RE.test(name || ''); }

// A lender is a NACL-competible incumbent if it is NOT NACL, NOT a debenture trustee, and either
// its known tier is at/weaker than NACL's band, OR its tier is UNKNOWN (candidate — review).
// (Known AAA seeds are excluded because rank(AAA) < COMPETIBLE_FROM_RANK.)
function isCompetible(name) {
  if (!name || isNACL(name) || isTrustee(name)) return false;
  const { tier, rank } = tierOf(name);
  if (tier === 'UNKNOWN') return true;          // candidate until rated
  return rank >= COMPETIBLE_FROM_RANK;          // AA- and weaker
}

module.exports = { TIER_ORDER, rankOf, NACL_TIER, tierOf, isNACL, isTrustee, isCompetible, norm, OVERRIDES_COUNT: Object.keys(OVERRIDES).length };

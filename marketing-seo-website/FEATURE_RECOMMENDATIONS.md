# SEO Automation Feature Recommendations

This document lists recommended next features for the Google Sheets SEO automation system.

The goal is to evolve the sheet from a data collector into an SEO action workspace that helps teams:
- identify the most important SEO issues
- generate practical recommendations
- prioritize fixes using real traffic and search data
- review and approve AI-assisted changes safely

## 1. AI SEO Prioritization

### Goal
Create a system that scores pages based on business importance and SEO opportunity using Search Console, GA4, and page audit data.

### Why It Is Needed
The sheet already collects a lot of useful data, but users still need to manually decide what to fix first. This feature would turn raw data into a ranked action list so teams can focus on the highest-impact pages.

## 2. AI Title Tag Recommendations

### Goal
Generate improved title tag suggestions for pages with missing, weak, duplicated, too-short, or too-long titles.

### Why It Is Needed
Meta descriptions are only one part of search snippet optimization. Titles strongly affect CTR and keyword targeting, so title recommendations are a natural next feature alongside AI meta descriptions.

## 3. Quick Wins Dashboard

### Goal
Create a summary view that highlights pages with the easiest and highest-value SEO improvements.

### Why It Is Needed
Users need a fast way to see what should be fixed right now. A dashboard can highlight issues like missing meta descriptions, low CTR pages with high impressions, and non-indexable pages that still matter.

## 4. Indexability Problem Assistant

### Goal
Translate Google index inspection results into plain-English explanations and recommended next actions.

### Why It Is Needed
The sheet already stores technical Google inspection data, but not every user will understand what those fields mean. This feature would make index problems easier to understand and act on.

## 5. AI CTR Improvement Suggestions

### Goal
Use Search Console impressions, clicks, CTR, and rankings to recommend better page titles and meta descriptions for underperforming pages.

### Why It Is Needed
Some pages already rank and get impressions but do not win clicks. This feature would help improve search snippet performance using real search visibility data instead of only static SEO checks.

## 6. Content Gap Suggestions

### Goal
Generate AI recommendations for missing content angles, weak search intent alignment, or unclear page positioning based on page metadata.

### Why It Is Needed
A page may be technically valid but still underperform because it does not match what users search for. This feature helps move beyond technical SEO into content quality and intent alignment.

## 7. Duplicate and Cannibalization Detection

### Goal
Detect pages with highly similar titles, H1s, or meta descriptions and flag likely keyword cannibalization issues.

### Why It Is Needed
Multiple pages competing for the same search intent can weaken SEO performance. This feature would help identify pages that need consolidation, differentiation, or stronger positioning.

## 8. Brand Voice and AI Guardrails

### Goal
Allow users to define tone, target audience, business summary, and restricted phrases that should guide all AI-generated outputs.

### Why It Is Needed
Without guardrails, AI outputs can become generic or off-brand. This feature would keep suggestions aligned with the business voice and reduce unwanted phrasing or hallucinated claims.

## 9. Prompt and AI Settings in Config

### Goal
Add configurable AI settings such as selected model, description length, tone preferences, CTA rules, and generation behavior.

### Why It Is Needed
Different businesses need different SEO writing styles and AI settings. Making these configurable reduces the need for code edits and makes the tool more reusable.

## 10. Human Review Workflow

### Goal
Add approval columns and review states for AI-generated suggestions, such as pending, approved, edited, rejected, and published.

### Why It Is Needed
SEO teams usually need review and approval before making updates live. This feature would make the sheet more useful in a real team workflow and safer for production use.

## 11. Internal Linking Suggestions

### Goal
Recommend related pages and possible anchor text opportunities based on page topics, titles, and URLs.

### Why It Is Needed
Internal linking is one of the most valuable SEO activities, but it is often handled manually. This feature would help users discover linking opportunities at scale.

## 12. Page-Type Rules

### Goal
Let users mark important pages such as homepage, service pages, landing pages, and blog posts so the automation can apply different SEO logic by page type.

### Why It Is Needed
Not all pages should be optimized the same way. Important commercial pages usually need safer and more intentional AI recommendations than lower-risk content pages.

## 13. Regenerate Only When Changed

### Goal
Only rerun AI generation for pages whose SEO-relevant inputs have changed since the last recommendation.

### Why It Is Needed
This reduces token usage, improves speed, and avoids unnecessary repeated AI calls for pages that have not changed.

## 14. Export for CMS or Team Handoff

### Goal
Provide a clean export-ready sheet of approved SEO updates that can be handed off to content, dev, or CMS teams.

### Why It Is Needed
Recommendations are more useful when they can easily move into implementation. This feature would make the sheet a stronger operations tool, not just an analysis tool.

## 15. SEO Health Summary Sheet

### Goal
Build an at-a-glance summary of overall SEO health for all tracked websites.

### Why It Is Needed
Users need a fast snapshot of system-wide health, such as total pages audited, pages missing metadata, pages with weak CTR, and AI recommendations still pending review.

## Recommended Build Order

If these are implemented in phases, this is the most practical order:

1. AI Title Tag Recommendations
2. Quick Wins Dashboard
3. AI SEO Prioritization
4. Indexability Problem Assistant
5. Brand Voice and AI Guardrails
6. Human Review Workflow
7. Internal Linking Suggestions

## Long-Term Product Direction

The strongest direction for this Google Sheet system is:

1. collect SEO and search performance signals
2. detect issues using code-based validation
3. use AI only where recommendations are actually needed
4. prioritize fixes using traffic and search data
5. support safe review and approval before implementation

That direction keeps the automation practical, data-driven, and useful for real SEO operations.

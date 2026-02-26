# Audit Presentation UI Guide

## Purpose

A monochrome, presentation-ready UI for translating squirrel audit outputs into clear business-facing reports. Designed for non-technical stakeholders and easy scanning in meetings.

## Output Location

Presentation HTML file is generated alongside the audit data file.

---

## Layout Structure

### 1. Hero Section

- Report title and context line (audience, generation date, scan snapshot date)
- Four summary metric cards showing key audit metrics

### 2. Executive Summary (Plain Language)

- What is working well
- What needs immediate action

### 3. Priority Findings

Categorized business impact areas:

- User Experience
- Technical Debt
- Performance Opportunities
- Governance & Compliance

### 4. Risk Assessment

- Lower scoring areas requiring attention
- Current strengths to maintain

### 5. Action Plan by Timeline

- Immediate (0-7 days)
- Near-Term (1-4 weeks)
- Ongoing maintenance

### 6. Priority Assets List

- Flagged items requiring optimization
- Expandable details with raw audit notes

---

## Visual System (Monochrome)

### Style Direction

Clean, corporate, presentation-first aesthetic using only grayscale tones. No color accents—all hierarchy through typography, spacing, and contrast.

### Palette

- **Background**: Subtle gray gradient (white to light gray)
- **Cards**: Pure white surfaces
- **Borders**: 1px light gray (`#e5e5e5`)
- **Text Primary**: Near-black (`#1a1a1a`)
- **Text Secondary**: Medium gray (`#6b7280`)
- **Dividers**: Light gray (`#e5e5e5`)

### Status Indicators (Monochrome)

Replace colored status pills with:

- **Critical**: Dark fill (`#1a1a1a`) with white text
- **Attention**: Outlined border with dark text
- **Good**: Light fill (`#f3f4f6`) with dark text

Or use icon patterns:

- Critical: Double border weight or filled indicator
- Attention: Single border, hollow
- Good: Lightest fill, subtle

### Typography

- **Base**: `"Avenir Next", "Segoe UI", Arial, sans-serif`
- **Hierarchy**: Clear distinction between H1 (report title), H2 (section headers), H3 (card titles), body text, and caption text
- **Scale**: Title (28-32px), Section (20-24px), Card (16-18px), Body (14px), Caption (12px)

### Surfaces & Elevation

- Cards: White background, 1px border, 4-8px border radius
- Shadows: None or very subtle (0 1px 3px rgba(0,0,0,0.08))
- Spacing: Consistent 24px section gaps, 16px card padding

---

## Design Tokens (CSS Variables)

```css
:root {
  --bg-page: #f5f5f5; /* Page background */
  --bg-hero: #fafafa; /* Hero section gradient end */
  --card: #ffffff; /* Card surface */
  --text-primary: #1a1a1a; /* Headings, primary content */
  --text-secondary: #6b7280; /* Metadata, captions */
  --text-muted: #9ca3af; /* Disabled, placeholders */
  --border: #e5e5e5; /* Card borders, dividers */
  --border-strong: #a3a3a3; /* Emphasis borders */
  --fill-critical: #1a1a1a; /* Critical status background */
  --fill-attention: #ffffff; /* Attention status background */
  --fill-good: #f3f4f6; /* Good status background */
}
```

---

## Responsiveness

- **Max container width**: 1100px
- **Grid patterns**: `repeat(auto-fit, minmax(260px, 1fr))` for cards
- **Breakpoints**: Fluid scaling from desktop (1100px) → tablet → mobile (320px)
- **Mobile adjustments**: Stack grid columns, reduce padding to 12px, hide secondary metadata

---

## Print Behavior

```css
@media print {
  body {
    background: white;
  }
  .card {
    box-shadow: none;
    border: 1px solid #ddd;
  }
  a {
    text-decoration: underline;
    color: #000;
  }
}
```

- Remove decorative shadows and gradients
- Ensure borders print clearly
- Preserve link underlines for PDF readability
- Force background colors to print (if supported)

---

## Implementation Notes

1. **Data Source**: Consume squirrel JSON output directly
2. **Scoring Display**: Show numeric scores with visual bar indicators (grayscale fill levels)
3. **Expandable Sections**: Use `<details>` elements or toggle buttons for raw audit notes
4. **Asset Links**: Direct clickable links to flagged files/resources
5. **No External Dependencies**: Pure HTML/CSS/JS, no frameworks required

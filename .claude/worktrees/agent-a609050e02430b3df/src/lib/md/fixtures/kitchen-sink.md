---
title: Kitchen Sink
tags: [markdown, spec-050]
---

# Heading One

## Heading Two

A plain paragraph, blank-line separated, with **bold**, _italic_, ~~strikethrough~~ and `inline code`.

A paragraph with a hard\
line break.

A paragraph with a [hyperlink](https://example.com/page).

```js
const x = 1;
console.log(x);
```

> A blockquote spanning\
> two physical lines.

---

- bullet one
- bullet two
  - nested bullet

A separating paragraph, so the numbered list below starts its own list run.

1. number one
2. number two

| Column A | Column B |
| --- | --- |
| cell 1 | cell 2 |
| cell 3 | cell 4 |

![a red dot](data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=)

<div class="raw-html-block">
  <p>This block is outside the supported subset and must round-trip verbatim.</p>
</div>

[ref-link]: https://example.com/reference "A reference-style link definition"

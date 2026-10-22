---
id: SPEC-021
title: Keyboard shortcuts and shortcut reference
status: draft
area: shell
depends_on: [SPEC-012, SPEC-015, SPEC-017]
---
## Why
Shortcuts are how fast users work, and by this point they are scattered across
six specs with no single owner. This spec consolidates them into one registry
so later features cannot silently collide.

## Acceptance criteria
- [ ] All shortcuts are declared in a single registry mapping command id → key binding per platform, and the app binds keys only from that registry
- [ ] Bindings follow platform convention: Cmd on macOS, Ctrl on Windows/Linux, resolved at runtime rather than by user-agent sniffing
- [ ] A build-time or test-time check fails if two commands claim the same binding on the same platform
- [ ] Ctrl/Cmd+/ opens a searchable "Keyboard shortcuts" dialog listing every registered shortcut grouped by category, generated from the registry
- [ ] The dialog is keyboard-navigable and closes with Escape
- [ ] Browser shortcuts we deliberately override (Ctrl/Cmd+B, +I, +U, +F, +P) are documented in the dialog; ones we must not override (Ctrl/Cmd+T, +W, +N) are never bound
- [ ] Shortcuts do not fire while focus is in a dialog, the find bar, or any other text input outside the document
- [ ] Every toolbar control with a shortcut shows it in the tooltip, read from the registry

## Out of scope
User-remappable shortcuts. A command palette. Word/LibreOffice compatibility
shortcut modes.

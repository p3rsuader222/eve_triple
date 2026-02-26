# AGENTS.md

## PowerShell Anti-Stall Rules
- Prefer `rg -n` for file discovery and contextual reads (`-A/-B`) over ranged `Get-Content` one-liners.
- Avoid commands like `$c=Get-Content ...; $c[start..end]`.
- Keep commands short and single-purpose; avoid long quoted command strings.
- Avoid chaining commands with `;`, `|`, `&&` unless necessary.
- Use `apply_patch` for file edits instead of shell text replacement.
- If a read command hits prompt friction, fallback immediately to `rg -n` context reads.

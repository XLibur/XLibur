# CLAUDE.md

## Shell Commands

- Do not use compound commands (e.g., `&&`, `||`, `;`) in Bash tool calls. Run each command as a separate Bash tool invocation.
- Never use compound commands with bash or git. Each command must be its own separate Bash tool call.
- Never use `cd <folder> && git <params>` style commands. Use absolute paths or set the working directory separately.

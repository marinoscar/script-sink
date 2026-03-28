# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

**script-sink** is a collection of cross-platform utility and productivity scripts for automation. Scripts are designed to run as:
- Linux cron jobs (Shell/Bash)
- Windows Scheduled Tasks (PowerShell)
- Node.js/TypeScript tasks
- Python tasks
- OpenClaw agent tasks

## Supported Languages

- **PowerShell** (`.ps1`) — Windows automation
- **Shell/Bash** (`.sh`) — Linux automation
- **Node.js / TypeScript** — cross-platform tasks
- **Python** — cross-platform tasks

## Repository Structure

All scripts live under `scripts/`, each in its own folder:

```
scripts/
└── <script-name>/
    ├── README.md       # Purpose, prerequisites, how to run, config reference
    ├── config.json     # Optional configuration (script uses defaults if missing)
    ├── output/         # Generated output files (gitignored)
    ├── logs/           # Log files (gitignored)
    └── <script files>
```

When adding a new script, create a new folder under `scripts/` following this pattern.

## Conventions

- Scripts should be self-contained or clearly document their dependencies.
- Include a brief comment header describing purpose, usage, and intended runtime (cron, scheduled task, agent).
- Each script folder must include a README.md with prerequisites, usage instructions, and configuration reference.
- Check for existing scripts before creating new ones to avoid duplication.

## Versioning

Every script must declare a version variable at the top and print it on startup:

- **PowerShell**: `$scriptVersion = "1.0.0"` followed by `Write-Host "ScriptName v$scriptVersion" -ForegroundColor Cyan`
- **Bash/Shell**: `SCRIPT_VERSION="1.0.0"` followed by `echo "ScriptName v${SCRIPT_VERSION}"`
- **Node.js/TypeScript**: `const SCRIPT_VERSION = "1.0.0";` followed by `console.log(\`ScriptName v${SCRIPT_VERSION}\`);`
- **Python**: `SCRIPT_VERSION = "1.0.0"` followed by `print(f"ScriptName v{SCRIPT_VERSION}")`

The version should also be included in log output where applicable.

**When modifying a script, always increment the version** using semver (MAJOR.MINOR.PATCH):
- PATCH: bug fixes, minor tweaks
- MINOR: new features, behavioral changes
- MAJOR: breaking changes, major rewrites

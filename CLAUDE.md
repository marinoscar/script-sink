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

## Conventions

- Scripts should be self-contained or clearly document their dependencies.
- Include a brief comment header describing purpose, usage, and intended runtime (cron, scheduled task, agent).
- Check for existing scripts before creating new ones to avoid duplication.

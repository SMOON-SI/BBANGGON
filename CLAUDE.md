# CLAUDE.md

This file provides guidance to AI assistants (Claude and others) working in this repository.

## Project Overview

**Repository**: BBANGGON (`SMOON-SI/BBANGGON`)
**Status**: Newly initialized — project structure is not yet established.

This repository was initialized on 2026-02-19 and is currently empty. As the project grows, update this file to reflect the actual tech stack, conventions, and workflows.

---

## Repository Structure

> Update this section as the project structure is established.

```
BBANGGON/
├── CLAUDE.md         # This file
└── .gitkeep          # Placeholder (remove once real files are added)
```

---

## Development Workflow

### Branching Strategy

- **`master`**: Stable, production-ready code
- **`claude/...`**: AI-assisted development branches (auto-created per session)
- Feature branches should follow a consistent naming convention once established

### Git Conventions

- Write clear, descriptive commit messages in the imperative mood (e.g., "Add user authentication", not "Added user authentication")
- Keep commits focused — one logical change per commit
- Always push to the correct branch; never push to `master` without review

### Making Changes

1. Ensure you are on the correct feature branch before editing files
2. Read existing files thoroughly before modifying them
3. Prefer editing existing files over creating new ones unless truly necessary
4. Run linters and tests before committing (commands TBD once tooling is configured)

---

## Tech Stack

> To be determined — update this section once the project technology choices are made.

- **Language**: TBD
- **Framework**: TBD
- **Package Manager**: TBD
- **Testing Framework**: TBD
- **Database**: TBD

---

## Code Conventions

> Update this section based on actual project conventions once established.

- Follow the style guide enforced by the project's linter/formatter
- Avoid over-engineering — implement only what is needed for the current task
- Do not add unnecessary comments; let code be self-explanatory where possible
- Do not introduce new dependencies without discussion

---

## Testing

> Update this section once a testing framework is configured.

- Write tests for all new functionality
- Ensure existing tests pass before pushing
- Run the full test suite with: `[command TBD]`

---

## Environment Setup

> Update this section once environment configuration is defined.

1. Clone the repository
2. Install dependencies: `[command TBD]`
3. Configure environment variables: copy `.env.example` to `.env` and fill in values
4. Start the development server: `[command TBD]`

---

## Key Notes for AI Assistants

- **This repository is at a very early stage.** There is no existing code to reference.
- When the user begins building the project, update this CLAUDE.md to reflect the actual structure, tech stack, and conventions.
- Always confirm the active branch before committing: `git branch --show-current`
- Use the development branch `claude/claude-md-mm25r21nrqdcuijn-esUJ0` for changes unless told otherwise
- Never push to `master` without explicit user permission
- Keep changes minimal and focused on what was explicitly requested
- Ask clarifying questions rather than making assumptions about project direction

---

## Maintenance

This file should be updated whenever:
- The tech stack is decided or changes
- New tooling (linters, formatters, CI/CD) is added
- Project structure significantly changes
- New conventions or workflows are adopted

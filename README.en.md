<div align="center">

# patent-builder

### A Codex skill for drafting patent disclosure documents

[中文](./README.md) | **English**

</div>

`patent-builder` is not a standalone app.
It is a skill that an agent can install and use to turn raw technical ideas into structured patent disclosure drafts, with final output in `Markdown` or `DOCX`.

## What This Skill Does

- Drafts patent disclosure documents through guided conversation
- Supports final output as `markdown` or `docx`
- Includes a default disclosure template
- Exports existing Markdown into `.docx`
- Works for methods, systems, devices, structures, materials, processes, controls, and algorithms

## Installation

If you use `Cursor`, `Codex`, `OpenClaw`, or any other tool that supports skills or agent workflows, the easiest approach is to give the agent either the **local project path** or the **repository URL** and let it install the skill for you.

### Option 1: Install from a local path

Send the project path to your agent:

```text
Please install this skill from the local project path:
/path/to/patent-builder
```

If you are on this machine, you can also use the real path directly:

```text
Please install this skill from:
/Users/yangchaoqun/Desktop/专利/patent-builder
```

### Option 2: Install from a repository URL

If the project is hosted in Git, send the repo URL to your agent:

```text
Please install the skill from this repository URL:
<your-repo-url>
```

### Recommended prompt

```text
Install the skill from this project and make it available in my skills list.
Use the skill name defined in the repo and finish the setup automatically.
```

## How To Use It

Once installed, you can ask your agent to use the skill directly. A typical flow is:

1. Tell the agent whether you want `markdown` or `docx`
2. Describe the invention, goals, key objects, and constraints
3. Ask it to draft the disclosure with this skill
4. If needed, ask it to export the result as `.docx`

Example:

```text
Use patent-builder to draft a patent disclosure for my idea.
Final output format: docx.
```

## Repository Contents

- [patent-disclosure-from-docx/SKILL.md](/Users/yangchaoqun/Desktop/专利/patent-builder/patent-disclosure-from-docx/SKILL.md): skill definition
- [patent-disclosure-from-docx/agents/openai.yaml](/Users/yangchaoqun/Desktop/专利/patent-builder/patent-disclosure-from-docx/agents/openai.yaml): agent configuration
- [patent-disclosure-from-docx/references/default-disclosure-template.md](/Users/yangchaoqun/Desktop/专利/patent-builder/patent-disclosure-from-docx/references/default-disclosure-template.md): default disclosure template
- [patent-disclosure-from-docx/scripts/markdown_to_docx.py](/Users/yangchaoqun/Desktop/专利/patent-builder/patent-disclosure-from-docx/scripts/markdown_to_docx.py): Markdown-to-DOCX exporter

## Requirements

- Python 3
- no third-party dependencies

## In One Line

If you want a skill that agents can install and use directly to produce patent disclosure drafts, this repository is built for exactly that.

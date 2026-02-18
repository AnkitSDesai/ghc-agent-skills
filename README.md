# GitHub Copilot Agent Skills

A growing collection of GitHub Copilot agents and reusable skills. This repository is continuously updated with new agents and domain-specific skills designed to extend Copilot's capabilities across a range of engineering and architecture workflows.

## Structure

```
.github/
├── agents/                          # Copilot agent definitions
│   ├── azure-architect.agent.md     # Azure Solutions Architect agent
│   └── gh-aw-builder.agent.md       # GitHub Agentic Workflow Builder agent
├── skills/                          # Reusable skill instructions
│   ├── api-security-review/         # API security audits and OWASP compliance
│   ├── apim-policy-authoring/       # Azure API Management policy generation
│   ├── apiops-deployment/           # APIOps CI/CD and IaC deployment
│   ├── architecture-design/         # Azure solution design and HLD documentation
│   ├── azure-apim-architecture/     # APIM architecture decisions and patterns
│   ├── cost-optimization/           # Azure cost analysis and savings recommendations
│   ├── drawio-mcp-diagramming/      # Architecture diagrams via Draw.io MCP
│   ├── gh-aw-operations/            # GitHub Agentic Workflow creation, compilation, and debugging
│   └── waf-assessment/              # Well-Architected Framework assessments
└── copilot-instructions.md          # Repo-wide Copilot governance and defaults
```

## Agents

| Agent | Description |
|-------|-------------|
| `azure-architect` | Azure Solutions Architect — designs production-ready architectures aligned to WAF and CAF |
| `gh-aw-builder` | GitHub Agentic Workflow Builder — creates and configures markdown-based AI-powered GitHub Agentic Workflows (gh-aw) with proper frontmatter, MCP servers, safe-outputs, and best practices |

## Skills

| Skill | Description |
|-------|-------------|
| `api-security-review` | Reviews API Management configurations against OWASP API Security Top 10 and Azure Security Benchmark |
| `apim-policy-authoring` | Generates production-ready APIM policy XML for auth, rate limiting, CORS, and transformations |
| `apiops-deployment` | Guides APIM deployments using Bicep/Terraform and CI/CD pipelines |
| `architecture-design` | Designs Azure architectures from requirements with service selection, cost estimates, and WAF alignment |
| `azure-apim-architecture` | Analyses APIM architecture decisions including VNet topology, multi-environment strategies, and component trade-offs |
| `cost-optimization` | Identifies cost reduction opportunities and quantifies savings across Azure workloads |
| `drawio-mcp-diagramming` | Creates and edits architecture diagrams using the Draw.io MCP integration |
| `gh-aw-operations` | Comprehensive skills for creating, compiling, debugging, and managing GitHub Agentic Workflows (gh-aw) |
| `waf-assessment` | Assesses architectures across all five WAF pillars and provides scored recommendations |

## Usage

1. Clone or fork this repository.
2. Open in VS Code with GitHub Copilot Chat enabled.
3. Reference an agent via `@azure-architect` or invoke a skill by asking Copilot to use it by name.
4. Skills are loaded automatically when referenced in `copilot-instructions.md` or triggered by relevant queries.

## Contributing

New agents and skills are added on an ongoing basis. Each skill lives in its own directory under `.github/skills/` and follows a consistent `SKILL.md` structure defining its purpose, inputs, and output format.

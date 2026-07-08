export interface LlmsFullSource {
  readonly title: string
  readonly relativePath: string
  readonly url: string
}

const llmsFullSourceSpecs = [
  ['Repository README', 'README.md'],
  ['Agent Install Context', 'llms-install.md'],
  ['Workbook Compatibility Report', 'docs/workbook-compatibility-report.md'],
  ['Workbook Compatibility Report Transcript', 'docs/workbook-compatibility-report-transcript.md'],
  ['Agent XLSX Risk Preflight', 'docs/agent-xlsx-risk-preflight.md'],
  ['Evaluate XLSX Formula Recalculation', 'docs/eval-xlsx-recalc.md'],
  ['External Workbook Recalculation Proof', 'docs/external-workbook-recalc-proof.md'],
  ['Evaluate WorkPaper In A Node Service', 'docs/eval-workpaper-service.md'],
  ['WorkPaper Host Handoff', 'docs/agent-adoption-kit.md'],
  ['Host Rule Chooser', 'docs/agent-rule-chooser.md'],
  ['Evaluate Bilig As An MCP Workbook Tool', 'docs/eval-agent-mcp.md'],
  ['WorkPaper Evaluator Matrix', 'docs/agent-proof-matrix.md'],
  ['WorkPaper Package README', 'packages/workpaper/README.md'],
  ['WorkPaper Host Handbook', 'docs/headless-workpaper-agent-handbook.md'],
  ['Agent Workbook Challenge', 'docs/agent-workbook-challenge.md'],
  ['WorkPaper Tool-Calling Recipe', 'docs/agent-workpaper-tool-calling-recipe.md'],
  ['MCP Spreadsheet Formula Server For Tool Hosts', 'docs/mcp-spreadsheet-formula-server-for-coding-agents.md'],
  ['Spreadsheet MCP Server Comparison', 'docs/spreadsheet-mcp-server-comparison.md'],
  ['Vercel AI SDK Spreadsheet Tool Formula Readback', 'docs/vercel-ai-sdk-spreadsheet-tool-formula-readback.md'],
  ['WorkPaper Tool For Node.js', 'docs/ai-agent-spreadsheet-tool-node.md'],
  ['Workbook Tools For MCP, Services, And Framework Integrations', 'docs/agent-framework-workbook-tools.md'],
  ['Browser Use WorkPaper Formula Tool', 'docs/browser-use-workpaper-formula-tool.md'],
  ['Workbook Runtime Intent API', 'docs/workbook-runtime-intent-api.md'],
  ['Workbook Runtime Intent API', 'docs/workbook-agent-intent-api.md'],
  ['Workbook Package README', 'packages/workbook/README.md'],
  ['Cloudflare Agents WorkPaper Spreadsheet Tool', 'docs/cloudflare-agents-workpaper-spreadsheet-tool.md'],
  ['CrewAI WorkPaper Spreadsheet Tool', 'docs/crewai-workpaper-spreadsheet-tool.md'],
  ['LlamaIndex.TS WorkPaper Spreadsheet Tool', 'docs/llamaindex-workpaper-spreadsheet-tool.md'],
  ['Gemini CLI WorkPaper Extension', 'docs/gemini-cli-workpaper-extension.md'],
  ['Open WebUI WorkPaper Tool Setup', 'docs/open-webui-workpaper-mcp.md'],
  ['Open Multi-Agent WorkPaper MCP Example', 'docs/open-multi-agent-workpaper-mcp.md'],
  ['LobeHub WorkPaper MCP Setup', 'docs/lobehub-workpaper-mcp.md'],
  ['AnythingLLM WorkPaper MCP Setup', 'docs/anythingllm-workpaper-mcp.md'],
  ['OpenHands WorkPaper MCP Setup', 'docs/openhands-workpaper-mcp.md'],
  ['Trae WorkPaper MCP Setup', 'docs/trae-workpaper-mcp.md'],
  ['Qodo WorkPaper MCP Setup', 'docs/qodo-workpaper-mcp.md'],
  ['OpenCode WorkPaper MCP Setup', 'docs/opencode-workpaper-mcp.md'],
  ['Aider WorkPaper Conventions', 'docs/aider-workpaper-conventions.md'],
  ['ChatGPT Apps WorkPaper MCP', 'docs/chatgpt-apps-workpaper-mcp.md'],
  ['Sim WorkPaper MCP Setup', 'docs/sim-workpaper-mcp.md'],
  ['n8n WorkPaper Formula Readback', 'docs/n8n-workpaper-formula-readback.md'],
  ['OpenAI Agents SDK WorkPaper Tool', 'docs/openai-agents-sdk-workpaper-tool.md'],
  ['MCP WorkPaper Tool Server', 'docs/mcp-workpaper-tool-server.md'],
  ['Agent XLSX Formula Recalculation Without LibreOffice', 'docs/agent-xlsx-formula-recalculation-without-libreoffice.md'],
  ['ExcelJS Formula Result Not Updating After Node Edits', 'docs/exceljs-formula-result-not-updating-after-node-edits.md'],
  ['Formula Bug Clinic', 'docs/formula-bug-clinic.md'],
  ['Google Sheets QUERY and SORTN in Node.js', 'docs/google-sheets-query-sortn-node-workpaper.md'],
  ['Try Bilig WorkPaper In Node', 'docs/try-bilig-headless-in-node.md'],
  ['Quote Approval WorkPaper API', 'docs/quote-approval-workpaper-api.md'],
  ['Compatibility Limits', 'docs/where-bilig-is-not-excel-compatible-yet.md'],
  ['npm Provenance And Package Trust', 'docs/npm-provenance-package-trust.md'],
] as const

export function buildLlmsFullSources(repositoryUrl: string): readonly LlmsFullSource[] {
  return llmsFullSourceSpecs.map(([title, relativePath]) => ({
    title,
    relativePath,
    url: `${repositoryUrl}/blob/main/${relativePath}`,
  }))
}

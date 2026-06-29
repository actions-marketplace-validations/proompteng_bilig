import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export type CliInspectLimit = number | 'all'

interface GithubActionWorkflowOptions {
  readonly workbooks: string
  readonly changedFilesOnly: boolean
  readonly failOnStale: boolean
  readonly inspectLimit: CliInspectLimit
  readonly jsonOutput: string
  readonly markdownOutput: string
  readonly packageVersion: string
  readonly workflowName: string
}

const printGithubActionOption = '--print-github-action'
const defaultGithubActionPackageVersion = readPackageVersion()

export function parseGithubActionWorkflowArgs(
  args: readonly string[],
  commandName: string,
  helpers: {
    readonly defaultInspectFormulaLimit: CliInspectLimit
    readonly parseInspectLimit: (raw: string) => CliInspectLimit
    readonly requireNextArg: (args: readonly string[], index: number, option: string) => string
  },
): GithubActionWorkflowOptions {
  let workbooks: string | undefined
  let changedFilesOnly = true
  let failOnStale = false
  let inspectLimit = helpers.defaultInspectFormulaLimit
  let jsonOutput = '${{ runner.temp }}/xlsx-cache-doctor.json'
  let markdownOutput = '${{ runner.temp }}/xlsx-cache-doctor.md'
  let packageVersion = defaultGithubActionPackageVersion
  let workflowName = 'xlsx-cache-doctor'

  for (let index = 0; index < args.length; index += 1) {
    const arg = args[index]
    if (arg === undefined) {
      throw new Error(`Unexpected missing ${commandName} argument`)
    }
    switch (arg) {
      case printGithubActionOption:
        break
      case '--workbook':
      case '--workbooks':
        workbooks = helpers.requireNextArg(args, index, arg)
        index += 1
        break
      case '--changed-files-only':
        changedFilesOnly = parseBooleanOption(helpers.requireNextArg(args, index, '--changed-files-only'), '--changed-files-only')
        index += 1
        break
      case '--fail-on-stale':
        failOnStale = parseBooleanOption(helpers.requireNextArg(args, index, '--fail-on-stale'), '--fail-on-stale')
        index += 1
        break
      case '--inspect-limit':
        inspectLimit = helpers.parseInspectLimit(helpers.requireNextArg(args, index, '--inspect-limit'))
        index += 1
        break
      case '--json-output':
        jsonOutput = helpers.requireNextArg(args, index, '--json-output')
        index += 1
        break
      case '--markdown-output':
        markdownOutput = helpers.requireNextArg(args, index, '--markdown-output')
        index += 1
        break
      case '--package-version':
        packageVersion = helpers.requireNextArg(args, index, '--package-version')
        index += 1
        break
      case '--workflow-name':
        workflowName = helpers.requireNextArg(args, index, '--workflow-name')
        index += 1
        break
      default:
        if (arg.startsWith('-')) {
          throw new Error(`Unknown ${commandName} option for ${printGithubActionOption}: ${arg}`)
        }
        if (workbooks !== undefined) {
          throw new Error(`Unexpected extra workbook glob for ${printGithubActionOption}: ${arg}`)
        }
        workbooks = arg
    }
  }

  if (!workbooks) {
    throw new Error(`Expected workbook path or glob after ${printGithubActionOption}`)
  }

  return {
    workbooks,
    changedFilesOnly,
    failOnStale,
    inspectLimit,
    jsonOutput,
    markdownOutput,
    packageVersion,
    workflowName,
  }
}

export function printGithubActionWorkflow(options: GithubActionWorkflowOptions, writeStdout: (text: string) => void): void {
  writeStdout(
    [
      `name: ${yamlDoubleQuote(options.workflowName)}`,
      '',
      'on:',
      '  pull_request:',
      '    paths:',
      '      - "**/*.xlsx"',
      '  workflow_dispatch:',
      '',
      'permissions:',
      '  contents: read',
      '',
      'jobs:',
      '  inspect-xlsx-formula-caches:',
      '    runs-on: ubuntu-latest',
      '    steps:',
      '      - uses: actions/checkout@v5',
      '        with:',
      '          fetch-depth: 0',
      '',
      '      - uses: actions/setup-node@v6',
      '        with:',
      '          node-version: "22"',
      '          package-manager-cache: false',
      '',
      '      - id: cache-doctor',
      '        uses: proompteng/bilig@v1',
      '        with:',
      `          workbooks: ${yamlDoubleQuote(options.workbooks)}`,
      `          changed-files-only: ${yamlDoubleQuote(String(options.changedFilesOnly))}`,
      `          package-version: ${yamlDoubleQuote(options.packageVersion)}`,
      `          inspect-limit: ${yamlDoubleQuote(String(options.inspectLimit))}`,
      `          json-output: ${yamlDoubleQuote(options.jsonOutput)}`,
      `          markdown-output: ${yamlDoubleQuote(options.markdownOutput)}`,
      `          fail-on-stale: ${yamlDoubleQuote(String(options.failOnStale))}`,
      '',
      '      - uses: actions/upload-artifact@v4',
      '        if: always()',
      '        with:',
      '          name: xlsx-cache-doctor-report',
      '          path: |',
      '            ${{ steps.cache-doctor.outputs.json }}',
      '            ${{ steps.cache-doctor.outputs.markdown }}',
      '',
    ].join('\n'),
  )
}

export function defaultGithubActionPackageVersionForHelp(): string {
  return defaultGithubActionPackageVersion
}

function readPackageVersion(): string {
  const packageJsonPath = join(dirname(fileURLToPath(import.meta.url)), '..', 'package.json')
  const parsed: unknown = JSON.parse(readFileSync(packageJsonPath, 'utf8'))
  if (typeof parsed !== 'object' || parsed === null || Array.isArray(parsed)) {
    throw new Error(`Expected package.json object at ${packageJsonPath}`)
  }
  const version = (parsed as { readonly version?: unknown }).version
  if (typeof version !== 'string' || version.length === 0) {
    throw new Error(`Expected package.json version at ${packageJsonPath}`)
  }
  return version
}

function parseBooleanOption(raw: string, option: string): boolean {
  if (raw === 'true') {
    return true
  }
  if (raw === 'false') {
    return false
  }
  throw new Error(`Expected ${option} to be "true" or "false", received: ${raw}`)
}

function yamlDoubleQuote(value: string): string {
  return JSON.stringify(value)
}

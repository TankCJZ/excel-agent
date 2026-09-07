import { z } from 'zod'

// Canonical meanings for localization; these are not tied to the website locale.
export const followupTemplates = {
  extremes: 'Show the highest and lowest “{dimension}” by “{metric}”',
  rank_details: 'List the full “{dimension}” ranking with “{metric}” details',
  create_chart: 'Summarize “{metric}” by “{dimension}” and create a chart',
  count_by_dimension: 'Count records by “{dimension}” and create a ranking',
  check_quality: 'Check for blanks, duplicates, and unusual records',
  export_chart: 'Export this chart to a new Excel file',
  export_analysis: 'Export this analysis to a new Excel file',
  summarize_findings: 'Summarize the 3 most important findings in this workbook',
  review_columns: 'Explain what each column in this spreadsheet is for',
  check_sources: 'List the data sources used for this spreadsheet',
  compare_findings: 'Compare these findings and highlight the most important differences',
  create_research_sheet: 'Turn these research findings into a downloadable Excel spreadsheet',
  rank_sources: 'Organize the sources used here by reliability'
} as const

export const followupSuggestionSchema = z.object({
  id: z.string().min(1).max(120),
  intent: z.enum(['analyze', 'chart', 'export', 'quality']),
  kind: z.enum(Object.keys(followupTemplates) as [keyof typeof followupTemplates, ...Array<keyof typeof followupTemplates>]),
  params: z.record(z.string(), z.string().max(1024)),
  prompt: z.string().trim().min(1).max(3000).optional()
})
export type FollowupSuggestion = z.infer<typeof followupSuggestionSchema>
export type FollowupIntent = FollowupSuggestion['intent']
export type FollowupKind = FollowupSuggestion['kind']

export const followupPresentationSchema = z.object({
  version: z.literal(1),
  language: z.string().regex(/^[a-zA-Z]{2,8}(?:-[a-zA-Z0-9]{1,8})*$/).max(35),
  title: z.string().trim().min(1).max(160),
  suggestions: z.array(followupSuggestionSchema.extend({ prompt: z.string().trim().min(1).max(3000) })).min(1).max(3)
})
export type FollowupPresentation = z.infer<typeof followupPresentationSchema>

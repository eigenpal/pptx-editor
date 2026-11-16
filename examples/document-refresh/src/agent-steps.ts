// The steps of the scripted agent on the replay page, shared by the mock server and the page.

/** One agent step: what the agent says it does. */
export interface AgentStep {
  readonly title: string;
  readonly detail: string;
}

export const AGENT_STEPS: readonly AgentStep[] = [
  {
    title: 'Update the revenue figure',
    detail: 'Reads the finance sheet and changes $4.0M to $4.6M on the revenue slide.',
  },
  {
    title: 'Make the growth number stand out',
    detail: 'Recolors the growth box green and updates it to 15%.',
  },
  {
    title: 'Retitle the deck',
    detail: 'Renames the title slide to "Q3 business review".',
  },
  {
    title: 'Add a risks slide',
    detail: 'Inserts a new slide about supplier delays after the roadmap.',
  },
  {
    title: 'Write speaker notes',
    detail: 'Adds talking points to the notes of the roadmap slide.',
  },
  {
    title: 'Remove the draft label',
    detail: 'Deletes the yellow DRAFT badge from the title slide.',
  },
  {
    title: 'Reorder the slides',
    detail: 'Moves the team slide before the roadmap.',
  },
];

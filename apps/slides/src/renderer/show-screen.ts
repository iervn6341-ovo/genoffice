/**
 * Choosing the screen for the show before it starts. With presenter view the audience sees the slides
 * on one of the screens other than the presenter's; whenever there is one the user is asked which,
 * every time, instead of the app guessing.
 */
export interface ShowScreen {
  id: number
  label: string
  primary: boolean
  /** the screen presenter view opens on */
  current: boolean
  width: number
  height: number
}

/** Screens the show can go to: every one but the presenter's. */
export const showCandidates = (screens: readonly ShowScreen[]): ShowScreen[] =>
  screens.filter((s) => !s.current)

/**
 * Ask on every start whenever another screen is connected (two or more screens in all): the app
 * never guesses where the audience should look.
 */
export const shouldAskScreen = (screens: readonly ShowScreen[]): boolean =>
  showCandidates(screens).length >= 1

/** The entry the dialog starts on: last time's screen when it is still there, else the first. */
export function preselectedScreen(
  candidates: readonly ShowScreen[],
  remembered: string | null,
): ShowScreen | undefined {
  return candidates.find((s) => s.label === remembered) ?? candidates[0]
}

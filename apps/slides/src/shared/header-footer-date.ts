/** Cached PowerPoint date field text for the format selector and slide preview. */
export function formatHeaderFooterDate(date: Date, field: string, locale: string): string | null {
  const kind = Number(/^datetime([1-9]|1[0-3])$/.exec(field)?.[1])
  if (!kind) return null
  const dateParts = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: 'long',
    day: '2-digit',
    weekday: 'long',
  }).formatToParts(date)
  const shortParts = new Intl.DateTimeFormat(locale, {
    year: '2-digit',
    month: 'short',
    day: '2-digit',
  }).formatToParts(date)
  const numericParts = new Intl.DateTimeFormat(locale, {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
  }).formatToParts(date)
  const pick = (parts: Intl.DateTimeFormatPart[], type: Intl.DateTimeFormatPartTypes) =>
    parts.find((part) => part.type === type)?.value ?? ''
  const year = pick(dateParts, 'year')
  const month = pick(dateParts, 'month')
  const day = pick(dateParts, 'day')
  const weekday = pick(dateParts, 'weekday')
  const shortMonth = pick(shortParts, 'month')
  const shortYear = pick(shortParts, 'year')
  const mm = pick(numericParts, 'month')
  const dd = pick(numericParts, 'day')
  const numericDate = `${mm}/${dd}/${year}`
  const time = (hour12: boolean, seconds: boolean) => {
    const parts = new Intl.DateTimeFormat(locale, {
      hour: hour12 ? 'numeric' : '2-digit',
      minute: '2-digit',
      ...(seconds ? { second: '2-digit' } : {}),
      hour12,
    }).formatToParts(date)
    const hh = pick(parts, 'hour')
    const min = pick(parts, 'minute')
    const sec = pick(parts, 'second')
    const period = pick(parts, 'dayPeriod')
    return `${hh}:${min}${seconds ? `:${sec}` : ''}${hour12 && period ? ` ${period}` : ''}`
  }
  switch (kind) {
    case 1:
      return numericDate
    case 2:
      return `${weekday}, ${month} ${day}, ${year}`
    case 3:
      return `${day} ${month} ${year}`
    case 4:
      return `${month} ${day}, ${year}`
    case 5:
      return `${day}-${shortMonth}-${shortYear}`
    case 6:
      return `${month} ${shortYear}`
    case 7:
      return `${shortMonth}-${shortYear}`
    case 8:
      return `${numericDate} ${time(true, false)}`
    case 9:
      return `${numericDate} ${time(true, true)}`
    case 10:
      return time(false, false)
    case 11:
      return time(false, true)
    case 12:
      return time(true, false)
    case 13:
      return time(true, true)
    default:
      return null
  }
}

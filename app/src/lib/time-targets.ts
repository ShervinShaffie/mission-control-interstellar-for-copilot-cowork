/**
 * Suggested starting targets for the user's own Outlook categories. Free and local: based only on
 * category names and this week's calendar. The user can change any value; once they do, theirs are saved.
 */
const RULES: { test: RegExp; weight: number }[] = [
  { test: /focus|deep|heads?.?down|think|maker|writing|build/i, weight: 25 },
  { test: /meeting|collab|sync|stand.?up|team/i, weight: 30 },
  { test: /customer|client|partner|sales|external|account/i, weight: 20 },
  { test: /project|plan|strategy|review|roadmap|design/i, weight: 15 },
  { test: /1:1|one.?on.?one|coach|mentor|manage|people|hiring|interview/i, weight: 10 },
  { test: /admin|email|inbox|ops|operation|process|expense|travel/i, weight: 10 },
  { test: /learn|training|develop|skill|course|read/i, weight: 5 },
  { test: /personal|break|lunch|wellbeing|well-being|health|family|exercise/i, weight: 5 },
];
/** Outlook's built-in color names ("Blue category") say nothing about the work, so they're skipped unless used. */
const COLOR_ONLY = /^(red|orange|peach|yellow|green|teal|olive|blue|purple|cranberry|steel|dark \w+|gray|grey|black|maroon)( category)?$/i;
const DEFAULT_WEIGHT = 10, MAX = 8;

export function suggestedTargets(categoryNames: string[], shares: { name: string; share: number }[]): Record<string, number> {
  const used = shares.filter(item => item.name !== "Uncategorized" && item.share > 0).map(item => item.name);
  const named = categoryNames.filter(name => !COLOR_ONLY.test(name.trim()));
  const names = [...new Set([...used, ...named])].slice(0, MAX);
  if (!names.length) return {};
  const weights = names.map(name => RULES.find(rule => rule.test.test(name))?.weight ?? DEFAULT_WEIGHT);
  const total = weights.reduce((sum, weight) => sum + weight, 0);
  // Scale to 100%, rounded to 5% steps, at least 5% each.
  const result: Record<string, number> = {};
  names.forEach((name, index) => { result[name] = Math.max(5, Math.round(weights[index] / total * 100 / 5) * 5); });
  // Nudge the largest so the total is exactly 100 when possible.
  const sum = Object.values(result).reduce((a, b) => a + b, 0), largest = names[weights.indexOf(Math.max(...weights))];
  if (sum !== 100 && result[largest] + (100 - sum) >= 5) result[largest] += 100 - sum;
  return result;
}

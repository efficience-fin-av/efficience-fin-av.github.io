export const EFFICIENCY_MIN = 0.82;

export function efficiencyState(value) {
  if (typeof value !== 'number' || !Number.isFinite(value)) return 'na';
  return value < EFFICIENCY_MIN ? 'bad' : 'good';
}

export function weeklyFeedback(value, week) {
  if (week === 'all') return {kind:'cumulative', deficit:null};
  const state=efficiencyState(value);
  if (state==='na') return {kind:'missing', deficit:null};
  return state==='good'
    ? {kind:'achieved', deficit:0}
    : {kind:'encouragement', deficit:(EFFICIENCY_MIN-value)*100};
}

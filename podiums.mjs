// The full period is supplied here, so a name search never changes the winners.
export function buildPodiums(rows,{metric='global',chef='all'}={}){
 const field=metric==='pure'?'pure':'global';
 const podium=members=>{
  const eligible=members.filter(p=>Number.isFinite(p[field])).sort((a,b)=>b[field]-a[field]||a.name.localeCompare(b.name,'fr')||a.key.localeCompare(b.key));
  return {total:members.length,eligible:eligible.length,leaders:eligible.slice(0,3).map((p,i)=>({...p,place:i+1,tied:(i>0&&eligible[i-1][field]===p[field])||(i+1<eligible.length&&eligible[i+1][field]===p[field])}))};
 };
 const chefs=[...new Set(rows.map(p=>p.chef))].sort((a,b)=>a.localeCompare(b,'fr'));
 return {metric:field,overall:podium(rows),teams:chefs.filter(c=>chef==='all'||c===chef).map(c=>({chef:c,...podium(rows.filter(p=>p.chef===c))}))};
}

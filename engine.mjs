export const EPOCH=Date.UTC(1899,11,30),DAY=86400000;
export const clean=x=>String(x??'').replace(/\u00a0/g,' ').trim();
export const norm=x=>clean(x).normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/œ/gi,'oe').toUpperCase().replace(/[-.'’]/g,' ').replace(/\s+/g,' ').trim();
const nameKey=x=>norm(x).split(' ').filter(Boolean).sort().join(' ');
const num=x=>typeof x==='number'&&Number.isFinite(x)?x:null;
const sum=a=>a.reduce((s,n)=>s+(n??0),0);
const idPart=x=>clean(x).includes('|')?clean(x).split('|')[0].trim().toUpperCase():'';
const labelPart=x=>clean(x).includes('|')?clean(x).slice(clean(x).indexOf('|')+1).trim():clean(x);
export function serialDate(x){if(typeof x==='number')return Math.floor(x);if(x instanceof Date)return Math.floor((x.getTime()-EPOCH)/DAY);if(typeof x==='string'&&/^\d{4}-\d{2}-\d{2}/.test(x))return Math.floor((Date.parse(x)-EPOCH)/DAY);return null;}
export function weekOf(serial){const d=new Date(EPOCH+serial*DAY),j=new Date(Date.UTC(d.getUTCFullYear(),0,1));return Math.floor(((d-j)/DAY+j.getUTCDay())/7)+1;}
function defaultStart(year,w){const j=Date.UTC(year,0,1);return (j-EPOCH)/DAY-new Date(j).getUTCDay()+7*(w-1);}

export function buildModel(raw){
 const issues=[],latest=new Map(),horoNames=new Map();
 const hRows=(raw.horoquartz??[]).map((r,i)=>({id:idPart(r[1]),name:labelPart(r[1]),chef:labelPart(r[0]).toUpperCase(),chefKey:idPart(r[0])||'N:'+norm(labelPart(r[0])),date:serialDate(r[2]),hours:num(r[7]),absence:num(r[10]),row:i+2}));
 for(const h of hRows){if(!h.id||h.date===null){issues.push({type:'Horoquartz',message:'Identifiant ou date non reconnu',name:h.name,row:h.row});continue;}if(!latest.has(h.id)||h.date>=latest.get(h.id).date)latest.set(h.id,h);horoNames.set(nameKey(h.name),horoNames.has(nameKey(h.name))&&horoNames.get(nameKey(h.name))!==h.id?null:h.id);}
 const latestChiefs=new Map();for(const h of hRows)if(h.id&&latest.get(h.id)?.date===h.date){if(!latestChiefs.has(h.id))latestChiefs.set(h.id,new Set());latestChiefs.get(h.id).add(h.chefKey);}
 const conflicts=new Set([...latestChiefs].filter(([,s])=>s.size>1).map(([id])=>id));
 const effNames=[...new Set((raw.efficience??[]).map(r=>clean(r[2])).filter(Boolean))],hpNames=[...new Set((raw.hp??[]).map(r=>clean(r[1])).filter(Boolean))];
 const chooseName=(names,n)=>{const matches=names.filter(s=>nameKey(s)===nameKey(n));return matches.length===1?matches[0]:null;};
 const people=[];const seen=new Set(),sourcePeople=raw.personnel??[],usedRows=new Set(),candidates=[];
 for(const hr of latest.values()){
  const matches=sourcePeople.filter(p=>clean(p.id).toUpperCase()===hr.id);
  const nameMatches=sourcePeople.filter(p=>!clean(p.id)&&nameKey(p.alias||p.name)===nameKey(hr.name));
  const match=matches[0]||(nameMatches.length===1?nameMatches[0]:null);
  if(matches.length>1)issues.push({type:'Personnel',message:'SGID présent plusieurs fois ; une seule fiche est retenue',name:hr.name});
  if(match)usedRows.add(match);candidates.push({...match,id:hr.id,chiefFormula:true,nameFormula:true});
 }
 for(const row of sourcePeople){if(usedRows.has(row)||latest.has(clean(row.id).toUpperCase()))continue;if(!clean(row.id)&&horoNames.get(nameKey(row.alias||row.name)))continue;candidates.push(row);}
 for(const row of candidates){
  let id=clean(row.id).toUpperCase(),alias=clean(row.alias),display=clean(row.name);
  if(!id&&display)id=horoNames.get(nameKey(display))||'';
  const hr=latest.get(id);let name=alias||(!row.nameFormula&&display?display:null)||chooseName(effNames,hr?.name||display)||chooseName(hpNames,hr?.name||display)||display||hr?.name||'';
  if(!name&&!id)continue;if(!name){issues.push({type:'Personnel',message:'SGID absent d’Horoquartz ; nom à compléter',name:id});continue;}
  const key=id||'N:'+norm(name),chef=conflicts.has(id)?'À vérifier':hr?.chef||clean(row.chef)||'Non rattaché';
  if(seen.has(key)){issues.push({type:'Personnel',message:'Salarié en double dans Personnel',name});continue;}seen.add(key);
  const p={id,key,name,chef,alias,horoName:hr?.name||'',weekend:id==='F1638025'||norm(name)==='FABIAN PENOT',index:people.length};people.push(p);
  if(conflicts.has(id))issues.push({type:'Horoquartz',message:'Chefs différents pour ce SGID au dernier jour connu ; attribution à vérifier',name});
  if(!id)issues.push({type:'Personnel',message:'SGID à compléter pour récupérer les heures Horoquartz',name});else if(!hr)issues.push({type:'Personnel',message:'Aucune ligne Horoquartz pour ce SGID',name});
 }
 const homonyms=new Map();for(const p of people){const k=norm(p.name);if(!homonyms.has(k))homonyms.set(k,[]);homonyms.get(k).push(p);}
 for(const group of homonyms.values())if(group.length>1)for(const p of group){issues.push({type:'Personnel',message:'Homonyme : le SGID distingue les fiches ; définir des noms export distincts si nécessaire',name:p.name});if(p.id)p.name+=' ['+p.id+']';}
 const exact=new Map(),canonical=new Map();
 function add(map,key,p){if(!key)return;if(!map.has(key))map.set(key,p);else if(map.get(key)!==p)map.set(key,null);}
 for(const p of people)for(const n of [p.name,p.alias,p.horoName]){add(exact,norm(n),p);add(canonical,nameKey(n),p);}
 const peopleById=new Map(people.filter(p=>p.id).map(p=>[p.id,p]));
 for(const h of hRows){const p=peopleById.get(h.id);if(p){add(exact,norm(h.name),p);add(canonical,nameKey(h.name),p);}}
 const lookup=n=>exact.get(norm(n))||canonical.get(nameKey(n))||null;
 const eff=[],hp=[],effIndex=new Map(),hpIndex=new Map(),horoIndex=new Map(),dupsE=new Map(),dupsH=new Map();
 const push=(m,k,v)=>{if(!m.has(k))m.set(k,[]);m.get(k).push(v);};
 const round=x=>typeof x==='number'?Math.round(x*1e6)/1e6:clean(x);
 for(let i=0;i<(raw.efficience??[]).length;i++){
  const r=raw.efficience[i],p=lookup(r[2]);if(!clean(r[2]))continue;
  if(!p){issues.push({type:'Efficience',message:'Nom absent ou ambigu dans Personnel',name:clean(r[2]),row:i+2});continue;}
  const x={week:Number(r[0]),day:Number(r[1]),key:p.key,name:clean(r[2]),activity:num(r[3]),allocated:num(r[4]),pieces:num(r[5]),row:i+2};
  eff.push(x);push(effIndex,`${x.week}|${x.day}|${p.key}`,x);push(dupsE,[x.week,x.day,p.key,...r.slice(3,6).map(round)].join('|'),x);
 }
 for(let i=0;i<(raw.hp??[]).length;i++){
  const r=raw.hp[i],p=lookup(r[1]);if(!clean(r[1]))continue;
  if(!p){issues.push({type:'HP',message:'Nom absent ou ambigu dans Personnel',name:clean(r[1]),row:i+2});continue;}
  const x={week:Number(r[0]),key:p.key,name:clean(r[1]),reason:clean(r[2]),hours:num(r[3]),row:i+2};hp.push(x);push(hpIndex,`${x.week}|${p.key}`,x);push(dupsH,[x.week,p.key,x.reason,round(r[3])].join('|'),x);
 }
 const uniqueDates=new Set();
 for(const h of hRows){if(!h.id||h.date===null)continue;push(horoIndex,h.id,h);if(!seen.has(h.id))issues.push({type:'Horoquartz',message:'SGID à ajouter dans Personnel',name:h.name,row:h.row});const k=h.id+'|'+h.date;if(uniqueDates.has(k))issues.push({type:'Horoquartz',message:'Plusieurs lignes pour un même salarié et jour',name:h.name,row:h.row});uniqueDates.add(k);}
 let duplicateEff=0,duplicateHp=0;
 for(const [type,map] of [['Efficience',dupsE],['HP',dupsH]])for(const group of map.values())if(group.length>1){if(type==='Efficience')duplicateEff+=group.length;else duplicateHp+=group.length;issues.push({type,message:`${group.length} lignes potentiellement identiques, conservées dans les calculs`,name:group[0].name,rows:group.map(x=>x.row)});}
 const years=hRows.filter(x=>x.date!==null).map(x=>new Date(EPOCH+x.date*DAY).getUTCFullYear());const year=raw.meta?.year||years.sort((a,b)=>years.filter(y=>y===b).length-years.filter(y=>y===a).length)[0]||2026;
 const productionWeeks=[...new Set([...eff.map(x=>x.week),...hp.map(x=>x.week)])].filter(w=>w>=1&&w<=53).sort((a,b)=>a-b);
 const horoWeeks=hRows.filter(x=>x.date!==null).map(x=>Math.floor((x.date-defaultStart(year,1))/7)+1);
 const populated=[...new Set([...productionWeeks,...horoWeeks])].filter(w=>w>=1&&w<=53).sort((a,b)=>a-b);
 const weekStarts=Object.fromEntries(Array.from({length:53},(_,i)=>[i+1,Number(raw.weekStarts?.[i+1])||defaultStart(year,i+1)]));
 const available=[...new Set([...populated,...Object.keys(raw.weekStarts??{}).map(Number)])].filter(w=>w>=1&&w<=53).sort((a,b)=>a-b);
 const cache=new Map();
 function weekly(p,week){
  const k=p.key+'|'+week;if(cache.has(k))return cache.get(k);
  const start=weekStarts[week],dates=p.weekend?[start+6,start+7]:Array.from({length:7},(_,i)=>start+i),daily=dates.map(date=>{
   const d=new Date(EPOCH+date*DAY).getUTCDate(),w=p.weekend?weekOf(date):week,rows=effIndex.get(`${w}|${d}|${p.key}`)||[];
   return {date,activity:rows.length?sum(rows.map(x=>x.activity)):null,allocated:rows.length?sum(rows.map(x=>x.allocated)):null,pieces:rows.length?sum(rows.map(x=>x.pieces)):null,rows};
  });
  const eCount=sum(daily.map(d=>d.rows.length)),activity=eCount?sum(daily.map(x=>x.activity)):null,allocated=eCount?sum(daily.map(x=>x.allocated)):null,pieces=eCount?sum(daily.map(x=>x.pieces)):null;
  const hpRows=hpIndex.get(`${week}|${p.key}`)||[],hpHours=hpRows.length?sum(hpRows.map(x=>x.hours)):null;
  const h=(horoIndex.get(p.id)||[]).filter(x=>x.date>=dates[0]&&x.date<=dates.at(-1)),hours=h.length?sum(h.map(x=>(x.hours??0)-(x.absence??0))):null;
  const pure=activity>0?allocated/activity:null,global=allocated!==null&&hours>(hpHours??0)?allocated/(hours-(hpHours??0)):null;
  const active=activity>0||allocated>0||hpHours>0,missing=active&&!(hours>0);
  const result={...p,week,activity,allocated,pieces,hp:hpHours,horo:hours,pure,global,gamma:activity!==null&&allocated!==null?allocated-activity:null,gap:hours!==null?(activity??0)+(hpHours??0)-hours:null,active,missing,daily,hpRows};cache.set(k,result);return result;
 }
 function aggregatePerson(p,weeks){
  const rows=weeks.map(w=>weekly(p,w)),has=k=>rows.some(r=>r[k]!==null),total=k=>has(k)?sum(rows.map(r=>r[k])):null;
  const activity=total('activity'),allocated=total('allocated'),hp=total('hp'),horo=total('horo'),missing=rows.some(x=>x.missing),pure=activity>0?allocated/activity:null,global=!missing&&horo>(hp??0)&&allocated!==null&&allocated!==0?allocated/(horo-(hp??0)):null;
  return {...p,activity,allocated,pieces:total('pieces'),hp,horo,pure,global:weeks.length===1?rows[0].global:global,gamma:activity!==null&&allocated!==null?allocated-activity:null,gap:horo!==null?(activity??0)+(hp??0)-horo:null,missing,active:rows.some(r=>r.active),weeks:rows};
 }
 function selection({week='all',chef='all',search=''}={}){
  const weeks=week==='all'?populated:[Number(week)],selected=people.filter(p=>(chef==='all'||p.chef===chef)&&(!search||norm(p.name).includes(norm(search))));
  const rows=selected.map(p=>aggregatePerson(p,weeks));return {rows,weeks,people:selected};
 }
 function totals(rows){
  const validGlobal=rows.filter(r=>r.global!==null),a=sum(rows.map(r=>r.activity)),b=sum(rows.map(r=>r.allocated)),g=sum(validGlobal.map(r=>r.horo)),h=sum(validGlobal.map(r=>r.hp)),bg=sum(validGlobal.map(r=>r.allocated));
  return {activity:a,allocated:b,pieces:sum(rows.map(r=>r.pieces)),hp:sum(rows.map(r=>r.hp)),horo:sum(rows.map(r=>r.horo)),pure:a>0?b/a:null,global:g>h?bg/(g-h):null,complete:validGlobal.length,withActivity:rows.filter(r=>r.activity!==null).length,people:rows.length,gapAlerts:rows.filter(r=>r.gap!==null&&r.gap>-1).length,gammaAlerts:rows.filter(r=>r.gamma!==null&&r.gamma<0).length,missing:rows.filter(r=>r.missing).length};
 }
 return {raw,people,issues,eff,hp,hRows,year,populated,weeks:available,weekStarts,duplicateEff,duplicateHp,weekly,selection,totals,chefs:[...new Set(people.map(p=>p.chef))].sort((a,b)=>a.localeCompare(b,'fr')),latestWeek:productionWeeks.at(-1)||populated.at(-1)||available.at(-1)||1};
}

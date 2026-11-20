import {createFireworks} from './celebration.mjs?v=seuil82-20261120';
import {EFFICIENCY_MIN,efficiencyState,weeklyFeedback} from './efficiency-rules.mjs?v=seuil82-20261120';
import {buildModel,clean,norm,EPOCH,DAY} from './engine.mjs?v=direction-20261120';
import {readExcel} from './xlsx-reader.mjs';
import {buildPodiums} from './podiums.mjs';

const $=id=>document.getElementById(id),esc=x=>String(x??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const number=(x,d=1)=>x===null||x===undefined?'n.d.':new Intl.NumberFormat('fr-FR',{minimumFractionDigits:d,maximumFractionDigits:d}).format(x);
const pct=x=>x===null||x===undefined?'n.d.':number(x*100,1)+' %';
const hours=(x,d=1)=>x===null||x===undefined?'n.d.':number(x,d)+' h';
const date=(serial,opts={day:'2-digit',month:'short'})=>new Intl.DateTimeFormat('fr-FR',{...opts,timeZone:'UTC'}).format(new Date(EPOCH+serial*DAY));
const score=x=>x===null?'<span class="na">n.d.</span>':`<span class="score ${efficiencyState(x)}">${pct(x)}</span>`;
const cell=(x,d=2,red=false,green=false)=>x===null?'<span class="na">n.d.</span>':`<span class="${red?'red-value':green?'green-value':''}" title="${esc(number(x,6))} h">${number(x,d)}</span>`;
let model,initialData,currentData,currentSelection,page='overview',topMetric='global',podiumMetric='global',lastVisibleRows=[],toastTimer,loading=false;
let filters={week:'all',chef:'all',search:''};
const fireworks=createFireworks($('celebrationCanvas'));
const celebratedWeeks=new Set();
let dbPromise;
function db(){if(!dbPromise)dbPromise=new Promise(resolve=>{try{const q=indexedDB.open('efficience-dashboard',1);q.onupgradeneeded=()=>q.result.createObjectStore('data');q.onsuccess=()=>resolve(q.result);q.onerror=()=>resolve(null);}catch{resolve(null);}});return dbPromise;}
async function stored(){const d=await db();if(!d)return null;return new Promise(resolve=>{try{const q=d.transaction('data').objectStore('data').get('last');q.onsuccess=()=>resolve(q.result||null);q.onerror=()=>resolve(null);}catch{resolve(null);}});}
async function persist(raw){const d=await db();if(!d)return;return new Promise(resolve=>{const t=d.transaction('data','readwrite');t.objectStore('data').put(raw,'last');t.oncomplete=resolve;t.onerror=resolve;});}
function toast(message){clearTimeout(toastTimer);$('toast').textContent=message;$('toast').hidden=false;toastTimer=setTimeout(()=>$('toast').hidden=true,5000);}
function setPage(next,scroll=true){page=['overview','ranking','podiums','sources'].includes(next)?next:'overview';for(const p of ['overview','ranking','podiums','sources'])$(p+'Page').hidden=p!==page;document.querySelectorAll('[data-page]').forEach(b=>b.classList.toggle('active',b.dataset.page===page));const titles={overview:['Performance des équipes FIN AV','Une lecture commune des résultats, des temps et des écarts.'],ranking:['Classement des performances','Chaque collaborateur, chaque indicateur, au même endroit.'],podiums:['Les podiums de vos équipes','Les trois meilleures performances, dans chaque équipe et au général.'],sources:['Contrôles & sources','Des informations fiables pour lire la performance.']};$('pageTitle').textContent=titles[page][0];$('pageSubtitle').textContent=titles[page][1];if(model)renderSelectionHeading();if(location.hash!=='#'+page)history.replaceState(null,'','#'+page);if(scroll)window.scrollTo({top:0,behavior:'smooth'});}
function periodText(){return filters.week==='all'?'Cumul des semaines renseignées':`Semaine ${filters.week}`;}
function install(raw,{keepChef=false}={}){
 currentData=raw;model=buildModel(raw);const previousChef=filters.chef;filters={week:String(model.latestWeek),chef:keepChef&&model.chefs.includes(previousChef)?previousChef:'all',search:''};
 $('weekFilter').innerHTML='<option value="all">Cumul des semaines</option>'+model.weeks.map(w=>`<option value="${w}">S${w} · ${date(model.weekStarts[w])} – ${date(model.weekStarts[w]+6)}</option>`).join('');
 $('chefFilter').innerHTML='<option value="all">Toutes les équipes</option>'+model.chefs.map(c=>`<option value="${esc(c)}">${esc(c)}</option>`).join('');
 $('weekFilter').value=filters.week;$('chefFilter').value=filters.chef;$('searchInput').value='';$('yearLabel').textContent=model.year;
 $('sourceName').textContent=raw.meta?.filename||'Classeur Excel';$('sourceName').title=$('sourceName').textContent;
 const latest=Math.max(...model.hRows.map(x=>x.date).filter(x=>x!==null));$('sourceDate').textContent=Number.isFinite(latest)?`Horoquartz jusqu’au ${date(latest,{day:'2-digit',month:'long',year:'numeric'})}`:'';
 $('sourceState').textContent=raw.meta?.importedAt?'Import local':'Version partagée';
 $('issueBadge').textContent=model.issues.length||'';renderSources();render();
}
function renderSelectionHeading(){const {weeks,people}=page==='podiums'?model.selection({week:filters.week,chef:filters.chef}):currentSelection;$('periodLabel').textContent=periodText();$('populationLabel').textContent=`${people.length} collaborateur${people.length>1?'s':''}${filters.chef!=='all'?' · '+filters.chef:''}`;
 $('selectionDate').textContent=weeks.length?`${date(model.weekStarts[weeks[0]])} — ${date(model.weekStarts[weeks.at(-1)]+6,{day:'2-digit',month:'long',year:'numeric'})}`:'Aucune semaine';
}
function render(){if(!model)return;currentSelection=model.selection(filters);const t=model.totals(currentSelection.rows);renderSelectionHeading();
 renderExecutive(t);renderWeeklyFeedback(t);renderTeamComparison();renderKpis(t);renderTrend();renderTop();renderPodiums();renderHours(t);renderAlerts(t);renderTable();
 $('qualityText').textContent=model.duplicateEff||model.duplicateHp?`${model.duplicateEff} lignes Efficience et ${model.duplicateHp} lignes HP potentiellement en double dans le fichier. Elles restent incluses.`:model.issues.length?`${model.issues.length} point${model.issues.length>1?'s':''} à vérifier dans les quatre onglets sources.`:'Aucun point de contrôle détecté dans les quatre sources.';
}
function renderKpis(t){
 const badge=x=>x===null?'':`<span class="metric-pill ${efficiencyState(x)}">${x>=EFFICIENCY_MIN?'≥ 82 %':'< 82 %'}</span>`;
 const cards=[{label:'Efficience pure',tone:efficiencyState(t.pure),value:pct(t.pure),note:`Alloué / activité · minimum 82 %`,badge:badge(t.pure),symbol:'◎'},{label:'Temps alloué',value:number(t.allocated)+' <small>h</small>',note:'Charge théorique des opérations réalisées',symbol:'◷'},{label:'Temps de production',value:number(t.activity)+' <small>h</small>',note:`Horoquartz effectif : ${hours(t.horo)}`,symbol:'◷'},{label:'Hors production',value:number(t.hp)+' <small>h</small>',note:t.horo>0?`${number(t.hp/t.horo*100,1)} % du temps Horoquartz effectif`:'Temps hors production renseigné',symbol:'▥'}];
 $('kpiGrid').innerHTML=cards.map(c=>`<article class="kpi ${c.featured?'featured':''}"><div class="kpi-label">${c.label}<span class="kpi-symbol">${c.symbol}</span></div><div class="kpi-value ${c.tone||''}">${c.value}${c.badge||''}</div><div class="kpi-note">${esc(c.note)}</div></article>`).join('');
}

let teamMetric='global';
function renderWeeklyFeedback(t){
 const feedback=weeklyFeedback(t.global,filters.week),box=$('weeklyFeedback');
 if(feedback.kind==='cumulative'){box.hidden=true;fireworks.stop();return;}
 box.hidden=false;
 if(feedback.kind==='missing'){
  fireworks.stop();box.className='weekly-feedback pending';box.innerHTML='<span class="weekly-feedback-icon" aria-hidden="true">i</span><div><strong>La semaine attend encore ses données</strong><p>Complétez ou actualisez les sources pour afficher son résultat.</p></div>';return;
 }
 const won=feedback.kind==='achieved';
 const deficit=feedback.deficit<.01?'moins de 0,01 point':number(feedback.deficit,feedback.deficit<.1?2:1)+' point'+(feedback.deficit>1?'s':'');
 box.className='weekly-feedback '+(won?'achieved':'encouragement');
 box.innerHTML=`<span class="weekly-feedback-icon" aria-hidden="true">${won?'★':'+'}</span><div><strong>${won?'Bravo, le minimum est atteint !':'On garde le cap !'}</strong><p>${won?`${pct(t.global)} d’efficience globale en semaine ${filters.week}. Le seuil de 82 % est atteint.`:`${pct(t.global)} d’efficience globale cette semaine. Encore ${deficit} pour atteindre 82 %. Chaque progrès compte.`}</p></div>${won?`<button class="button secondary" data-replay-fireworks ${fireworks.reducedMotion?'hidden':''}>Rejouer le feu d’artifice</button>`:''}`;
 if(!won){fireworks.stop();return;}
 const key=[filters.week,filters.chef,filters.search,t.global].join('|');
 if(!celebratedWeeks.has(key)&&!fireworks.reducedMotion){celebratedWeeks.add(key);fireworks.play();}
}
function teamResults(){const groups=new Map();for(const r of currentSelection.rows){if(!groups.has(r.chef))groups.set(r.chef,[]);groups.get(r.chef).push(r);}return [...groups].map(([chef,rows])=>({chef,rows,...model.totals(rows)}));}
function renderExecutive(t){
 const active=currentSelection.rows.filter(r=>r.active),complete=active.filter(r=>r.global!==null),coverage=active.length?complete.length/active.length:null;
 const previous=filters.week==='all'?null:model.populated.filter(w=>w<Number(filters.week)).at(-1);
 const before=previous?model.totals(currentSelection.people.map(p=>model.weekly(p,previous))):null;
 const delta=before?.global!==null&&before?.global!==undefined&&t.global!==null?(t.global-before.global)*100:null;
 const deltaText=delta===null?(filters.week==='all'?`${currentSelection.weeks.length} semaines dans le cumul`:'Aucune période précédente comparable'):`${delta>0?'+':''}${number(delta,1)} points vs S${previous}`;
 const teams=teamResults(),above=teams.filter(x=>x.global!==null&&x.global>=EFFICIENCY_MIN).length,eligible=teams.filter(x=>x.global!==null).length;
 const achieved=t.global!==null&&t.global>=EFFICIENCY_MIN,progress=t.global===null?0:Math.min(1,Math.max(0,t.global/EFFICIENCY_MIN));
 const tone=efficiencyState(t.global);
 const status=t.global===null?'Données insuffisantes':achieved?'Minimum 82 % atteint':'Sous le minimum de 82 %';
 $('executiveHero').innerHTML=`<section class="executive-hero"><div class="hero-main"><p class="hero-kicker">${esc(periodText())} · ${esc(filters.chef==='all'?'Toutes les équipes':filters.chef)}</p><div class="hero-label">Efficience globale</div><div class="hero-rate ${tone}">${pct(t.global)}</div><div class="hero-trend ${delta===null?'':delta>=0?'up':'down'}">${esc(deltaText)}</div><p class="hero-formula">Alloué / (Horoquartz effectif − hors production)</p></div><div class="objective-ring"><svg viewBox="0 0 160 160" aria-hidden="true"><circle cx="80" cy="80" r="64" class="ring-track"/><circle cx="80" cy="80" r="64" class="ring-value ${tone}" stroke-dasharray="${progress*402.124} 402.124"/></svg><div><span>MINIMUM</span><strong>82<span>%</span></strong></div><p>${status}</p></div><div class="hero-facts"><div><strong>${above}<small> / ${eligible}</small></strong><span>équipes à 82 % ou plus</span></div><div><strong>${coverage===null?'n.d.':number(coverage*100,0)+' %'}</strong><span>des dossiers actifs calculables</span></div><div><strong>${complete.length}<small> / ${active.length}</small></strong><span>collaborateurs actifs avec taux global</span></div></div></section>`;
 const gammaLoss=currentSelection.rows.reduce((s,r)=>s+(r.gamma!==null&&r.gamma<0?-r.gamma:0),0),gapLoss=currentSelection.rows.reduce((s,r)=>s+(r.gap!==null&&r.gap<-1?-r.gap:0),0);
 const insights=[{tone:t.global===null?'neutral':achieved?'positive':'critical',label:'Lecture de la période',text:t.global===null?'Les heures disponibles ne permettent pas encore de calculer un taux global.':`${number(t.global*100,1)} % d’efficience globale, soit ${number(Math.abs((t.global-EFFICIENCY_MIN)*100),1)} points ${achieved?'au-dessus':'en dessous'} du minimum de 82 %.`},{tone:t.gammaAlerts?'attention':'positive',label:'Temps de gamme',text:t.gammaAlerts?`${t.gammaAlerts} collaborateurs présentent un écart négatif, pour ${hours(gammaLoss)} au total.`:'Aucun écart de gamme négatif dans la sélection.'},{tone:t.gapAlerts?'attention':'positive',label:'Rapprochement des heures',text:t.gapAlerts?`${t.gapAlerts} écarts Horo / Power BI sous −1 h, représentant ${hours(gapLoss)} à rapprocher.`:'Aucun écart Horo / Power BI inférieur à −1 h.'}];
 $('executiveBrief').innerHTML=insights.map(x=>`<article class="brief-card ${x.tone}"><span>${x.label}</span><p>${esc(x.text)}</p></article>`).join('');
}
function renderTeamComparison(){
 const groups=teamResults().sort((a,b)=>(b[teamMetric]??-Infinity)-(a[teamMetric]??-Infinity));
 const max=Math.max(1.2,...groups.map(x=>x[teamMetric]??0))*1.08;
 document.querySelectorAll('#teamMetric button').forEach(b=>b.classList.toggle('selected',b.dataset.teamMetric===teamMetric));
 $('teamComparison').innerHTML=groups.length?groups.map((g,i)=>`<button class="team-result" data-team="${esc(g.chef)}" title="Voir l’équipe ${esc(g.chef)}"><span class="team-label"><span class="team-initial">${esc(g.chef.split(' ').map(s=>s[0]).slice(0,2).join(''))}</span><span><strong>${esc(g.chef)}</strong><small>${g.rows.length} collaborateur${g.rows.length>1?'s':''} · ${g.complete} taux ${g.complete>1?'globaux calculables':'global calculable'}</small></span></span><span class="team-track"><i class="team-target" style="left:${EFFICIENCY_MIN*100/max}%"></i><i class="team-bar ${efficiencyState(g[teamMetric])} ${g[teamMetric]!==null&&g[teamMetric]>=EFFICIENCY_MIN?'achieved':''}" style="width:${Math.max(0,(g[teamMetric]??0)/max*100)}%"></i></span><strong class="team-rate ${efficiencyState(g[teamMetric])}">${pct(g[teamMetric])}</strong></button>`).join(''):'<div class="empty">Aucune équipe pour cette sélection.</div>';
}
function renderTrend(){
 const weeks=model.populated,people=currentSelection.people;
 const points=weeks.map(week=>{const rows=people.map(p=>model.weekly(p,week));return {week,...model.totals(rows)};});
 if(!points.some(p=>p.pure!==null||p.global!==null)){$('trendChart').innerHTML='<div class="empty">Aucune donnée d’efficience pour cette sélection.</div>';return;}
 const W=720,H=270,L=43,R=25,T=24,B=36,range=Math.max(1.2,...points.flatMap(p=>[p.pure??0,p.global??0]))*1.12,max=Math.ceil(range*5)/5,step=max>2?.5:.25;
 const x=i=>points.length===1?(L+(W-R))/2:L+i*(W-L-R)/(points.length-1),y=v=>T+(H-T-B)*(1-v/max);
 let grid='';for(let v=0;v<=max+.0001;v+=step)grid+=`<line x1="${L}" x2="${W-R}" y1="${y(v)}" y2="${y(v)}" stroke="#edf2f6"/><text x="${L-9}" y="${y(v)+4}" text-anchor="end" fill="#9aabba" font-size="12">${Math.round(v*100)}%</text>`;
 const draw=(metric,color)=>{let line='',active=false,shapes='';points.forEach((p,i)=>{const v=p[metric];if(v===null){active=false;return;}line+=`${active?'L':'M'}${x(i)},${y(v)} `;active=true;shapes+=`<circle class="chart-point" data-week="${p.week}" tabindex="0" role="button" aria-label="Semaine ${p.week}, efficience ${metric==='global'?'globale':'pure'} ${pct(v)}" cx="${x(i)}" cy="${y(v)}" r="4.1" fill="${efficiencyState(v)==='bad'?'#c73545':'#168064'}" stroke="white" stroke-width="2"><title>S${p.week} · ${metric==='global'?'Globale':'Pure'} : ${pct(v)}</title></circle>`;});return `<path d="${line}" fill="none" stroke="${color}" stroke-width="2.8" stroke-linejoin="round"/>`+shapes;};
 const selected=points.findIndex(p=>String(p.week)===filters.week),highlight=selected<0?'':`<rect x="${x(selected)-24}" y="${T-8}" width="48" height="${H-T-B+17}" fill="#edf4fd" rx="7"/>`;
 $('trendChart').innerHTML=`<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Évolution de l’efficience pure et globale par semaine">${highlight}${grid}<line x1="${L}" x2="${W-R}" y1="${y(EFFICIENCY_MIN)}" y2="${y(EFFICIENCY_MIN)}" stroke="#a7b6c6" stroke-dasharray="5 5"/>${draw('pure','#19ac92')}${draw('global','#2463eb')}${points.map((p,i)=>`<text data-week="${p.week}" x="${x(i)}" y="${H-10}" text-anchor="middle" fill="${String(p.week)===filters.week?'#2463eb':'#7e94a7'}" font-size="13" font-weight="${String(p.week)===filters.week?700:500}" style="cursor:pointer">S${p.week}</text>`).join('')}</svg>`;
}
function sorted(rows,metric='global'){return [...rows].sort((a,b)=>metric==='name'?a.name.localeCompare(b.name,'fr'):a[metric]===null&&b[metric]===null?a.index-b.index:a[metric]===null?1:b[metric]===null?-1:metric==='gamma'?a[metric]-b[metric]:b[metric]-a[metric]);}
function renderTop(){
 document.querySelectorAll('#topMetric button').forEach(b=>b.classList.toggle('selected',b.dataset.metric===topMetric));const rows=sorted(currentSelection.rows,topMetric).filter(p=>p[topMetric]!==null).slice(0,7),max=rows[0]?.[topMetric]||1;
 $('topRanking').innerHTML=rows.length?rows.map((p,i)=>`<div class="rank-row"><span class="rank-no ${i===0?'first':''}">${i+1}</span><div class="rank-main"><button class="rank-name" data-person="${esc(p.key)}">${esc(p.name)}</button><div class="rank-track"><div class="rank-fill ${efficiencyState(p[topMetric])}" style="width:${Math.min(100,p[topMetric]/max*100)}%"></div></div></div><span class="rank-score ${efficiencyState(p[topMetric])}">${pct(p[topMetric])}</span></div>`).join(''):'<div class="empty">Aucun résultat disponible.</div>';
 $('viewRanking').textContent=`Voir tout le classement (${currentSelection.rows.length})`;
}
function renderPodiums(){
 const result=buildPodiums(model.selection({week:filters.week}).rows,{metric:podiumMetric,chef:filters.chef});
 document.querySelectorAll('#podiumMetric button').forEach(b=>{const selected=b.dataset.podiumMetric===podiumMetric;b.classList.toggle('selected',selected);b.setAttribute('aria-pressed',String(selected));});
 const metricName=podiumMetric==='global'?'Efficience globale':'Efficience pure';
 const trophy='<svg viewBox="0 0 32 32" aria-hidden="true"><path d="M10 5h12v8a6 6 0 0 1-12 0V5Zm0 2H5v4a6 6 0 0 0 6 6m11-10h5v4a6 6 0 0 1-6 6M16 19v6m-6 3h12m-9-3h6"/></svg>';
 const card=(group,title,general=false)=>{
  const count=`${group.eligible} résultat${group.eligible>1?'s':''} disponible${group.eligible>1?'s':''} sur ${group.total}`;
  const slots=Array.from({length:3},(_,i)=>{const p=group.leaders[i],place=i+1,label=['1re','2e','3e'][i];
   return `<li class="podium-place place-${place}${p?'':' vacant'}"${p?` data-podium-person="${esc(p.key)}"`:''}><div class="podium-athlete"><span class="podium-medal" aria-label="${label} place">${place===1?trophy:''}<span>${place}</span></span>${p?`<button class="podium-name" data-person="${esc(p.key)}">${esc(p.name)}</button>${general?`<span class="podium-team">${esc(p.chef)}</span>`:''}<strong class="podium-score ${efficiencyState(p[podiumMetric])}">${pct(p[podiumMetric])}</strong>${p.tied?'<span class="podium-tie">Ex æquo</span>':''}`:'<span class="podium-name">Place disponible</span><span class="podium-unavailable">Données insuffisantes</span>'}</div><div class="podium-step" aria-hidden="true">${place}</div></li>`;
  }).join('');
  return `<article class="podium-card${general?' podium-general':''}" aria-label="${esc(title)}"><div class="podium-card-heading"><div><p class="section-kicker">${general?'TOUTES LES ÉQUIPES':'ÉQUIPE'}</p><h${general?'2':'3'}>${esc(title)}</h${general?'2':'3'}><p class="podium-count">${count}</p></div><span class="podium-metric-label">${metricName}</span></div>${group.leaders.length?`<ol class="podium-stage" aria-label="Les trois meilleurs collaborateurs">${slots}</ol>`:'<div class="podium-empty">Aucun résultat d’efficience disponible pour cette période.</div>'}</article>`;
 };
 $('generalPodium').innerHTML=card(result.overall,'Podium général',true);
 $('teamPodiumsTitle').textContent=filters.chef==='all'?'Podiums par équipe':'Podium de l’équipe sélectionnée';
 $('teamPodiums').innerHTML=result.teams.map(team=>card(team,team.chef)).join('')||'<div class="empty">Aucune équipe disponible pour cette sélection.</div>';
}
function renderHours(t){const max=Math.max(t.horo,t.activity,t.allocated,t.hp,1),items=[['Horoquartz',t.horo,'#b0c3d5'],['Activité',t.activity,'#416f9e'],['Alloué',t.allocated,'#17ab94'],['HP',t.hp,'#eab667']];$('hoursChart').innerHTML=items.map(([l,v,c])=>`<div class="hours-row"><span>${l}</span><div class="hours-track"><div class="hours-fill" style="width:${v/max*100}%;background:${c}"></div></div><span class="hours-value">${hours(v)}</span></div>`).join('')+'<p class="hours-note">Heures cumulées sur les collaborateurs de la sélection. Horoquartz est présenté après déduction des absences.</p>';$('piecesBadge').textContent=`${number(t.pieces,0)} pièces bonnes`;}
function renderAlerts(t){$('alertsSummary').innerHTML=[['gap',t.gapAlerts,'Écart Horo / Power BI','Résultat strictement inférieur à −1 h',''],['gamma',t.gammaAlerts,'Écart gamme négatif','Temps alloué inférieur au temps d’activité','warning'],['missing',t.missing,'Horoquartz manquant','Activité ou HP sans temps Horoquartz positif','info']].map(([key,count,label,sub,cls])=>`<button data-alert="${key}" class="alert-row ${cls}"><span class="alert-count">${count}</span><span class="alert-text">${label}<small>${sub}</small></span><span class="muted">Voir</span></button>`).join('');}
function tableRows(){const alert=$('alertFilter').value;return sorted(currentSelection.rows.filter(r=>alert==='all'||alert==='gap'&&r.gap!==null&&r.gap<-1||alert==='gamma'&&r.gamma!==null&&r.gamma<0||alert==='missing'&&r.missing),$('sortMetric').value);}
function renderTable(){
 const rows=tableRows();lastVisibleRows=rows;$('rankingCount').textContent=rows.length;$('rankingNote').textContent=`${currentSelection.rows.filter(r=>r.global!==null).length} résultats d’efficience globale disponibles. Les collaborateurs sans résultat restent visibles.`;
 const metric=$('sortMetric').value;let rank=0;$('rankingTable').querySelector('tbody').innerHTML=rows.map(r=>{const ranked=['pure','global'].includes(metric)?r[metric]!==null:false;if(ranked)rank++;return `<tr><td>${ranked?rank:'—'}</td><td><button class="person-button" data-person="${esc(r.key)}">${esc(r.name)}</button>${r.weekend?' <span class="na" title="Équipe week-end">WE</span>':''}</td><td>${esc(r.chef)}</td><td>${cell(r.activity)}</td><td>${cell(r.allocated)}</td><td>${cell(r.hp)}</td><td>${cell(r.horo)}</td><td>${score(r.pure)}</td><td>${score(r.global)}</td><td>${cell(r.gamma,3,r.gamma!==null&&r.gamma<0,r.gamma!==null&&r.gamma>0)}</td><td>${cell(r.gap,3,r.gap!==null&&r.gap<-1,r.gap!==null&&r.gap>-1)}</td></tr>`;}).join('');$('tableEmpty').hidden=rows.length>0;
}
function renderSources(){
 const cards=[['Personnel',currentData.personnel.length,`${model.people.length} collaborateurs reconnus`],['Efficience',currentData.efficience.length,`${model.eff.length} lignes rapprochées`],['HP',currentData.hp.length,`${model.hp.length} lignes rapprochées`],['Horoquartz',currentData.horoquartz.length,`${new Set(model.hRows.map(x=>x.id).filter(Boolean)).size} identifiants salariés`]];
 $('sourceCards').innerHTML=cards.map(([n,total,sub])=>`<article class="source-card"><h3>${n}</h3><div class="source-number">${number(total,0)}</div><small>${sub}</small></article>`).join('');
 $('issuesList').innerHTML=model.issues.length?model.issues.map(i=>`<article class="issue-item"><span class="issue-type">${esc(i.type)}</span><div class="issue-main"><strong>${esc(i.name)}</strong><p>${esc(i.message)}</p>${i.rows?`<small>Lignes ${i.rows.join(', ')}</small>`:i.row?`<small>Ligne ${i.row}</small>`:''}</div></article>`).join(''):'<div class="empty">Aucun point de contrôle détecté.</div>';
}
function showPerson(key){
 const r=currentSelection.rows.find(p=>p.key===key)||model.selection({week:filters.week}).rows.find(p=>p.key===key);if(!r)return;$('detailName').textContent=r.name;$('detailChef').textContent=r.chef+(r.weekend?' · ÉQUIPE WEEK-END':'');$('detailPeriod').textContent=periodText()+(r.id?' · SGID '+r.id:' · SGID à compléter');
 const m=[['Efficience pure',pct(r.pure),efficiencyState(r.pure)],['Efficience globale',pct(r.global),efficiencyState(r.global)],['Horoquartz effectif',hours(r.horo,2)],['Hors production',hours(r.hp,2)]];
 const weekly=r.weeks.map(w=>`<tr><td>S${w.week}</td><td>${cell(w.activity)}</td><td>${cell(w.allocated)}</td><td>${cell(w.hp)}</td><td>${cell(w.horo)}</td><td>${score(w.pure)}</td><td>${score(w.global)}</td></tr>`).join('');
 const daily=r.weeks.flatMap(w=>w.daily.map(d=>`<tr><td>${date(d.date,{day:'2-digit',month:'2-digit',year:'numeric'})}</td><td>${cell(d.activity,3)}</td><td>${cell(d.allocated,3)}</td><td>${d.pieces===null?'<span class="na">n.d.</span>':number(d.pieces,0)}</td></tr>`)).join('');
 const reasons=new Map();for(const w of r.weeks)for(const h of w.hpRows){if(h.hours>0)reasons.set(h.reason||'Sans motif',(reasons.get(h.reason||'Sans motif')||0)+h.hours);}
 $('detailContent').innerHTML=`<div class="detail-metrics">${m.map(([l,v,tone])=>`<div class="detail-metric"><span>${l}</span><strong class="${tone||''}">${v}</strong></div>`).join('')}</div><div class="detail-badges"><span class="${r.gamma!==null&&r.gamma<0?'red-value':''}">Écart gamme : ${hours(r.gamma,3)}</span><span class="${r.gap!==null&&r.gap<-1?'red-value':''}">Écart Horo / BI : ${hours(r.gap,3)}</span><span>${number(r.pieces??0,0)} pièces bonnes</span></div>${r.missing?'<p class="import-message error">Une semaine active n’a pas de temps Horoquartz positif. L’efficience globale cumulée est indisponible.</p>':''}<section class="detail-section"><h3>Synthèse par semaine</h3><div class="table-scroll"><table><thead><tr><th>Semaine</th><th>Activité</th><th>Alloué</th><th>HP</th><th>Horo.</th><th>Eff. pure</th><th>Eff. globale</th></tr></thead><tbody>${weekly}</tbody></table></div></section><section class="detail-section"><h3>Production par jour</h3><div class="table-scroll"><table><thead><tr><th>Jour</th><th>Activité (h)</th><th>Alloué (h)</th><th>Pièces bonnes</th></tr></thead><tbody>${daily}</tbody></table></div></section><section class="detail-section"><h3>Motifs hors production</h3>${reasons.size?`<table><thead><tr><th>Motif</th><th>Heures</th></tr></thead><tbody>${[...reasons].sort((a,b)=>b[1]-a[1]).map(([reason,h])=>`<tr><td>${esc(reason)}</td><td>${hours(h,3)}</td></tr>`).join('')}</tbody></table>`:'<p class="muted">Aucune heure HP positive renseignée sur cette période.</p>'}</section>`;$('detailDialog').showModal();
}
async function importFile(file){
 if(!file||loading)return;if(!/\.(xlsx|xlsm)$/i.test(file.name)){toast('Sélectionnez un classeur .xlsx ou .xlsm.');return;}
 loading=true;$('loadButton').disabled=true;$('importMessage').className='import-message';$('importMessage').hidden=false;
 try{const raw=await readExcel(file,m=>$('importMessage').textContent=m);if(!raw.personnel.length||!raw.horoquartz.length)throw new Error('Les onglets Personnel et Horoquartz doivent contenir des données.');install(raw,{keepChef:true});await persist(raw);$('importMessage').hidden=true;toast(`${file.name} chargé · ${model.people.length} collaborateurs reconnus.`);}catch(e){$('importMessage').className='import-message error';$('importMessage').textContent=e.message||'Ce fichier n’a pas pu être lu. Les données précédentes sont conservées.';}finally{loading=false;$('loadButton').disabled=false;$('fileInput').value='';}
}
function exportCsv(){const cols=['Collaborateur','Chef d’équipe','Activité (h)','Alloué (h)','HP (h)','Horoquartz effectif (h)','Efficience pure (%)','Efficience globale (%)','Écart gamme (h)','Écart Horo / Power BI (h)'];const value=x=>{let s=typeof x==='number'?String(Math.round(x*1e6)/1e6).replace('.',','):String(x??'');if(typeof x==='string'&&/^[=+\-@]/.test(s))s="'"+s;return '"'+s.replaceAll('"','""')+'"';};const lines=[cols,...lastVisibleRows.map(r=>[r.name,r.chef,r.activity,r.allocated,r.hp,r.horo,r.pure===null?'n.d.':r.pure*100,r.global===null?'n.d.':r.global*100,r.gamma,r.gap])];const blob=new Blob(['\uFEFF'+lines.map(row=>row.map(value).join(';')).join('\r\n')],{type:'text/csv;charset=utf-8'});const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=`Classement_${filters.week==='all'?'Cumul':'S'+filters.week}_${model.year}.csv`;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1000);toast('Le classement affiché a été exporté.');}
document.querySelectorAll('[data-page]').forEach(b=>b.addEventListener('click',()=>setPage(b.dataset.page)));
$('weekFilter').addEventListener('change',e=>{filters.week=e.target.value;render();});$('chefFilter').addEventListener('change',e=>{filters.chef=e.target.value;render();});let searchTimer;$('searchInput').addEventListener('input',e=>{clearTimeout(searchTimer);searchTimer=setTimeout(()=>{filters.search=e.target.value;render();},130);});
$('resetFilters').addEventListener('click',()=>{filters={week:String(model.latestWeek),chef:'all',search:''};$('weekFilter').value=filters.week;$('chefFilter').value='all';$('searchInput').value='';$('alertFilter').value='all';render();});
$('weeklyFeedback').addEventListener('click',e=>{if(e.target.closest('[data-replay-fireworks]'))fireworks.play();});
$('teamMetric').addEventListener('click',e=>{const b=e.target.closest('[data-team-metric]');if(b){teamMetric=b.dataset.teamMetric;renderTeamComparison();}});
$('teamComparison').addEventListener('click',e=>{const b=e.target.closest('[data-team]');if(b){filters.chef=b.dataset.team;filters.search='';$('chefFilter').value=filters.chef;$('searchInput').value='';render();}});
$('meetingButton').addEventListener('click',async()=>{const on=!document.body.classList.contains('meeting-mode');document.body.classList.toggle('meeting-mode',on);$('meetingButton').textContent=on?'Quitter le mode réunion':'Mode réunion';$('meetingButton').setAttribute('aria-pressed',String(on));if(on){setPage('overview');try{await document.documentElement.requestFullscreen?.();}catch{}}else if(document.fullscreenElement){try{await document.exitFullscreen();}catch{}}});
$('sortMetric').addEventListener('change',renderTable);$('alertFilter').addEventListener('change',renderTable);
$('topMetric').addEventListener('click',e=>{const b=e.target.closest('[data-metric]');if(b){topMetric=b.dataset.metric;renderTop();}});
$('podiumMetric').addEventListener('click',e=>{const b=e.target.closest('[data-podium-metric]');if(b){podiumMetric=b.dataset.podiumMetric;renderPodiums();}});
$('viewPodiums').addEventListener('click',()=>{podiumMetric=topMetric;renderPodiums();setPage('podiums');});
$('viewRanking').addEventListener('click',()=>{$('sortMetric').value=topMetric;$('alertFilter').value='all';renderTable();setPage('ranking');});
$('alertsSummary').addEventListener('click',e=>{const b=e.target.closest('[data-alert]');if(b){$('alertFilter').value=b.dataset.alert;renderTable();setPage('ranking');}});
$('qualityBanner').addEventListener('click',()=>setPage('sources'));
document.addEventListener('click',e=>{const b=e.target.closest('[data-person]');if(b)showPerson(b.dataset.person);const c=e.target.closest('[data-close]');if(c)$(c.dataset.close).close();});
for(const d of [$('detailDialog'),$('helpDialog')])d.addEventListener('click',e=>{if(e.target===d){const r=d.getBoundingClientRect();if(e.clientX<r.left||e.clientX>r.right||e.clientY<r.top||e.clientY>r.bottom)d.close();}});
$('trendChart').addEventListener('click',e=>{const b=e.target.closest('[data-week]');if(b){filters.week=b.dataset.week;$('weekFilter').value=filters.week;render();}});$('trendChart').addEventListener('keydown',e=>{if(['Enter',' '].includes(e.key)&&e.target.dataset.week){e.preventDefault();filters.week=e.target.dataset.week;$('weekFilter').value=filters.week;render();}});
$('helpButton').addEventListener('click',()=>$('helpDialog').showModal());$('loadButton').addEventListener('click',()=>$('fileInput').click());$('fileInput').addEventListener('change',e=>importFile(e.target.files[0]));$('printButton').addEventListener('click',()=>window.print());$('exportButton').addEventListener('click',exportCsv);
$('restoreButton').addEventListener('click',async()=>{install(initialData);await persist(initialData);$('importMessage').hidden=true;toast('Le fichier initial du tableau de bord est à nouveau affiché.');});
let dragCount=0;window.addEventListener('dragenter',e=>{if([...e.dataTransfer.types].includes('Files')){e.preventDefault();dragCount++;$('dropOverlay').hidden=false;}});window.addEventListener('dragleave',e=>{e.preventDefault();if(--dragCount<=0){dragCount=0;$('dropOverlay').hidden=true;}});window.addEventListener('dragover',e=>e.preventDefault());window.addEventListener('drop',e=>{e.preventDefault();dragCount=0;$('dropOverlay').hidden=true;importFile(e.dataTransfer.files[0]);});
window.addEventListener('hashchange',()=>setPage(location.hash.slice(1),false));
try{initialData=window.__INITIAL_DATA__||await (await fetch('./seed.json',{cache:'no-cache'})).json();const saved=await stored();install(saved?.meta?.importedAt?saved:initialData);setPage(location.hash.slice(1)||'overview',false);}catch(e){$('importMessage').hidden=false;$('importMessage').className='import-message error';$('importMessage').textContent='Les données initiales ne sont pas disponibles. Utilisez Actualiser Excel pour charger votre classeur.';console.error(e);}
// Read-only diagnostics used by the dashboard verification. No file is written.
window.efficienceDashboard={get model(){return model;},get selection(){return currentSelection;},get data(){return currentData;},readExcel,importFile};



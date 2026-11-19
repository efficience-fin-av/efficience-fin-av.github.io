// Read-only OOXML loader using browser ZIP decompression and XML parsing.
// No spreadsheet writer and no file-system write permission are used.
const decode=new TextDecoder(),NS='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const nodes=(e,n)=>[...e.getElementsByTagNameNS(NS,n)];
function xml(s){const d=new DOMParser().parseFromString(s,'application/xml');if(d.querySelector('parsererror'))throw new Error('Le fichier contient un composant XML illisible.');return d;}
const path=(base,target)=>{if(target.startsWith('/'))return target.slice(1);const parts=(base+'/'+target).split('/'),r=[];for(const p of parts){if(p==='..')r.pop();else if(p&&p!=='.')r.push(p);}return r.join('/');};
export async function readExcel(file,onProgress=()=>{}){
 const bytes=new Uint8Array(await file.arrayBuffer()),view=new DataView(bytes.buffer);let end=-1;
 for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50){end=i;break;}
 if(end<0)throw new Error('Choisis un classeur Excel .xlsx ou .xlsm.');
 const count=view.getUint16(end+10,true),offset=view.getUint32(end+16,true);const entries=new Map();let pos=offset;
 for(let i=0;i<count;i++){
  if(view.getUint32(pos,true)!==0x02014b50)throw new Error('L’archive Excel est incomplète.');
  const len=view.getUint16(pos+28,true),extra=view.getUint16(pos+30,true),comment=view.getUint16(pos+32,true),name=decode.decode(bytes.subarray(pos+46,pos+46+len));
  entries.set(name,{method:view.getUint16(pos+10,true),size:view.getUint32(pos+20,true),expanded:view.getUint32(pos+24,true),offset:view.getUint32(pos+42,true)});pos+=46+len+extra+comment;
 }
 async function extract(name){
  const e=entries.get(name);if(!e)throw new Error(`Composant Excel absent : ${name}`);
  if(e.expanded>80*1024*1024)throw new Error('Un onglet est trop volumineux pour ce chargement.');
  const p=e.offset,start=p+30+view.getUint16(p+26,true)+view.getUint16(p+28,true),b=bytes.subarray(start,start+e.size);
  if(e.method===0)return decode.decode(b);if(e.method!==8)throw new Error('Compression Excel non prise en charge.');
  let stream;try{stream=new Blob([b]).stream().pipeThrough(new DecompressionStream('deflate-raw'));}catch{throw new Error('Ouvre ce tableau de bord dans une version récente de Chrome ou Edge.');}
  return new Response(stream).text();
 }
 onProgress('Lecture de la structure Excel…');
 const wb=xml(await extract('xl/workbook.xml')),rels=xml(await extract('xl/_rels/workbook.xml.rels')),relmap=new Map([...rels.getElementsByTagName('Relationship')].map(r=>[r.getAttribute('Id'),path('xl',r.getAttribute('Target'))]));
 const sheets=new Map(nodes(wb,'sheet').map(s=>[s.getAttribute('name'),relmap.get(s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id'))]));
 const getSheet=n=>[...sheets.keys()].find(k=>k.normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim().toLowerCase()===n.toLowerCase());
 for(const n of ['Personnel','Efficience','HP','Horoquartz'])if(!getSheet(n))throw new Error(`L’onglet « ${n} » est absent du classeur.`);
 const strings=entries.has('xl/sharedStrings.xml')?nodes(xml(await extract('xl/sharedStrings.xml')),'si').map(si=>nodes(si,'t').map(t=>t.textContent).join('')):[];
 function cell(c){const t=c.getAttribute('t'),v=nodes(c,'v')[0]?.textContent;if(t==='inlineStr')return nodes(c,'t').map(t=>t.textContent).join('');if(t==='s')return strings[Number(v)]??'';if(t==='str')return v??'';if(t==='e')return null;if(t==='b')return v==='1';if(t==='d')return (Date.parse(v)-Date.UTC(1899,11,30))/86400000;return v===undefined?null:Number.isFinite(Number(v))?Number(v):v;}
 function col(a){let n=0;for(const c of a.match(/^[A-Z]+/)[0])n=n*26+c.charCodeAt(0)-64;return n-1;}
 const raw={personnel:[],efficience:[],hp:[],horoquartz:[],weekStarts:{},meta:{filename:file.name,importedAt:new Date().toISOString(),lastModified:file.lastModified}};
 for(const [name,key,maxcol] of [['Personnel','personnel',4],['Efficience','efficience',6],['HP','hp',4],['Horoquartz','horoquartz',15]]){
  onProgress(`Lecture de ${name}…`);await new Promise(r=>setTimeout(r,0));const doc=xml(await extract(sheets.get(getSheet(name))));
  for(const row of nodes(doc,'row')){
   if(Number(row.getAttribute('r'))<2)continue;const a=Array(maxcol).fill(null),formula={};
   for(const c of nodes(row,'c')){const i=col(c.getAttribute('r'));if(i>=maxcol)continue;a[i]=cell(c);formula[i]=nodes(c,'f').length>0;}
   if(!a.some(v=>v!==null&&v!==''))continue;
   if(key==='personnel')raw[key].push({chef:a[0],name:a[1],id:a[2],alias:a[3],chiefFormula:!!formula[0],nameFormula:!!formula[1]});else raw[key].push(a);
  }
 }
 for(const [name,p] of sheets)if(/^S\d{1,2}$/i.test(name)){
  const doc=xml(await extract(p)),c=nodes(doc,'c').find(c=>c.getAttribute('r')==='B3');if(c&&typeof cell(c)==='number')raw.weekStarts[Number(name.slice(1))]=cell(c);
 }
 const dates=raw.horoquartz.map(r=>r[2]).filter(x=>typeof x==='number'),yearCounts=new Map();
 for(const d of dates){const y=new Date(Date.UTC(1899,11,30)+d*86400000).getUTCFullYear();yearCounts.set(y,(yearCounts.get(y)||0)+1);}raw.meta.year=[...yearCounts].sort((a,b)=>b[1]-a[1])[0]?.[0]||Number(file.name.match(/20\d{2}/)?.[0])||2026;
 return raw;
}

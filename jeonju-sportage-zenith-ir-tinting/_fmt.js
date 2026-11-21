window.__fmt=function(html){
 // v2: 링크 통째 보호
 const MAX=21;
 const doc=new DOMParser().parseFromString('<body>'+html+'</body>','text/html');
 const vis=s=>s.replace(/<[^>]+>/g,'').replace(/&[a-z#0-9]+;/g,'x').length;
 const wrapSent=s=>{const ws=s.split(' ');const L=[];let c='';for(const w of ws){if(!w)continue;
   if(c&&(vis(c+' '+w)>MAX || (/[,·]$/.test(c.replace(/<[^>]+>$/,'')) && vis(c)>=11))){L.push(c);c=w}else c=c?c+' '+w:w}
   if(c)L.push(c);
   if(L.length>1&&vis(L[L.length-1])<=4){const last=L.pop();L[L.length-1]+=' '+last}
   return L.join('<br>')};
 const wrap=(inner,group)=>{
  if(inner.includes('<br')) return null;
  let prot=inner.replace(/<a\b[^>]*>[\s\S]*?<\/a>/g,m=>m.replace(/\s+/g,'\u0001'))
             .replace(/<[^>]+>/g,m=>m.replace(/ /g,'\u0001')).replace(/\s+/g,' ').trim();
  prot=prot.replace(/\b([QA])\. /g,'$1.\u0002');
  const sents=prot.split(/(?<=[.?!](?:<\/\w+>)*)\s+(?=\S)/);
  const out=[];sents.forEach((s,i)=>{out.push(wrapSent(s.replace(/\u0002/g,' ')));if(i<sents.length-1)out.push(group&&i%2===1?'<br><br>':'<br>')});
  return out.join('').replace(/\u0001/g,' ');
 };
 let n=0;
 doc.body.querySelectorAll('p,li').forEach(el=>{
  if(el.querySelector('img,video,table'))return;
  if(el.tagName==='LI'&&el.querySelector('a'))return;
  const card=el.closest('div[style]'); if(card&&/border-radius:\s*8px/.test(card.getAttribute('style')))return;
  if(el.closest('td,th'))return;
  if(vis(el.innerHTML)<=MAX)return;
  const r=wrap(el.innerHTML, el.tagName==='P'); if(r!==null){el.innerHTML=r;n++}
 });
 return {html:doc.body.innerHTML,n};
};

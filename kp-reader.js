/* Local legacy Word 97–2003 reader (OLE compound file, Word piece table). */
function readLegacyWord(buffer){
 const a=new Uint8Array(buffer),v=new DataView(buffer); const u=(p)=>v.getUint32(p,true),s=(p)=>v.getUint16(p,true);
 if(a[0]!==208||a[1]!==207)throw Error('Неизвестный формат DOC');
 const size=2**s(30),mini=2**s(32),END=0xfffffffe; if(![512,4096].includes(size)||mini!==64)throw Error('Некорректный DOC');
 const sector=n=>{if(n>=Math.floor(a.length/size)-1)throw Error('Повреждён DOC');return a.subarray((n+1)*size,(n+2)*size)};
 const words=b=>new Uint32Array(b.buffer,b.byteOffset,b.byteLength/4); let fatIds=[];for(let i=0;i<109;i++){let n=u(76+i*4);if(n<END)fatIds.push(n)}
 let dif=u(68);for(let i=0;i<u(72)&&dif<END;i++){let w=words(sector(dif));fatIds.push(...Array.from(w.slice(0,-1)).filter(n=>n<END));dif=w[w.length-1]}
 const fat=fatIds.flatMap(n=>Array.from(words(sector(n))));
 function chain(n,table,get,limit=a.length){let chunks=[],seen=new Set(),len=0;while(n<END&&len<limit){if(seen.has(n)||n>=table.length)throw Error('Повреждена цепочка DOC');seen.add(n);let b=get(n);chunks.push(b);len+=b.length;n=table[n]}let out=new Uint8Array(Math.min(len,limit)),p=0;for(let b of chunks){let c=b.subarray(0,out.length-p);out.set(c,p);p+=c.length}return out}
 let dir=chain(u(48),fat,sector),entries={};let dv=new DataView(dir.buffer);for(let p=0;p+128<=dir.length;p+=128){let l=dv.getUint16(p+64,true);if(l<2||l>64)continue;let name=new TextDecoder('utf-16le').decode(dir.subarray(p,p+l-2));entries[name]={start:dv.getUint32(p+116,true),length:dv.getUint32(p+120,true),type:dir[p+66]}}
 let root=Object.values(entries).find(x=>x.type===5);if(!root)throw Error('Нет корневого потока DOC');let ms=chain(root.start,fat,sector,root.length),mf=Array.from(words(chain(u(60),fat,sector,u(64)*size)));
 function stream(name){let e=entries[name];if(!e)throw Error('Нет потока '+name);return e.length<4096?chain(e.start,mf,n=>ms.subarray(n*mini,(n+1)*mini),e.length):chain(e.start,fat,sector,e.length)}
 let w=stream('WordDocument'),f=new DataView(w.buffer),flags=f.getUint16(10,true);if(flags&0x100)throw Error('DOC защищён паролем');let t=stream(flags&0x200?'1Table':'0Table');let p=32; p+=2+2*f.getUint16(p,true);p+=2+4*f.getUint16(p,true);let pairs=f.getUint16(p,true);p+=2;if(pairs<=33)throw Error('Старая версия DOC не поддерживается');let fc=f.getUint32(p+33*8,true),lc=f.getUint32(p+33*8+4,true);let cl=t.subarray(fc,fc+lc),cv=new DataView(cl.buffer,cl.byteOffset,cl.length);let q=0;while(cl[q]===1)q+=3+cv.getUint16(q+1,true);if(cl[q++]!==2)throw Error('Нет таблицы текста DOC');let n=(cv.getUint32(q,true)-4)/12;q+=4;if(!Number.isInteger(n)||n<1)throw Error('Некорректная таблица DOC');let out='';for(let i=0;i<n;i++){let count=cv.getUint32(q+4*(i+1),true)-cv.getUint32(q+4*i,true),off=cv.getUint32(q+4*(n+1)+8*i+2,true),compressed=!!(off&0x40000000);off&=0x3fffffff;if(compressed)off/=2;out+=new TextDecoder(compressed?'windows-1251':'utf-16le').decode(w.subarray(off,off+count*(compressed?1:2)))}return out.replace(/[\x07\x0b\x0c\r]/g,'\n').replace(/[\x00-\x06\x08\x0e-\x1f]/g,'');
}
function kpPdfText(items,width){
 const clean=items.filter(x=>x.str&&x.str.trim()&&x.transform);
 const lines=arr=>{const rows=[];for(const it of arr){const y=it.transform[5];let row=rows.find(r=>Math.abs(r.y-y)<3);if(!row){row={y,items:[]};rows.push(row)}row.items.push(it)}return rows.sort((a,b)=>b.y-a.y).map(r=>r.items.sort((a,b)=>a.transform[4]-b.transform[4]).map(x=>x.str).join(' ')).join('\n')};
 const all=lines(clean);if(!/свидетельство о технической компетентности/i.test(all)||!/аттестат соответствия на/i.test(all))return items.map(it=>it.str+(it.hasEOL?'\n':' ')).join('');
 const left=clean.filter(x=>x.transform[4]<width*.52),right=clean.filter(x=>x.transform[4]>=width*.52);const l=lines(left),r=lines(right);if(/свидетельство о технической компетентности/i.test(l)&&/аттестат соответствия на/i.test(r))return l+'\n'+r;return items.map(it=>it.str+(it.hasEOL?'\n':' ')).join('');
}
function kpSections(text){
 const lines=text.split(/\r?\n/).map(x=>x.trim()).filter(Boolean), spk=[],att=[],licenses=[],tools=[];let mode='';
 for(const line of lines){if(/лицензия\s+(мчс|мвд)/i.test(line)&&/\(/.test(line))licenses.push(line);
 if(/^свидетельство о технической компетентности\s*:/i.test(line)){mode='spk';continue}if(/^аттестат соответствия на\s*:/i.test(line)){mode='att';continue}if(/^перечень (?:необходимого оборудования|инструментов)/i.test(line)){mode='tools';continue}if(/^(?:важность наличия|ценовое предложение|условия сотрудничества|требования к аттестату|преимущества работы|услуга по подбору)/i.test(line)){mode='';continue}if(/^виды работ\s*:/i.test(line)){if(!mode)mode='spk';continue}
 if(mode==='spk'||mode==='att'){const target=mode==='spk'?spk:att;if(target.length&&!/^\d+[.)]|^[-•]|^(?:устройство|монтаж|возведение|строительство|заполнение|штукатурные|облицовочные|малярные|обойные|стекольные|выполнение|геодезические|тепловая|внутренние|наружные)/i.test(line)&&!/[;.]$/.test(target[target.length-1]))target[target.length-1]+=' '+line;else target.push(line)}if(mode==='tools')tools.push(line)}
 return {spk,att,licenses:[...new Set(licenses)],tools};
}


(function(){
function crc32(bytes){let table=crc32.table;if(!table){table=[];for(let n=0;n<256;n++){let c=n;for(let k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);table[n]=c>>>0}crc32.table=table}
let c=0xFFFFFFFF;for(const b of bytes)c=table[(c^b)&255]^(c>>>8);return (c^0xFFFFFFFF)>>>0}
const te=new TextEncoder();
function u16(n){return new Uint8Array([n&255,(n>>>8)&255])}
function u32(n){return new Uint8Array([n&255,(n>>>8)&255,(n>>>16)&255,(n>>>24)&255])}
function cat(...arrs){let n=arrs.reduce((a,b)=>a+b.length,0),o=new Uint8Array(n),p=0;for(const a of arrs){o.set(a,p);p+=a.length}return o}
function zip(entries){
 const parts=[],central=[];let offset=0;
 for(const [name,data] of entries){
  const nb=te.encode(name), b=data instanceof Uint8Array?data:te.encode(data), crc=crc32(b);
  const local=cat(new Uint8Array([80,75,3,4]),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(b.length),u32(b.length),u16(nb.length),u16(0),nb,b);
  parts.push(local);
  central.push(cat(new Uint8Array([80,75,1,2]),u16(20),u16(20),u16(0),u16(0),u16(0),u16(0),u32(crc),u32(b.length),u32(b.length),u16(nb.length),u16(0),u16(0),u16(0),u16(0),u32(0),u32(offset),nb));
  offset+=local.length;
 }
 const c=cat(...central), body=cat(...parts), end=cat(new Uint8Array([80,75,5,6]),u16(0),u16(0),u16(entries.length),u16(entries.length),u32(c.length),u32(body.length),u16(0));
 return cat(body,c,end);
}
function esc(s){return String(s??"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;").replace(/'/g,"&apos;")}
function sheetXml(rows){
 let r=rows.map((row,ri)=>`<row r="${ri+1}">`+row.map((v,ci)=>{let col="",n=ci+1;while(n){let x=(n-1)%26;col=String.fromCharCode(65+x)+col;n=Math.floor((n-1)/26)}return `<c r="${col}${ri+1}" t="inlineStr"><is><t xml:space="preserve">${esc(v)}</t></is></c>`}).join("")+"</row>").join("");
 return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>${r}</sheetData></worksheet>`;
}
function xlsx(rowsBySheet){
 const names=Object.keys(rowsBySheet), entries=[];
 entries.push(["[Content_Types].xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${names.map((_,i)=>`<Override PartName="/xl/worksheets/sheet${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>`]);
 entries.push(["_rels/.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`]);
 entries.push(["xl/workbook.xml",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>${names.map((n,i)=>`<sheet name="${esc(n)}" sheetId="${i+1}" r:id="rId${i+1}"/>`).join("")}</sheets></workbook>`]);
 entries.push(["xl/_rels/workbook.xml.rels",`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${names.map((_,i)=>`<Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i+1}.xml"/>`).join("")}</Relationships>`]);
 names.forEach((n,i)=>entries.push([`xl/worksheets/sheet${i+1}.xml`,sheetXml(rowsBySheet[n])]));
 return new Blob([zip(entries)],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
}
window.WorkLogXLSX={xlsx};
})();

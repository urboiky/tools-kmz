const $=x=>document.getElementById(x);let ports=[];
function typ(c){c=String(c).trim().toUpperCase();return c.endsWith("01")?"SB":c.endsWith("02")?"EB":"NODE"}
function cv(ws,a){return ws[a]?ws[a].v:""}
function parse(ws){let rg=XLSX.utils.decode_range(ws["!ref"]||"A1:I4"),out=[],cur=null;
 for(let r=4;r<=rg.e.r+1;r++){let b=String(cv(ws,"B"+r)||"").trim(),c=String(cv(ws,"C"+r)||"").trim(),v=cv(ws,"I"+r);if(!c)continue;
  if(b){cur={line:b,items:[]};out.push(cur)} if(!cur)continue;
  let n=typeof v==="number"?v:Number(String(v).replace(",","."));cur.items.push({row:r,code:c,length:Number.isFinite(n)?n:"",type:typ(c)})}return out}
function demo(){ports=[
 {line:"A",items:[{row:4,code:"FXXXXXS01A01",length:62,type:"SB"},{row:5,code:"FXXXXXS01A02",length:46,type:"EB"}]},
 {line:"B",items:[{row:6,code:"FXXXXXS01B01",length:32,type:"SB"},{row:7,code:"FXXXXXS01B02",length:88,type:"EB"}]},
 {line:"C",items:[{row:8,code:"FXXXXXS02C01",length:33,type:"SB"},{row:9,code:"FXXXXXS02C02",length:68,type:"EB"}]},
 {line:"D",items:[{row:10,code:"FXXXXXS02D01",length:195,type:"SB"},{row:11,code:"FXXXXXS02D02",length:54,type:"EB"}]},
 {line:"E",items:[{row:12,code:"FXXXXXS03E01",length:106,type:"SB"},{row:13,code:"FXXXXXS03E02",length:61,type:"EB"}]},
 {line:"F",items:[{row:14,code:"FXXXXXS03F01",length:203,type:"SB"},{row:15,code:"FXXXXXS03F02",length:203,type:"EB"}]}];render()}
function layout(){const n=Math.max(1,ports.length),gap=105,y0=80,H=Math.max(480,y0+(n-1)*gap+90);
 const hbY=y0+(n-1)*gap/2;
 return{n,gap,y0,H,W:1120,hbX:105,hbY,fanStart:205,fanStep:18,portX:360,sbX:650,ebX:960}}
function hex(cx,cy,r){const a=[];for(let i=0;i<6;i++){let ang=Math.PI/3*i;a.push([cx+r*Math.cos(ang),cy+r*Math.sin(ang)])}return a}
function ps(a){return a.map(q=>q[0].toFixed(2)+","+q[1].toFixed(2)).join(" ")}
function table(){let s="<table><tr><th>Row</th><th>Line</th><th>FAT CODE</th><th>I</th></tr>";ports.forEach(p=>p.items.forEach((x,i)=>s+=`<tr><td>${x.row}</td><td>${i?"↳":p.line}</td><td>${x.code}</td><td>${x.length}m</td></tr>`));$("table").innerHTML=s+"</table>"}
function render(){table();const q=layout();let z=`<rect width="${q.W}" height="${q.H}" fill="#222a31"/>`;
 z+=`<polygon points="${ps(hex(q.hbX,q.hbY,48))}" fill="none" stroke="#ff2525" stroke-width="2"/><text x="${q.hbX}" y="${q.hbY}" text-anchor="middle" dominant-baseline="middle" fill="#ff2525" font-size="23">HB</text><text x="${q.hbX}" y="${q.hbY+70}" text-anchor="middle" fill="#fff" font-size="10">FXXXXX</text>`;
 ports.forEach((p,i)=>{const y=q.y0+i*q.gap;
   // unique trunk position, then a clean horizontal lead-in to the port line
   const center=(ports.length-1)/2, d=i-center;
   // Symmetric fan: closest-to-center ports leave HB first; outer ports step outward evenly.
   const bx=q.fanStart+Math.abs(d)*q.fanStep+(d>0?q.fanStep/2:0);
   z+=`<polyline points="${q.hbX+48},${q.hbY} ${bx},${q.hbY} ${bx},${y} ${q.portX},${y}" fill="none" stroke="#164cff" stroke-width="2"/>`;
   z+=`<text x="${q.portX+10}" y="${y-16}" fill="#fff" font-size="12">PORT ${i+1} FOR LINE ${p.line}</text>`;
   let prev=q.portX,prevNode=false;
   p.items.slice(0,2).forEach((n,k)=>{const x=k===0?q.sbX:q.ebX, R=14;
     const x1=prevNode?prev+R:q.portX, x2=x-R;
     z+=`<line x1="${x1}" y1="${y}" x2="${x2}" y2="${y}" stroke="#164cff" stroke-width="2"/>`;
     z+=`<polygon points="${ps(hex(x,y,R))}" fill="none" stroke="#ff2525" stroke-width="2"/>
     <text x="${x}" y="${y}" text-anchor="middle" dominant-baseline="middle" fill="#ff2525" font-size="7">${n.type}</text>
     <text x="${x}" y="${y-22}" text-anchor="middle" fill="#fff" font-size="10">${n.code}</text>
     <text x="${x}" y="${y+32}" text-anchor="middle" fill="#fff" font-size="10">${n.length}m</text>`;
     prev=x;prevNode=true});
 });$("svg").setAttribute("viewBox",`0 0 ${q.W} ${q.H}`);$("svg").setAttribute("width",q.W);$("svg").setAttribute("height",q.H);$("svg").innerHTML=z;$("status").textContent=ports.length+" PORT • REGULAR HEXAGON"}
$("file").onchange=async e=>{try{let f=e.target.files[0];if(!f)return;let wb=XLSX.read(await f.arrayBuffer(),{type:"array"}),ws=wb.Sheets[wb.SheetNames[0]];ports=parse(ws);if(!ports.length)throw Error("Data B4/C4/I4 tidak ditemukan");render()}catch(e){$("status").textContent=e.message}};$("demo").onclick=demo;
function pair(c,v){return c+"\r\n"+v+"\r\n"}
function L(x1,y1,x2,y2,lay="0",col=7){return pair(0,"LINE")+pair(8,lay)+pair(62,col)+pair(10,x1.toFixed(3))+pair(20,(-y1).toFixed(3))+pair(30,0)+pair(11,x2.toFixed(3))+pair(21,(-y2).toFixed(3))+pair(31,0)}
function T(x,y,h,s,lay="TEXT",col=7,center=false){
 s=String(s).replace(/[^\x20-\x7E]/g,"");
 let out=pair(0,"TEXT")+pair(8,lay)+pair(62,col)+pair(10,x.toFixed(3))+pair(20,(-y).toFixed(3))+pair(30,0)+pair(40,h)+pair(1,s)+pair(50,0);
 if(center){
   // DXF TEXT alignment: 72=1 center horizontal, 73=2 middle vertical.
   // Alignment point is group 11/21/31.
   out+=pair(72,1)+pair(73,2)+pair(11,x.toFixed(3))+pair(21,(-y).toFixed(3))+pair(31,0);
 }
 return out;
}
function P(a,lay,col){let s="";for(let i=0;i<a.length;i++){let b=a[(i+1)%a.length];s+=L(a[i][0],a[i][1],b[0],b[1],lay,col)}return s}
$("export").onclick=()=>{if(!ports.length)return alert("Upload Excel dulu.");const q=layout();let e="";
 e+=P(hex(q.hbX,q.hbY,48),"HB",1)+T(q.hbX,q.hbY,20,"HB","HB",1,true)+T(q.hbX-18,q.hbY+69,8,"FXXXXX","TEXT",7);
 ports.forEach((p,i)=>{const y=q.y0+i*q.gap,center=(ports.length-1)/2,d=i-center,bx=q.fanStart+Math.abs(d)*q.fanStep;
   e+=L(q.hbX+48,q.hbY,bx,q.hbY,"ROUTE",5)+L(bx,q.hbY,bx,y,"ROUTE",5)+L(bx,y,q.portX,y,"ROUTE",5)+T(q.portX+10,y-16,9,`PORT ${i+1} FOR LINE ${p.line}`,"TEXT",7);
   let prev=q.portX,prevNode=false;p.items.slice(0,2).forEach((n,k)=>{const x=k===0?q.sbX:q.ebX,R=14,x1=prevNode?prev+R:q.portX,x2=x-R;
   e+=L(x1,y,x2,y,"CABLE",5)+P(hex(x,y,R),n.type,1)+T(x,y,5,n.type,n.type,1,true)+T(x-36,y-21,7,n.code,"TEXT",7)+T(x-8,y+31,7,n.length+"m","TEXT",7);prev=x;prevNode=true});
 });
 const d=pair(0,"SECTION")+pair(2,"HEADER")+pair(9,"$ACADVER")+pair(1,"AC1009")+pair(0,"ENDSEC")+pair(0,"SECTION")+pair(2,"TABLES")+pair(0,"ENDSEC")+pair(0,"SECTION")+pair(2,"BLOCKS")+pair(0,"ENDSEC")+pair(0,"SECTION")+pair(2,"ENTITIES")+e+pair(0,"ENDSEC")+pair(0,"EOF");
 const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([d],{type:"application/dxf"}));let typed=String($("filename")?.value||"").trim();
 let chosen=window.prompt("Nama file DXF:",typed||"line-diagram");
 if(chosen===null){URL.revokeObjectURL(a.href);a.remove();return;}
 let fn=String(chosen).trim().replace(/[\/:*?"<>|]+/g,"_").replace(/\.dxf$/i,"");
 if(!fn)fn="line-diagram";
 if($("filename"))$("filename").value=fn;
 a.download=fn+".dxf";document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(a.href),1000)};
demo();
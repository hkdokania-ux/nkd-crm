// v2.1
import{useState,useEffect,useRef,useMemo,Fragment}from"react";
import*as XLSX from"xlsx";
import{jsPDF}from"jspdf";
import{createClient}from"@supabase/supabase-js";
const _sb=createClient("https://acwpqkywhxxnxylewtbo.supabase.co","eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImFjd3Bxa3l3aHh4bnh5bGV3dGJvIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODQ2MjI0MDYsImV4cCI6MjEwMDE5ODQwNn0.0_d4LLJ1loNklPCVayHk8M2B0rn-r8ut66siLcRbLt8");
function _dbGet(k){return Promise.resolve(_sb.from("store").select("value").eq("key",k).maybeSingle()).then(({data})=>data?.value??null).catch(()=>null);}
function _dbSet(k,v){Promise.resolve(_sb.from("store").upsert({key:k,value:v})).catch(()=>{});}

const RC={DY13:{n:"CT110 X ES NXT",cat:"CT",ex:71184,cAcc:0,hdl:600,ins:6834,reg:8506,onRoad:88824,amc:1050},DY08:{n:"CT 110 X ES",cat:"CT",ex:65207,cAcc:0,hdl:600,ins:6716,reg:8087,onRoad:82310,amc:1050},JK39:{n:"Platina 100 ES NXT",cat:"PLATINA",ex:69142,cAcc:0,hdl:600,ins:6788,reg:6793,onRoad:85023,amc:1050},JK38:{n:"Platina 110 Drum ES NXT",cat:"PLATINA",ex:73142,cAcc:0,hdl:600,ins:6872,reg:8643,onRoad:90957,amc:1050},JH31:{n:"Freedom 125 Drum OBD2A",cat:"FREEDOM",ex:91625,cAcc:788,hdl:600,ins:7236,reg:9937,onRoad:111886,amc:1322},JH36:{n:"Freedom 125 Disc LED",cat:"FREEDOM",ex:111604,cAcc:788,hdl:600,ins:7631,reg:13567,onRoad:135890,amc:1322},JH37:{n:"Freedom 125 Drum LED",cat:"FREEDOM",ex:104117,cAcc:788,hdl:600,ins:7484,reg:12894,onRoad:127583,amc:1322},JZ51:{n:"Pulsar N125 Top",cat:"N125",ex:92530,cAcc:788,hdl:600,ins:7254,reg:10000,onRoad:112872,amc:1322},JZ53:{n:"Pulsar N125 Mid",cat:"N125",ex:86836,cAcc:738,hdl:600,ins:7254,reg:9602,onRoad:106730,amc:1322},JZ56:{n:"Pulsar N125 Top OBD2B",cat:"N125",ex:92990,cAcc:788,hdl:600,ins:7254,reg:10032,onRoad:113364,amc:1322},DH65:{n:"PULSAR 125 Neon DISC",cat:"P125",ex:82081,cAcc:788,hdl:600,ins:7001,reg:9269,onRoad:101439,amc:1322},DH62:{n:"PULSAR 125 Carbon Disc",cat:"P125",ex:86175,cAcc:788,hdl:600,ins:7114,reg:9555,onRoad:105932,amc:1322},DH70:{n:"Pulsar 125 CF Split Seat LED",cat:"P125",ex:93934,cAcc:788,hdl:600,ins:7282,reg:10098,onRoad:114402,amc:1322},DH71:{n:"Pulsar 125 CF Disc LED",cat:"P125",ex:91757,cAcc:788,hdl:600,ins:7240,reg:9946,onRoad:112031,amc:1322},JF58:{n:"Pulsar NS 125 UG OBD2B",cat:"NS125",ex:99409,cAcc:876,hdl:600,ins:7390,reg:10482,onRoad:120457,amc:1322},JF60:{n:"Pulsar NS 125 ABS OBD2B",cat:"NS125",ex:99999,cAcc:876,hdl:600,ins:7402,reg:10523,onRoad:121100,amc:1322},DH60:{n:"PULSAR 150 SD UG",cat:"P150",ex:107941,cAcc:876,hdl:600,ins:7559,reg:13238,onRoad:131914,amc:1322},DH69:{n:"PULSAR 150 TD UG",cat:"P150",ex:117794,cAcc:876,hdl:600,ins:7753,reg:14124,onRoad:142847,amc:1322},DH63:{n:"PULSAR 150 TD OBD2B",cat:"P150",ex:111982,cAcc:876,hdl:600,ins:7639,reg:13601,onRoad:136398,amc:1322},JF43:{n:"PULSAR NS 160 UG",cat:"NS160",ex:129476,cAcc:876,hdl:600,ins:12130,reg:15176,onRoad:159958,amc:1699},JF55:{n:"PULSAR NS 160 Twin Channel",cat:"NS160",ex:124638,cAcc:876,hdl:600,ins:12035,reg:14740,onRoad:154589,amc:1699},JF45:{n:"PULSAR NS 160 OBD2B",cat:"NS160",ex:133139,cAcc:876,hdl:600,ins:12202,reg:15506,onRoad:164023,amc:1699},JR50:{n:"Pulsar N 160 TD",cat:"N160",ex:114053,cAcc:876,hdl:600,ins:11826,reg:13788,onRoad:142843,amc:1699},JR37:{n:"Pulsar N 160",cat:"N160",ex:117658,cAcc:876,hdl:600,ins:11897,reg:14112,onRoad:146843,amc:1699},JR33:{n:"Pulsar N 160 UG OBD2B",cat:"N160",ex:129381,cAcc:876,hdl:600,ins:12128,reg:15167,onRoad:159852,amc:1699},JR83:{n:"Pulsar N160 Split Seat USD",cat:"N160",ex:130355,cAcc:876,hdl:600,ins:12148,reg:15255,onRoad:160934,amc:1699},JL30:{n:"PULSAR NS 200 Twin Channel",cat:"NS200",ex:137795,cAcc:876,hdl:600,ins:12294,reg:15925,onRoad:169190,amc:1699},JL25:{n:"PULSAR NS 200 OBD2B",cat:"NS200",ex:147978,cAcc:876,hdl:600,ins:12495,reg:16841,onRoad:180490,amc:1699},DJ16:{n:"PULSAR 180",cat:"P180",ex:123552,cAcc:876,hdl:600,ins:12013,reg:14643,onRoad:153384,amc:1699},DT22:{n:"PULSAR RS 200",cat:"RS200",ex:177384,cAcc:0,hdl:600,ins:13075,reg:19488,onRoad:212247,amc:1699},JR34:{n:"PULSAR N 250 OBD2B",cat:"N250",ex:139405,cAcc:876,hdl:600,ins:12326,reg:16069,onRoad:170976,amc:1699},PD39:{n:"AVENGER 160 Street",cat:"AVENGER",ex:113103,cAcc:876,hdl:600,ins:11807,reg:13702,onRoad:141788,amc:1699},PD40:{n:"AVENGER 220 STREET",cat:"AVENGER",ex:131293,cAcc:876,hdl:600,ins:12166,reg:15339,onRoad:161974,amc:1699},JF53:{n:"Dominar 250 UG",cat:"DOMINAR",ex:182924,cAcc:876,hdl:600,ins:13185,reg:19986,onRoad:219271,amc:1699},JF62:{n:"Dominar 400 UG",cat:"DOMINAR",ex:204370,cAcc:876,hdl:600,ins:13608,reg:21916,onRoad:243070,amc:1699},DK14:{n:"Pulsar 220 F LED",cat:"P220F",ex:134732,cAcc:0,hdl:600,ins:12234,reg:15649,onRoad:165791,amc:1699},JL32:{n:"Pulsar NS 400Z OBD2B",cat:"NS400",ex:182174,cAcc:876,hdl:600,ins:13170,reg:19919,onRoad:218439,amc:1699}};

const SM=["Amit Kumar","Ravi Singh","Suresh Yadav","Priya Sharma","Deepak Gupta"];
const BRANCHES=["Hirak Road","Saraidhela","Chirkunda"];
const OFFICE_WA="7033099010";
const SM_BRANCH={"Amit Kumar":"Hirak Road","Ravi Singh":"Hirak Road","Suresh Yadav":"Saraidhela","Priya Sharma":"Saraidhela","Deepak Gupta":"Chirkunda"};
const ST_C={Hot:"#ef4444",Warm:"#f97316",Cold:"#3b82f6",Booked:"#8b5cf6",Billed:"#10b981",Lost:"#6b7280"};
const FU={Hot:1,Warm:3,Cold:7};
const CATS=[...new Set(Object.values(RC).map(r=>r.cat))];
const DEFAULT_USERS={manager:[{name:"Manager",pin:"1234"}],owner:[{name:"Owner",pin:"0000"}],admin:[{name:"Admin",pin:"9999"}],tech:[{name:"Tech",pin:"1111"}]};
// role helpers
function isOwner(r){return r==="owner"||r==="tech";}
function isPortalRole(r){return r==="owner"||r==="admin"||r==="tech";}

function td(){return new Date().toISOString().split("T")[0];}
function aD(d,n){const x=new Date(d);x.setDate(x.getDate()+n);return x.toISOString().split("T")[0];}
function fd(d){if(!d)return"—";return new Date(d).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"});}
function fc(n){return"₹"+(Number(n)||0).toLocaleString("en-IN");}
const DRIVE_UPLOAD_URL="https://script.google.com/macros/s/AKfycbzYG7g9u3zOy87f5Nl8a0s7Mf6QGH5bi0mVueyzwYG5FYZLAy7RSSGgb9D-H6CQgkbH/exec";
function uploadToDrive(fileName,dataUrl,mimeType,customerName,docType,cb){
  try{
    const monthFolder=new Date().toLocaleDateString("en-IN",{month:"long",year:"numeric"});
    fetch(DRIVE_UPLOAD_URL,{method:"POST",body:JSON.stringify({fileName,fileData:dataUrl,mimeType:mimeType||"image/jpeg",customerName,docType,monthFolder})})
      .then(r=>r.json())
      .then(d=>{if(d&&d.success&&d.id){cb("https://drive.google.com/uc?id="+d.id+"&export=view",d.url);}else{cb(dataUrl,null);}})
      .catch(()=>cb(dataUrl,null));
  }catch(e){cb(dataUrl,null);}
}
function compressImg(file,cb){
  try{
  var rd=new FileReader();
  rd.onload=function(e){
    var img=new Image();
    img.onload=function(){
      var MAX=900,w=img.width,h=img.height;
      if(w>MAX||h>MAX){if(w>h){h=Math.round(h*MAX/w);w=MAX;}else{w=Math.round(w*MAX/h);h=MAX;}}
      var cv=document.createElement("canvas");cv.width=w;cv.height=h;
      cv.getContext("2d").drawImage(img,0,0,w,h);
      cb(cv.toDataURL("image/jpeg",0.7));
    };
    img.onerror=function(){cb(e.target.result);};
    img.src=e.target.result;
  };
  rd.readAsDataURL(file);
  }catch(err){var r2=new FileReader();r2.onload=function(e){cb(e.target.result);};r2.readAsDataURL(file);}
}
function dlFile(content,filename,mime){
  try{
    var blob=new Blob([content],{type:mime||"text/html"});
    var url=URL.createObjectURL(blob);
    var a=document.createElement("a");a.href=url;a.download=filename;
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    setTimeout(function(){URL.revokeObjectURL(url);},2000);
  }catch(e){alert("Download blocked in this preview — will work in the deployed app/browser");}
}
function makeMRDoc(cust,b,c,pageOnly){
  const doc=new jsPDF({unit:"mm",format:"a4"});
  const W=210,pad=18;
  let y=18;
  function line(text,x,yy,size,style,align){doc.setFontSize(size||11);doc.setFont("helvetica",style||"normal");doc.text(text,x,yy,{align:align||"left"});}
  function hline(yy){doc.setDrawColor(200);doc.line(pad,yy,W-pad,yy);}
  function row(l,v,yy){line(l,pad,yy,10,"normal");line(v,W-pad,yy,10,"bold","right");}
  // Header
  line("NKD BAJAJ",W/2,y,18,"bold","center");y+=7;
  line("Authorised Bajaj Dealer | Dhanbad",W/2,y,9,"normal","center");y+=5;
  hline(y);y+=6;
  // Title
  line("MONEY RECEIPT",W/2,y,15,"bold","center");y+=8;
  hline(y);y+=6;
  // MR details row
  line("MR No: "+(b.mrNo||"—"),pad,y,10);line("Date: "+fd(cust.billedDate||td()),W-pad,y,10,"normal","right");y+=10;
  // Customer block
  const custRows=[["Customer Name",cust.name||""],["Father / Husband",cust.fatherName||""],["Address",cust.address||""],["Phone",cust.phone||""],["Aadhar",cust.aadhar||""],["PAN",cust.pan||""]];
  custRows.forEach(([l,v])=>{if(v){row(l,v,y);y+=7;}});
  y+=2;hline(y);y+=6;
  // Vehicle block
  line("VEHICLE DETAILS",pad,y,10,"bold");y+=7;
  const vehRows=[["Model",cust.model||""],["Model Code",cust.modelCode||""],["Chassis No",b.chassis||""],["Engine No",b.engine||""],["Colour",b.color||""],["Delivery Date",fd(b.deliveryDate)],["Reg No",b.registrationNo||""],["Finance By",b.financeBank||"Cash"]];
  vehRows.forEach(([l,v])=>{if(v&&v!=="—"){row(l,v,y);y+=7;}});
  y+=2;hline(y);y+=6;
  // Amount block
  line("AMOUNT DETAILS",pad,y,10,"bold");y+=7;
  const amtRows=[["On-Road Price (C)",fc(c.C)],["Consumer Offer",c.cof?"-"+fc(c.cof):""],["Special Discount",c.sdis?"-"+fc(c.sdis):""],["Corporate Scheme",c.corp?"-"+fc(c.corp):""],["DEAL PRICE (E)",fc(c.E)],["Booking Amount",c.bk?"-"+fc(c.bk):""],["Exchange Value",c.exv?"-"+fc(c.exv):""],["Loan / Disbursal",c.loan?"-"+fc(c.loan):""],["Balance from Customer",fc(c.I)]];
  amtRows.forEach(([l,v])=>{if(v&&v!=="₹0"&&v!==""){const isTot=l.startsWith("DEAL");if(isTot){doc.setFontSize(11);doc.setFont("helvetica","bold");}else{doc.setFontSize(10);doc.setFont("helvetica","normal");}doc.text(l,pad,y);doc.text(v,W-pad,y,{align:"right"});y+=isTot?8:6;}});
  const pmts=(b.payments&&b.payments.length?b.payments:[{mode:b.payMode||"",amt:c.paid,ref:""}]).filter(p=>Number(p.amt||0)>0);
  if(pmts.length>1){doc.setFontSize(10);doc.setFont("helvetica","bold");doc.text("Payment Received",pad,y);doc.text(fc(c.paid),W-pad,y,{align:"right"});y+=7;pmts.forEach(p=>{doc.setFontSize(9);doc.setFont("helvetica","normal");const lbl="  "+p.mode+(p.ref?" ["+p.ref+"]":"");doc.text(lbl,pad,y);doc.text(fc(Number(p.amt)),W-pad,y,{align:"right"});y+=6;});}else{doc.setFontSize(10);doc.setFont("helvetica","normal");const lbl="Amount Received"+(pmts[0]?(" ("+pmts[0].mode+(pmts[0].ref?" — "+pmts[0].ref:"")+")"):"")+(" (J)");doc.text(lbl,pad,y);doc.text(fc(c.paid),W-pad,y,{align:"right"});y+=6;}
  doc.setFontSize(11);doc.setFont("helvetica","bold");doc.text("BALANCE (K)",pad,y);doc.text(fc(Math.max(c.K,0)),W-pad,y,{align:"right"});y+=8;
  y+=4;hline(y);y+=10;
  // Signature
  line("Customer Signature",pad,y+16,9);line("Authorised Signatory",W-pad,y+16,9,"normal","right");
  line("______________________",pad,y+14,9);line("______________________",W-pad,y+14,9,"normal","right");
  line("NKD Bajaj, Dhanbad",W-pad,y+22,8,"italic","right");
  return doc;
}
function makeCalcDoc(cust,b,c){
  const doc=new jsPDF({unit:"mm",format:"a4"});
  const W=210,pad=18;
  let y=18;
  function line(text,x,yy,size,style,align){doc.setFontSize(size||11);doc.setFont("helvetica",style||"normal");doc.text(text,x,yy,{align:align||"left"});}
  function hline(yy){doc.setDrawColor(200);doc.line(pad,yy,W-pad,yy);}
  function row(l,v,yy,bold){doc.setFontSize(bold?11:10);doc.setFont("helvetica",bold?"bold":"normal");doc.text(l,pad,yy);doc.text(v,W-pad,yy,{align:"right"});}
  line("NKD BAJAJ",W/2,y,18,"bold","center");y+=7;
  line("CALCULATION SHEET (INTERNAL — OFFICE COPY)",W/2,y,9,"normal","center");y+=5;hline(y);y+=6;
  line("Customer: "+cust.name,pad,y,10,"bold");line("Date: "+fd(cust.billedDate||td()),W-pad,y,10,"normal","right");y+=7;
  line("Phone: "+cust.phone,pad,y,10);line("MR No: "+(b.mrNo||"—"),W-pad,y,10,"normal","right");y+=7;
  line("Model: "+(cust.model||"")+" ("+cust.modelCode+")",pad,y,10);y+=7;
  line("Chassis: "+(b.chassis||""),pad,y,10);line("Engine: "+(b.engine||""),W-pad,y,10,"normal","right");y+=7;
  line("Pay Mode: "+((b.payments&&b.payments.length?b.payments.filter(p=>Number(p.amt||0)>0).map(p=>p.mode+(p.ref?" ("+p.ref+")":"")).join(" + "):b.payMode)||""),pad,y,10);line("Financed By: "+(b.financeBank||"Cash"),W-pad,y,10,"normal","right");y+=7;
  line("Salesman: "+(cust.salesman||""),pad,y,10);line("Branch: "+(cust.branch||""),W-pad,y,10,"normal","right");y+=5;
  hline(y);y+=6;
  line("CALCULATION",pad,y,11,"bold");y+=8;
  const calcPmts=(b.payments&&b.payments.length?b.payments:[{mode:b.payMode||"",amt:c.paid,ref:""}]).filter(p=>Number(p.amt||0)>0);
  const payItems=calcPmts.length>1?[["Payment Received (J)",fc(c.paid),"bold"],...calcPmts.map(p=>["  "+p.mode+(p.ref?" ["+p.ref+"]":""),fc(Number(p.amt))])]:[[("Payment Received"+(calcPmts[0]?(" — "+calcPmts[0].mode+(calcPmts[0].ref?" ("+calcPmts[0].ref+")":"")):"")+" (J)"),fc(c.paid)]];
  const items=[["(A) Ex-Showroom",fc(c.ex)],["+ Comp. Accessories",fc(c.ca)],["+ Handling",fc(c.hdl)],["+ Insurance (5yr)",fc(c.ins)],["+ Registration",fc(c.reg)],["+ Accessories",fc(c.acc)],["+ Teflon",fc(c.tef)],["+ Hypothication",fc(c.hyp)],["+ AMC",fc(c.amcV)],["TOTAL ON-ROAD (C)",fc(c.C),"bold"],[""," "],["— Consumer Offer",c.cof?fc(c.cof):""],["— Special Discount",c.sdis?fc(c.sdis):""],["— Corporate",c.corp?fc(c.corp):""],["DEAL PRICE (E)",fc(c.E),"bold"],[""," "],["— Booking Amount",c.bk?fc(c.bk):""],["— Exchange Value",c.exv?fc(c.exv):""],["NET (G)",fc(c.G),"bold"],["— Loan / Disbursal",c.loan?fc(c.loan):""],["Balance from Customer (I)",fc(c.I),"bold"],...payItems,["DIFFERENCE (K)",fc(c.K),"bold"]];
  items.forEach(([l,v,style])=>{if(l===""){y+=3;return;}if(!v||v==="₹0")return;row(l,v,y,style==="bold");y+=style==="bold"?8:6;});
  y+=4;hline(y);y+=6;
  line("Approved By: "+(cust.approvedBy||"—"),pad,y,10);line("Enquiry Date: "+fd(cust.enquiryDate),W-pad,y,10,"normal","right");
  return doc;
}
function makeCombinedDoc(cust,b,c){
  const doc=makeCalcDoc(cust,b,c);
  doc.addPage();
  // re-draw MR on page 2
  const W=210,pad=18;let y=18;
  function line(text,x,yy,size,style,align){doc.setFontSize(size||11);doc.setFont("helvetica",style||"normal");doc.text(text,x,yy,{align:align||"left"});}
  function hline(yy){doc.setDrawColor(200);doc.line(pad,yy,W-pad,yy);}
  function row(l,v,yy){line(l,pad,yy,10,"normal");line(v,W-pad,yy,10,"bold","right");}
  line("NKD BAJAJ",W/2,y,18,"bold","center");y+=7;
  line("Authorised Bajaj Dealer | Dhanbad",W/2,y,9,"normal","center");y+=5;hline(y);y+=6;
  line("MONEY RECEIPT",W/2,y,15,"bold","center");y+=8;hline(y);y+=6;
  line("MR No: "+(b.mrNo||"—"),pad,y,10);line("Date: "+fd(cust.billedDate||td()),W-pad,y,10,"normal","right");y+=10;
  [["Customer",cust.name||""],["Father / Husband",cust.fatherName||""],["Phone",cust.phone||""],["Model",cust.model||""],["Chassis No",b.chassis||""],["Colour",b.color||""]].forEach(([l,v])=>{if(v){row(l,v,y);y+=7;}});
  y+=2;hline(y);y+=6;
  line("AMOUNT DETAILS",pad,y,10,"bold");y+=7;
  const cmbPmts=(b.payments&&b.payments.length?b.payments:[{mode:b.payMode||"",amt:c.paid,ref:""}]).filter(p=>Number(p.amt||0)>0);
  const cmbAmtRows=[["Deal Price",fc(c.E)],["Loan Amount",c.loan?fc(c.loan):""],["Balance from Customer",fc(c.I)]];
  cmbAmtRows.forEach(([l,v])=>{if(v&&v!=="₹0"){doc.setFontSize(10);doc.setFont("helvetica","normal");doc.text(l,pad,y);doc.text(v,W-pad,y,{align:"right"});y+=6;}});
  if(cmbPmts.length>1){doc.setFontSize(11);doc.setFont("helvetica","bold");doc.text("Amount Received",pad,y);doc.text(fc(c.paid),W-pad,y,{align:"right"});y+=7;cmbPmts.forEach(p=>{doc.setFontSize(9);doc.setFont("helvetica","normal");doc.text("  "+p.mode+(p.ref?" ["+p.ref+"]":""),pad,y);doc.text(fc(Number(p.amt)),W-pad,y,{align:"right"});y+=6;});}else{doc.setFontSize(11);doc.setFont("helvetica","bold");const lbl="Amount Received"+(cmbPmts[0]?(" ("+cmbPmts[0].mode+(cmbPmts[0].ref?" — "+cmbPmts[0].ref:"")+")"):"")+":";doc.text(lbl,pad,y);doc.text(fc(c.paid),W-pad,y,{align:"right"});y+=8;}
  doc.setFontSize(11);doc.setFont("helvetica","bold");doc.text("BALANCE",pad,y);doc.text(fc(Math.max(c.K,0)),W-pad,y,{align:"right"});y+=8;
  y+=6;hline(y);y+=10;
  line("Customer Signature",pad,y+16,9);line("Authorised Signatory",W-pad,y+16,9,"normal","right");
  line("______________________",pad,y+14,9);line("______________________",W-pad,y+14,9,"normal","right");
  return doc;
}
function makeBookingPdf(cust){
  const bk=cust.booking||{};
  const doc=new jsPDF({unit:"mm",format:"a4"});
  const W=210,pad=18;let y=18;
  function line(text,x,yy,size,style,align){doc.setFontSize(size||11);doc.setFont("helvetica",style||"normal");doc.text(text,x,yy,{align:align||"left"});}
  function hline(yy){doc.setDrawColor(200);doc.line(pad,yy,W-pad,yy);}
  function row(l,v,yy){line(l,pad,yy,10,"normal");line(v,W-pad,yy,10,"bold","right");}
  line("NKD BAJAJ",W/2,y,18,"bold","center");y+=7;
  line("Authorised Bajaj Dealer | Dhanbad",W/2,y,9,"normal","center");y+=5;
  hline(y);y+=6;
  line("BOOKING RECEIPT",W/2,y,15,"bold","center");y+=8;
  hline(y);y+=6;
  line("Date: "+fd(bk.date||td()),W-pad,y,10,"normal","right");y+=10;
  [["Customer Name",cust.name||""],["Father / Husband",cust.fatherName||""],["Address",cust.address||""],["Phone",cust.phone||""],["Model",cust.model||""],["Colour",cust.colour||""]].forEach(([l,v])=>{if(v){row(l,v,y);y+=7;}});
  y+=2;hline(y);y+=6;
  line("BOOKING DETAILS",pad,y,10,"bold");y+=8;
  [["Booking Amount",fc(bk.amt||0)],["Payment Mode",bk.mode||""],["Booking Date",fd(bk.date)],["Exchange Asked",cust.exchangeAsked||""]].forEach(([l,v])=>{if(v&&v!=="—"){const isTot=l==="Booking Amount";doc.setFontSize(isTot?12:10);doc.setFont("helvetica",isTot?"bold":"normal");doc.text(l,pad,y);doc.text(v,W-pad,y,{align:"right"});y+=isTot?9:7;}});
  if(bk.note){y+=2;line("Note: "+bk.note,pad,y,9,"italic");y+=6;}
  y+=4;hline(y);y+=8;
  line("This is a provisional booking receipt. Final billing subject to document verification.",pad,y,8,"italic");
  y+=12;
  line("Customer Signature",pad,y+16,9);line("Authorised Signatory",W-pad,y+16,9,"normal","right");
  line("______________________",pad,y+14,9);line("______________________",W-pad,y+14,9,"normal","right");
  line("NKD Bajaj, Dhanbad",W-pad,y+22,8,"italic","right");
  return doc;
}
function savePdfToDrive(doc,filename,customerName,docType){
  try{
    const now=new Date();
    const monthFolder=now.toLocaleDateString("en-IN",{month:"long",year:"numeric"});
    const base64=doc.output("datauristring");
    fetch(DRIVE_UPLOAD_URL,{method:"POST",body:JSON.stringify({fileName:filename,fileData:base64,mimeType:"application/pdf",customerName:(customerName||"Unknown").replace(/[<>:"/\\|?*]/g," ").trim()||"Unknown",docType:docType||"doc",monthFolder:monthFolder})}).catch(()=>{});
  }catch(e){}
}
async function sharePdf(doc,filename,phone,msg){
  const blob=doc.output("blob");
  const file=new File([blob],filename,{type:"application/pdf"});
  try{
    if(navigator.share&&navigator.canShare&&navigator.canShare({files:[file]})){
      await navigator.share({files:[file],title:"NKD Bajaj",text:msg||""});
      return;
    }
  }catch(e){if(e&&e.name==="AbortError")return;}
  // Fallback: download + open WhatsApp
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");a.href=url;a.download=filename;a.click();
  setTimeout(()=>URL.revokeObjectURL(url),3000);
  if(phone)setTimeout(()=>window.open("https://wa.me/91"+phone,"_blank"),800);
}
function sv(k,v){try{localStorage.setItem(k,JSON.stringify(v));return true;}catch(e){return false;}}
function ld(k,d){try{const v=localStorage.getItem(k);return v?JSON.parse(v):d;}catch(e){return d;}}

try{Object.assign(RC,ld("nkd_rc",{}));}catch(e){}

function calcB(f,r){
  const ex=Number(r?.ex||0),ca=Number(r?.cAcc||0),hdl=Number(f.hdl||r?.hdl||600);
  const ins=Number(f.ins||r?.ins||0),reg=Number(f.reg||r?.reg||0),acc=Number(f.acc||0);
  const tef=Number(f.tef||0),hyp=Number(f.hyp||0),amcV=f.addAmc?(Number(r?.amc)||0):0;
  const B=hdl+ins+reg+ca+acc+tef+hyp+amcV,C=ex+B;
  const cof=Number(f.cof||0),sdis=Number(f.sdis||0),corp=Number(f.corp||0),D=cof+sdis+corp,E=C-D;
  const bk=Number(f.bk||0),exv=Number(f.exv||0),F=bk+exv,G=E-F;
  const loan=Number(f.loan||0),I=G-loan;
  const paid=f.payments&&f.payments.length?(f.payments.reduce((s,p)=>s+Number(p.amt||0),0)):Number(f.paid||0);
  const K=I-paid;
  return{ex,ca,hdl,ins,reg,acc,tef,hyp,amcV,B,C,cof,sdis,corp,D,E,bk,exv,F,G,loan,I,paid,K};
}

function seedData(){
  const s=ld("nkd6",null);if(s)return s;
  const names=["Rajesh Verma","Sunil Gupta","Pooja Singh","Amit Sharma","Ravi Kumar","Meena Devi","Deepak Yadav","Sunita Prasad","Vijay Tiwari","Anita Kumari","Pankaj Jha","Kavita Singh","Rohit Sahu","Nisha Roy","Arun Mishra"];
  const codes=Object.keys(RC);
  const sts=["Hot","Warm","Hot","Cold","Warm","Hot","Warm","Cold","Hot","Warm","Hot","Booked","Warm","Hot","Cold"];
  return names.map((name,i)=>{
    const st=sts[i],mc=codes[i%codes.length],r=RC[mc],eq=aD(td(),-Math.floor(Math.random()*60+5));
    return{id:"C"+(1000+i),name,phone:"90"+String(1e7+i*1234567).slice(1),fatherName:"",address:"House "+(i+1)+", Main Road, Dhanbad",dob:"198"+i%9+"-0"+(i%9+1)+"-15",aadhar:"",pan:"",modelCode:mc,model:r.n,cat:r.cat,enquiryDate:eq,status:st,salesman:SM[i%SM.length],branch:SM_BRANCH[SM[i%SM.length]],finance:i%2===0?"Finance":"Cash",exchangeAsked:i%3===0?"Honda Shine 2019":"",exchangeOffered:i%3===0?"25000":"",remarks:"["+eq+"] INTERESTED: Visited showroom, enquired about "+r.n,followupDate:aD(eq,FU[st]||3),attempts:0,stopped:false,billed:false,billedDate:null,photos:{},billing:null,managerApproval:null,callLog:[]};
  });
}

const inp={background:"#f8fafc",border:"1px solid #6b8fb5",borderRadius:10,padding:"10px 12px",fontSize:13,color:"#1e293b",width:"100%",boxSizing:"border-box",outline:"none"};
const lbl={fontSize:11,color:"#475569",display:"block",marginBottom:4,fontWeight:600};
const btn=(bg,col="#fff")=>({background:bg,border:"none",borderRadius:10,padding:"10px 14px",color:col,fontWeight:700,cursor:"pointer",fontSize:13});

function PhIcon({s=18,c="#22c55e"}){return(<svg width={s} height={s} viewBox="0 0 24 24" fill="none" stroke={c} strokeWidth="2.2" strokeLinecap="round" strokeLinejoin="round"><path d="M22 16.92v3a2 2 0 01-2.18 2 19.79 19.79 0 01-8.63-3.07A19.5 19.5 0 013.07 11.5a19.79 19.79 0 01-3.07-8.63A2 2 0 012 .82h3a2 2 0 012 1.72c.128.96.347 1.903.7 2.81a2 2 0 01-.45 2.11L6.09 8.91a16 16 0 006 6l1.27-1.27a2 2 0 012.11-.45c.907.353 1.85.572 2.81.7A2 2 0 0122 15.92z"/></svg>);}
function WAIcon({s=18}){return(<svg width={s} height={s} viewBox="0 0 24 24" fill="#25D366"><path d="M17.472 14.382c-.297-.149-1.758-.867-2.03-.967-.273-.099-.471-.148-.67.15-.197.297-.767.966-.94 1.164-.173.199-.347.223-.644.075-.297-.15-1.255-.463-2.39-1.475-.883-.788-1.48-1.761-1.653-2.059-.173-.297-.018-.458.13-.606.134-.133.298-.347.446-.52.149-.174.198-.298.298-.497.099-.198.05-.371-.025-.52-.075-.149-.669-1.612-.916-2.207-.242-.579-.487-.5-.669-.51-.173-.008-.371-.01-.57-.01-.198 0-.52.074-.792.372-.272.297-1.04 1.016-1.04 2.479 0 1.462 1.065 2.875 1.213 3.074.149.198 2.096 3.2 5.077 4.487.709.306 1.262.489 1.694.625.712.227 1.36.195 1.871.118.571-.085 1.758-.719 2.006-1.413.248-.694.248-1.289.173-1.413-.074-.124-.272-.198-.57-.347m-5.421 7.403h-.004a9.87 9.87 0 01-5.031-1.378l-.361-.214-3.741.982.998-3.648-.235-.374a9.86 9.86 0 01-1.51-5.26c.001-5.45 4.436-9.884 9.888-9.884 2.64 0 5.122 1.03 6.988 2.898a9.825 9.825 0 012.893 6.994c-.003 5.45-4.437 9.884-9.885 9.884m8.413-18.297A11.815 11.815 0 0012.05 0C5.495 0 .16 5.335.157 11.892c0 2.096.547 4.142 1.588 5.945L.057 24l6.305-1.654a11.882 11.882 0 005.683 1.448h.005c6.554 0 11.89-5.335 11.893-11.893a11.821 11.821 0 00-3.48-8.413z"/></svg>);}

function Badge({s}){return(<span style={{fontSize:10,fontWeight:800,background:ST_C[s]||"#374151",color:"#fff",padding:"2px 8px",borderRadius:20}}>{(s||"").toUpperCase()}</span>);}
function Card({c,onClick,showSM}){
  const r=RC[c.modelCode];
  const ovd=!c.billed&&!c.stopped&&c.followupDate<td();
  const last=(c.remarks||"").trim().split("\n").filter(Boolean).slice(-1)[0]||"";
  const lastClean=last.replace(/^\[.*?\]\s*/,"").replace(/^[A-Z_]+:\s*/,"");
  return(
    <div onClick={onClick} className="fu glass" style={{background:"#ffffff",border:"1px solid "+(ovd?"rgba(239,68,68,0.5)":"#6b8fb5"),borderRadius:13,padding:"12px 14px",marginBottom:9,cursor:"pointer",borderLeft:"3px solid "+(ST_C[c.status]||"#374151")}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
        <div style={{flex:1,minWidth:0}}>
          <div style={{display:"flex",alignItems:"center",gap:7,flexWrap:"wrap"}}>
            <span style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>{c.name}</span>
            <Badge s={c.status}/>
            {c.stopped&&<span style={{fontSize:9,color:"#ef4444",fontWeight:800,background:"rgba(239,68,68,0.15)",padding:"2px 6px",borderRadius:6}}>STOPPED</span>}{c.managerApproval==="rejected"&&!c.billed&&<span style={{fontSize:9,color:"#fff",fontWeight:800,background:"#ef4444",padding:"2px 6px",borderRadius:6}}>❌ REJECTED — FIX CALC</span>}
          </div>
          <div style={{fontSize:12,color:"#64748b",marginTop:2}}>{c.phone} · {c.model}</div>
          <div style={{fontSize:11,color:"#94a3b8"}}>{c.modelCode} · {c.finance}{showSM&&<span> · <span style={{color:"#60a5fa",fontWeight:700}}>{c.salesman}</span></span>}</div>
          {lastClean&&<div style={{marginTop:6,borderLeft:"2px solid rgba(249,115,22,0.5)",paddingLeft:7,fontSize:11,color:"#64748b",lineHeight:1.4}}>{lastClean.length>80?lastClean.slice(0,80)+"…":lastClean}</div>}
        </div>
        <div style={{textAlign:"right",marginLeft:10,flexShrink:0}}>
          {!c.billed&&!c.stopped&&<div style={{fontSize:10,color:ovd?"#ef4444":"#5a6478",fontWeight:ovd?700:400}}>{ovd?"OVERDUE":"DUE"}<br/><span style={{fontSize:11}}>{fd(c.followupDate)}</span></div>}
          {c.billed&&(()=>{const K=c.billing&&c.billing.calc?c.billing.calc.K:null;return(<div style={{textAlign:"right"}}><div style={{fontSize:11,color:"#34d399",fontWeight:700}}>✓ BILLED · {fd(c.billedDate)}</div>{K!==null&&K!==0&&<div style={{fontSize:12,fontWeight:800,color:"#ef4444",marginTop:2}}>⚠️ Due: {fc(K)}</div>}{K===0&&<div style={{fontSize:11,fontWeight:700,color:"#22c55e",marginTop:2}}>✅ Fully Paid</div>}</div>);})()}
        </div>
      </div>
      {r&&<div style={{marginTop:6,fontSize:11,color:"#64748b"}}>On-Road: <span style={{color:"#475569",fontWeight:600}}>{fc(r.onRoad)}</span></div>}
    </div>
  );
}

function Login({onLogin,nkdUsers}){
  const users=nkdUsers||DEFAULT_USERS;
  const [role,setRole]=useState("salesman");
  const [user,setUser]=useState(SM[0]);
  const [uname,setUname]=useState("");
  const [pin,setPin]=useState("");
  const [br,setBr]=useState(BRANCHES[0]);
  const [spw,setSpw]=useState("");
  const [showPw,setShowPw]=useState(false);
  const [chg,setChg]=useState(false);
  const [npw1,setNpw1]=useState("");
  const [npw2,setNpw2]=useState("");
  function go(){
    if(role==="salesman"){
      const pws=ld("nkd_pw",{});
      const rec=pws[user]||{pw:"1111",must:true,fails:0,locked:false};
      if(rec.locked){alert("🔒 Account locked after 5 wrong attempts.\nAsk Admin/Owner to reset your password.");return;}
      if(spw!==rec.pw){rec.fails=(rec.fails||0)+1;if(rec.fails>=5){rec.locked=true;alert("🔒 Account LOCKED (5 wrong attempts). Admin has been flagged.");}else{alert("Wrong password ("+rec.fails+"/5 attempts)");}pws[user]=rec;sv("nkd_pw",pws);_dbSet("passwords",pws);return;}
      rec.fails=0;
      if(rec.must&&!chg){setChg(true);pws[user]=rec;sv("nkd_pw",pws);_dbSet("passwords",pws);return;}
      if(rec.must&&chg){
        if(!npw1||npw1.length<4){alert("New password must be at least 4 characters");return;}
        if(npw1!==npw2){alert("New passwords do not match");return;}
        rec.pw=npw1;rec.must=false;
      }
      pws[user]=rec;sv("nkd_pw",pws);_dbSet("passwords",pws);
      onLogin(role,user,null);return;
    }
    // manager / owner / admin — check against nkdUsers
    const roleList=users[role]||[];
    const match=roleList.find(u=>u.name.trim().toLowerCase()===uname.trim().toLowerCase()&&u.pin===pin);
    if(!match){alert("Wrong username or PIN");return;}
    onLogin(role,match.name,(role==="manager")?br:null);
  }
  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,#f0f7ff 0%,#f8fafc 100%)",display:"flex",alignItems:"center",justifyContent:"center",padding:16}}>
      <div className="fu glass" style={{width:"100%",maxWidth:360,background:"#ffffff",borderRadius:26,padding:32,border:"1px solid #6b8fb5",boxShadow:"0 8px 40px rgba(15,23,42,.12)"}}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{width:120,borderRadius:16,background:"#fff",padding:"8px 12px",margin:"0 auto 14px",display:"flex",alignItems:"center",justifyContent:"center",animation:"glow 3s ease infinite"}}><img src="/logo.png" alt="NKD Bajaj" style={{width:"100%",display:"block"}}/></div>
          <div style={{fontWeight:800,fontSize:20,color:"#1e293b"}}>NKD Bajaj CRM</div>
          <div style={{color:"#94a3b8",fontSize:12,marginTop:3}}>Dhanbad · 3 Showrooms</div>
        </div>
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div><label style={lbl}>Role</label><select style={inp} value={role} onChange={e=>{setRole(e.target.value);setUname("");setPin("");}}>
            <option value="salesman">Sales Executive</option>
            <option value="manager">Manager</option>
            <option value="owner">Owner</option>
            <option value="admin">Admin (Documents)</option>
            <option value="tech">Tech</option>
          </select></div>
          {/* ── SALESMAN ── */}
          {role==="salesman"&&<div><label style={lbl}>Your Name</label><select style={inp} value={user} onChange={e=>setUser(e.target.value)}>{SM.map(s=><option key={s}>{s}</option>)}</select></div>}
          {role==="salesman"&&<div><label style={lbl}>Password <span style={{color:"#94a3b8",fontWeight:400}}>(first time: 1111)</span></label>
            <div style={{position:"relative"}}>
              <input type={showPw?"text":"password"} style={inp} value={spw} onChange={e=>setSpw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}/>
              <button onClick={()=>setShowPw(!showPw)} style={{position:"absolute",right:8,top:9,background:"transparent",border:"none",color:"#94a3b8",cursor:"pointer",fontSize:14}}>{showPw?"🙈":"👁️"}</button>
            </div>
          </div>}
          {role==="salesman"&&chg&&<div style={{background:"rgba(249,115,22,0.08)",border:"1px solid rgba(249,115,22,0.35)",borderRadius:12,padding:12}}>
            <div style={{fontSize:11,color:"#f97316",fontWeight:700,marginBottom:8}}>🔐 First login — set your own password</div>
            <input type="password" placeholder="New password (min 4 chars)" style={{...inp,marginBottom:8}} value={npw1} onChange={e=>setNpw1(e.target.value)}/>
            <input type="password" placeholder="Re-enter new password" style={inp} value={npw2} onChange={e=>setNpw2(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()}/>
          </div>}
          {/* ── MANAGER / OWNER / ADMIN ── */}
          {role!=="salesman"&&<div><label style={lbl}>Username</label><input style={inp} value={uname} onChange={e=>setUname(e.target.value)} placeholder={"Your "+role+" username"} autoComplete="off"/></div>}
          {role==="manager"&&<div><label style={lbl}>Your Branch</label><select style={inp} value={br} onChange={e=>setBr(e.target.value)}>{BRANCHES.map(b=><option key={b}>{b}</option>)}</select></div>}
          {role!=="salesman"&&<div><label style={lbl}>PIN</label><input type="password" style={inp} value={pin} onChange={e=>setPin(e.target.value)} onKeyDown={e=>e.key==="Enter"&&go()} placeholder="Enter PIN"/></div>}
          {role!=="salesman"&&(
            <div style={{background:"rgba(249,115,22,0.07)",border:"1px solid rgba(249,115,22,0.25)",borderRadius:10,padding:"9px 12px",fontSize:11,color:"#94a3b8"}}>
              <div style={{fontWeight:700,color:"#f97316",marginBottom:5}}>🔑 Default credentials (first time)</div>
              {[["Manager","Manager","1234"],["Owner","Owner","0000"],["Admin","Admin","9999"],["Tech","Tech","1111"]].map(([r,u,p])=>(
                <div key={r} style={{display:"flex",gap:6,marginBottom:2}}>
                  <span style={{color:"#64748b",minWidth:60}}>{r}:</span>
                  <span style={{color:"#1e293b",fontWeight:600}}>Username: <b>{u}</b> · PIN: <b>{p}</b></span>
                </div>
              ))}
              <div style={{marginTop:6,color:"#f97316",fontWeight:600}}>Change these from Owner / Tech portal → 👤 User Accounts</div>
            </div>
          )}
          <button onClick={go} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),padding:14,fontSize:15,borderRadius:13,marginTop:4}}>Login →</button>
        </div>
      </div>
    </div>
  );
}

function DuePayRow({c,K,role,onOpen,onAddPayment}){
  const [open,setOpen]=useState(false);
  const [mode,setMode]=useState("Cash");
  const [amt,setAmt]=useState("");
  const [ref,setRef]=useState("");
  function save(){
    const a=Number(amt);
    if(!a||a<=0){alert("Enter a valid amount");return;}
    if(a>K){alert("Amount ₹"+a+" exceeds balance due ₹"+K);return;}
    onAddPayment(c.id,{mode,amt:a,ref,date:td()});
    setOpen(false);setAmt("");setRef("");
  }
  return(
    <div style={{background:"rgba(239,68,68,0.05)",border:"1px solid rgba(239,68,68,0.4)",borderRadius:12,padding:"11px 13px",marginBottom:7}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div onClick={()=>onOpen(c,"billing")} style={{flex:1,minWidth:0,cursor:"pointer"}}>
          <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{c.name}</div>
          <div style={{fontSize:11,color:"#64748b"}}>{c.model} · Billed {fd(c.billedDate)}{role!=="salesman"?" · "+c.salesman:""}</div>
          {c.billing?.mrNo&&<div style={{fontSize:10,color:"#8b5cf6",fontWeight:700,marginTop:2}}>MR# {c.billing.mrNo}</div>}
        </div>
        <div style={{textAlign:"right",flexShrink:0,marginLeft:10}}>
          <div style={{fontWeight:900,fontSize:15,color:"#ef4444"}}>{fc(K)}</div>
          <div style={{fontSize:9,color:"#94a3b8"}}>DUE</div>
        </div>
      </div>
      {!open&&<button onClick={()=>setOpen(true)} style={{marginTop:8,width:"100%",background:"linear-gradient(135deg,#22c55e,#16a34a)",border:"none",borderRadius:9,padding:"9px",fontSize:12,color:"#fff",fontWeight:700,cursor:"pointer",boxShadow:"0 2px 8px rgba(34,197,94,0.3)"}}>💳 Receive Payment</button>}
      {open&&(
        <div style={{marginTop:8,background:"#f8fafc",border:"1px solid #6b8fb5",borderRadius:10,padding:"10px 12px"}}>
          <div style={{fontSize:11,fontWeight:700,color:"#1e293b",marginBottom:8}}>Record Payment — Balance: {fc(K)}</div>
          <div style={{display:"flex",gap:6,marginBottom:7}}>
            <select value={mode} onChange={e=>setMode(e.target.value)} style={{...inp,flex:"0 0 110px",padding:"7px 8px",fontSize:12}}>
              {["Cash","Cheque","UPI","RTGS","Finance"].map(m=><option key={m}>{m}</option>)}
            </select>
            <input type="number" value={amt} onChange={e=>setAmt(e.target.value)} placeholder={"Amount (max "+fc(K)+")"} style={{...inp,flex:1,padding:"7px 8px",fontSize:12}}/>
          </div>
          <input value={ref} onChange={e=>setRef(e.target.value)} placeholder="Cheque no / Ref (optional)" style={{...inp,marginBottom:8,padding:"7px 8px",fontSize:12}}/>
          <div style={{display:"flex",gap:6}}>
            <button onClick={save} style={{...btn("linear-gradient(135deg,#22c55e,#16a34a)"),flex:1,padding:"8px",borderRadius:9,fontSize:12}}>✅ Save Receipt</button>
            <button onClick={()=>{setOpen(false);setAmt("");setRef("");}} style={{flex:"0 0 70px",background:"transparent",border:"1px solid #6b8fb5",borderRadius:9,padding:"8px",fontSize:12,color:"#64748b",cursor:"pointer"}}>Cancel</button>
          </div>
        </div>
      )}
    </div>
  );
}
function ExchDashGroup({exchName,list,onUpd,notify}){
  const [open,setOpen]=useState(false);
  const [edits,setEdits]=useState({});
  function getExv(c){return Number(c.billing?.exv||c.billing?.calc?.exv||0);}
  function setE(id,field,val){setEdits(p=>({...p,[id]:{...(p[id]||{}),[ field]:val}}));}
  function getE(c,field){return edits[c.id]?.[field]??c[field]??"";}
  function getPhone(c){return c.billing?.details?.exchangePhone||c.exchangePhone||"";}
  function receiveAndSend(){
    const rows=list.map(c=>({name:c.name,model:c.model||"",regNo:c.billing?.details?.exchangeRegNo||c.exchangeRegNo||"",exv:getExv(c),amtRec:Number(getE(c,"exchAmtRec")||0),comm:Number(getE(c,"exchComm")||0),disc:Number(getE(c,"exchDisc")||0)}));
    const doc=makeExchMRDoc(exchName,rows,td());
    const fname="ExchMR_"+exchName.replace(/ /g,"_")+"_"+td()+".pdf";
    const offNum=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;
    const exchPhone=list.map(c=>getPhone(c)).find(p=>p)||"";
    const msg="Exchange Settlement MR for "+exchName+" — "+list.length+" vehicle(s)";
    if(exchPhone)sharePdf(doc,fname,exchPhone,msg);
    sharePdf(doc,fname,offNum,msg);
    savePdfToDrive(doc,fname,exchName,"ExchMR");
    list.forEach(c=>{savePdfToDrive(doc,fname,c.name,"ExchMR");onUpd(c.id,{exchAmtRec:Number(getE(c,"exchAmtRec")||0),exchComm:Number(getE(c,"exchComm")||0),exchDisc:Number(getE(c,"exchDisc")||0),exchMrIssued:true,exchMrDate:td()});});
    notify("✅ MR sent to "+(exchPhone?"exchanger & ":"")+"office for "+exchName);
  }
  return(
    <div style={{background:"rgba(245,158,11,0.06)",border:"1px solid rgba(245,158,11,0.4)",borderRadius:12,padding:"11px 13px",marginBottom:7}}>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}} onClick={()=>setOpen(p=>!p)}>
        <div>
          <span style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{exchName}</span>
          <span style={{fontSize:10,color:"#f59e0b",fontWeight:800,marginLeft:6}}>{list.length} vehicle{list.length>1?"s":""}</span>
        </div>
        <span style={{fontSize:12,color:"#f59e0b"}}>{open?"▲":"▼"}</span>
      </div>
      {!open&&list.map(c=><div key={c.id} style={{fontSize:11,color:"#64748b",marginTop:3}}>• {c.name} — {c.model} · EXV: {fc(getExv(c))}{c.billing?.mrNo?" · MR# "+c.billing.mrNo:""}</div>)}
      {open&&(
        <div style={{marginTop:10}}>
          {list.map(c=>(
            <div key={c.id} style={{background:"#fff",border:"1px solid #e2e8f0",borderRadius:10,padding:"10px 12px",marginBottom:8}}>
              <div style={{fontWeight:700,fontSize:12,color:"#1e293b",marginBottom:6}}>{c.name} · {c.model}</div>
              <div style={{fontSize:11,color:"#64748b",marginBottom:6}}>EXV: {fc(getExv(c))}{c.billing?.mrNo?" · MR# "+c.billing.mrNo:""}{getPhone(c)?" · 📞 "+getPhone(c):""}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                {[["exchAmtRec","Amt Recd"],["exchComm","Commission"]].map(([field,label])=>(
                  <div key={field}>
                    <div style={{fontSize:9,color:"#64748b",marginBottom:2}}>{label}</div>
                    <input type="number" value={getE(c,field)} onChange={e=>setE(c.id,field,e.target.value)} placeholder="₹0" style={{width:"100%",border:"1px solid #cbd5e1",borderRadius:7,padding:"6px 8px",fontSize:12,boxSizing:"border-box"}}/>
                  </div>
                ))}
              </div>
            </div>
          ))}
          <button onClick={receiveAndSend} style={{width:"100%",background:"linear-gradient(135deg,#f59e0b,#d97706)",border:"none",borderRadius:9,padding:"10px",fontSize:12,color:"#fff",fontWeight:700,cursor:"pointer",boxShadow:"0 2px 8px rgba(245,158,11,0.35)"}}>📲 Issue MR & Send to Exchanger + Office</button>
        </div>
      )}
    </div>
  );
}
function Dashboard({custs,role,onOpen,onNav,onNavF,onSvcDone,onTeamTap,onAddPayment,onUpd,notify}){
  const hot=custs.filter(c=>c.status==="Hot"&&!c.billed);
  const stats=[
    {l:"Hot",st:"Hot",v:custs.filter(c=>c.status==="Hot"&&!c.billed).length,c:"#ef4444"},
    {l:"Warm",st:"Warm",v:custs.filter(c=>c.status==="Warm"&&!c.billed).length,c:"#f97316"},
    {l:"Cold",st:"Cold",v:custs.filter(c=>c.status==="Cold"&&!c.billed).length,c:"#3b82f6"},
    {l:"Booked",st:"Booked",v:custs.filter(c=>c.status==="Booked"&&!c.billed).length,c:"#8b5cf6"},
    {l:"Billed",st:"Billed",v:custs.filter(c=>c.billed).length,c:"#10b981"},
    {l:"MTD",v:custs.filter(c=>c.billed&&(c.billedDate||"").startsWith(new Date().toISOString().slice(0,7))).length,c:"#34d399"},
  ];
  return(
    <div>
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:16}}>
        <div><div style={{fontWeight:800,fontSize:19,color:"#1e293b"}}>Dashboard</div><div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"long"})}</div></div>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:18}}>
        {stats.map(s=>(
          <div key={s.l} onClick={()=>onNavF(s.st||"All")} className="glass" style={{background:"#fff",border:"2px solid "+s.c+"30",borderRadius:16,padding:"18px 8px",cursor:"pointer",textAlign:"center",boxShadow:"0 2px 12px rgba(15,23,42,.06)"}}>
            <div style={{fontSize:26,fontWeight:900,color:s.c,lineHeight:1}}>{s.v}</div>
            <div style={{fontSize:10,color:"#64748b",marginTop:3}}>{s.l}</div>
          </div>
        ))}
      </div>
      {(()=>{
        const withK=custs.filter(c=>c.billed&&c.billing).map(c=>{
          const r=RC[c.modelCode]||{};
          const stored=c.billing.calc?Number(c.billing.calc.K):NaN;
          const K=(!isNaN(stored)&&stored>0)?stored:Number(calcB(c.billing,r).K)||0;
          return{...c,_K:K};
        }).filter(c=>c._K>0).sort((a,b)=>b._K-a._K);
        return withK.length>0?(
          <div style={{marginBottom:18}}>
            <div style={{fontSize:12,fontWeight:800,color:"#ef4444",marginBottom:8}}>⚠️ BALANCE DUE — BILLED CUSTOMERS ({withK.length})</div>
            {withK.map(c=><DuePayRow key={c.id} c={c} K={c._K} role={role} onOpen={onOpen} onAddPayment={onAddPayment}/>)}
            <div style={{fontSize:10,color:"#94a3b8",textAlign:"right",marginTop:4}}>Total pending: {fc(withK.reduce((s,c)=>s+c._K,0))}</div>
          </div>
        ):null;
      })()}
      {(()=>{const svc=custs.filter(c=>{if(!c.billed||c.serviceDone)return false;const d=(new Date(aD(c.billedDate,45))-new Date())/864e5;return d<=7&&d>=-30;});return svc.length>0?(
        <div style={{marginBottom:16}}>
          <div style={{fontSize:12,fontWeight:800,color:"#60a5fa",marginBottom:8}}>🔧 1st FREE SERVICE DUE ({svc.length})</div>
          {svc.map(c=><div key={c.id} style={{background:"rgba(96,165,250,0.07)",border:"1px solid rgba(96,165,250,0.3)",borderRadius:12,padding:"11px 13px",marginBottom:7}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div onClick={()=>onOpen(c)} style={{cursor:"pointer",flex:1}}>
                <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{c.name}</div>
                <div style={{fontSize:11,color:"#64748b"}}>{c.model} · Due {fd(aD(c.billedDate,45))}</div>
              </div>
              <div style={{display:"flex",gap:6}}>
                <a href={"https://wa.me/91"+c.phone+"?text="+encodeURIComponent("Hello "+c.name+", NKD Bajaj here! Your "+(c.model||"bike")+" is due for its 1st FREE service. Please visit our workshop at Hirak Road, Dhanbad. Thank you!")} target="_blank" rel="noreferrer" style={{background:"rgba(37,211,102,0.1)",border:"1px solid rgba(37,211,102,0.3)",borderRadius:8,padding:"6px 9px",fontSize:10,color:"#25D366",fontWeight:700,textDecoration:"none"}}>📲 Remind</a>
                <button onClick={()=>onSvcDone(c.id)} style={{background:"#c2d6ec",border:"1px solid #6b8fb5",borderRadius:8,padding:"6px 9px",fontSize:10,color:"#64748b",fontWeight:700,cursor:"pointer"}}>✓ Done</button>
              </div>
            </div>
          </div>)}
        </div>
      ):null;})()}
      {(()=>{const rej=custs.filter(c=>c.managerApproval==="rejected"&&!c.billed);return rej.length>0?(
        <div style={{marginBottom:16}}>
          <div style={{fontSize:12,fontWeight:800,color:"#ef4444",marginBottom:8}}>❌ REJECTED BILLINGS — FIX &amp; RE-BILL ({rej.length})</div>
          {rej.map(c=><div key={c.id} onClick={()=>onOpen(c)} style={{background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.4)",borderRadius:12,padding:"11px 13px",marginBottom:7,cursor:"pointer"}}>
            <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{c.name} <span style={{fontSize:10,color:"#ef4444",fontWeight:800}}>· REJECTED</span></div>
            <div style={{fontSize:11,color:"#64748b"}}>{c.model} · {role!=="salesman"?c.salesman+" · ":""}Correct calc sheet &amp; bill again</div>
          </div>)}
        </div>
      ):null;})()}
      {(()=>{
        const exchPending=custs.filter(c=>c.billed&&c.billing&&Number(c.billing.exv||c.billing.calc?.exv||0)>0&&!c.exchMrIssued);
        if(exchPending.length===0)return null;
        const byExch={};
        exchPending.forEach(c=>{const n=c.billing?.details?.exchangeName||c.exchangeName||"Unknown";if(!byExch[n])byExch[n]=[];byExch[n].push(c);});
        const totalExv=exchPending.reduce((s,c)=>s+Number(c.billing?.exv||c.billing?.calc?.exv||0),0);
        return(
          <div style={{marginBottom:18}}>
            <div style={{fontSize:12,fontWeight:800,color:"#f59e0b",marginBottom:8}}>🔄 EXCHANGER DUE — PENDING SETTLEMENT ({exchPending.length})</div>
            {Object.entries(byExch).map(([name,list])=>(
              <ExchDashGroup key={name} exchName={name} list={list} onUpd={onUpd} notify={notify}/>
            ))}
            <div style={{fontSize:10,color:"#94a3b8",textAlign:"right",marginTop:4}}>Total exch value pending: {fc(totalExv)}</div>
          </div>
        );
      })()}
      <div style={{fontSize:12,fontWeight:700,color:"#ef4444",marginBottom:8}}>🔥 HOT LEADS</div>
      {hot.length===0&&<div style={{color:"#64748b",fontSize:13,textAlign:"center",padding:24}}>No hot leads</div>}
      {hot.slice(0,4).map(c=><Card key={c.id} c={c} onClick={()=>onOpen(c)} showSM={role!=="salesman"}/>)}
      {role!=="salesman"&&(
        <div style={{marginTop:16}}>
          <div style={{fontSize:12,fontWeight:700,color:"#64748b",marginBottom:8}}>TEAM</div>
          {SM.map(s=>{
            const m=custs.filter(c=>c.salesman===s);
            const b=m.filter(c=>c.billed).length;
            const due=m.filter(c=>!c.billed&&!c.stopped&&c.followupDate<=td()).length;
            return(
              <div key={s} onClick={()=>onTeamTap(s)} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:"13px 14px",marginBottom:7,display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}}>
                <div><div style={{fontWeight:600,fontSize:13,color:"#1e293b"}}>{s}</div><div style={{fontSize:11,color:"#94a3b8"}}>{m.length} leads{due>0?" · ":""}{due>0&&<span style={{color:"#f97316",fontWeight:700}}>{due} due</span>}</div></div>
                <div style={{textAlign:"right"}}><div style={{fontSize:20,fontWeight:900,color:"#34d399"}}>{b}</div><div style={{fontSize:9,color:"#94a3b8"}}>BILLED</div></div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

function Followups({items,onOpen,onLog,onCallLog,showSMFilter,initSM}){
  const [smF,setSmF]=useState(initSM||"All");
  useEffect(()=>{if(initSM)setSmF(initSM);},[initSM]);
  items=smF==="All"?items:items.filter(c=>c.salesman===smF);
  const [active,setActive]=useState(null);
  const [form,setForm]=useState({out:"interested",note:"",nxt:"",reason:"",competitor:""});
  const [timer,setTimer]=useState(null);
  const [elapsed,setElapsed]=useState(0);
  const iv=useRef(null);
  useEffect(()=>()=>clearInterval(iv.current),[]);
  function startCall(c){
    setTimer(c);setElapsed(0);clearInterval(iv.current);
    iv.current=setInterval(()=>setElapsed(p=>p+1),1000);
    
  }
  function endCall(){
    clearInterval(iv.current);
    const dur=String(Math.floor(elapsed/60)).padStart(2,"0")+":"+String(elapsed%60).padStart(2,"0");
    onCallLog(timer,dur);setTimer(null);setElapsed(0);
  }
  function submit(c){
    if(form.out==="not_interested"&&!form.reason){alert("Select reason for losing customer");return;}
    onLog(c,form.out,form.note,form.nxt||null,{reason:form.reason,competitor:form.competitor});
    setActive(null);setForm({out:"interested",note:"",nxt:"",reason:"",competitor:""});
  }
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:3}}>Today's Followups</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:10}}>🔥 Hot first — auto prioritised</div>
      {showSMFilter&&<div style={{display:"flex",gap:5,overflowX:"auto",marginBottom:12,paddingBottom:2}}>
        {["All",...SM].map(s=><button key={s} onClick={()=>setSmF(s)} style={{padding:"5px 11px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",flexShrink:0,background:smF===s?"#f97316":"#6b8fb5",color:smF===s?"#fff":"#8892a4",border:"none"}}>{s==="All"?"All Team":s.split(" ")[0]}</button>)}
      </div>}
      {timer&&(
        <div style={{background:"rgba(34,197,94,0.1)",border:"1px solid rgba(34,197,94,0.4)",borderRadius:13,padding:"14px 16px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{color:"#22c55e",fontWeight:700,fontSize:13}}>📞 Calling — {timer.name}</div><div style={{fontFamily:"monospace",fontSize:24,fontWeight:900,color:"#1e293b"}}>{String(Math.floor(elapsed/60)).padStart(2,"0")}:{String(elapsed%60).padStart(2,"0")}</div></div>
          <button onClick={endCall} style={{...btn("#ef4444"),padding:"10px 18px"}}>End</button>
        </div>
      )}
      {items.length===0&&<div className="fu" style={{textAlign:"center",padding:"40px 20px",background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:16}}><div style={{fontSize:40,marginBottom:8}}>🏆</div><div style={{fontWeight:800,fontSize:15,color:"#1e293b"}}>All caught up!</div><div style={{fontSize:12,color:"#94a3b8",marginTop:4}}>No followups due right now. Great work.</div></div>}
      {items.map(c=>{
        const last=(c.remarks||"").trim().split("\n").filter(Boolean).slice(-1)[0]||"";
        return(
          <div key={c.id} className="fu glass" style={{background:"#ffffff",border:"1px solid "+(ST_C[c.status]||"#6b8fb5")+"50",borderRadius:18,marginBottom:14,overflow:"hidden"}}>
            <div style={{padding:"12px 14px",cursor:"pointer"}} onClick={()=>onOpen(c)}>
              <div style={{display:"flex",alignItems:"center",gap:7}}><span style={{fontWeight:700,fontSize:15,color:"#1e293b"}}>{c.name}</span><Badge s={c.status}/></div>
              <div style={{fontSize:12,color:"#64748b",marginTop:2}}>{c.phone} · {c.model}</div>
              <div style={{fontSize:11,color:"#94a3b8"}}>Attempt #{(c.attempts||0)+1}/3 · <span style={{color:"#60a5fa",fontWeight:700}}>{c.salesman}</span></div>
              {(()=>{const tip={Hot:"🔥 Push for booking today — offer a test ride",Warm:"💬 Send offer on WhatsApp + invite for visit",Cold:"📞 Light check-in — share price/festival offer",Booked:"📄 Complete finance & docs — plan delivery"}[c.status];
                return tip?<div style={{marginTop:6,fontSize:11,color:"#60a5fa",fontWeight:600}}>{tip}</div>:null;})()}
              {last&&<div style={{marginTop:7,background:"rgba(249,115,22,0.07)",border:"1px solid rgba(249,115,22,0.2)",borderRadius:8,padding:"7px 10px"}}>
                <div style={{fontSize:9,color:"#f97316",fontWeight:700,marginBottom:2}}>LAST REMARK</div>
                <div style={{fontSize:12,color:"#334155",lineHeight:1.4}}>{last.replace(/^\[.*?\]\s*/,"").replace(/^[A-Z_]+:\s*/,"")}</div>
              </div>}
            </div>
            <div style={{padding:"6px 14px 2px",display:"flex",gap:5,flexWrap:"wrap"}}>
              {[["interested","👍 Interested"],["no_response","📵 No Resp"],["switched_off","🔕 Off"],["booked","✅ Booked"],["price_issue","💸 Price"],["visit_scheduled","🏪 Visit"]].map(([o,l])=>(
                <button key={o} onClick={()=>onLog(c,o,"",null,{})} style={{background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:16,padding:"5px 10px",fontSize:11,color:"#64748b",fontWeight:600,cursor:"pointer"}}>{l}</button>
              ))}
            </div>
            <div style={{borderTop:"1px solid #6b8fb5",padding:"9px 14px",display:"flex",gap:7,marginTop:6}}>
              <a href={"tel:"+c.phone} onClick={()=>startCall(c)} style={{flex:1,background:"rgba(34,197,94,0.1)",border:"1px solid rgba(34,197,94,0.3)",borderRadius:10,padding:"9px 4px",display:"flex",alignItems:"center",justifyContent:"center",gap:5,textDecoration:"none",color:"#22c55e",fontWeight:600,fontSize:12}}><PhIcon s={14}/>Call</a>
              <a href={"https://wa.me/91"+c.phone+"?text="+encodeURIComponent("Hello "+c.name+", NKD Bajaj here. Following up on your "+c.model+" enquiry.")} target="_blank" rel="noreferrer" style={{flex:1,background:"rgba(37,211,102,0.08)",border:"1px solid rgba(37,211,102,0.3)",borderRadius:10,padding:"9px 4px",display:"flex",alignItems:"center",justifyContent:"center",gap:5,textDecoration:"none",fontWeight:600,fontSize:12}}><WAIcon s={14}/><span style={{color:"#25D366"}}>WA</span></a>
              <button onClick={()=>setActive(active===c.id?null:c.id)} style={{flex:1,background:active===c.id?"rgba(249,115,22,0.2)":"rgba(249,115,22,0.08)",border:"1px solid rgba(249,115,22,0.3)",borderRadius:10,padding:"9px 4px",cursor:"pointer",color:"#f97316",fontWeight:600,fontSize:12}}>{active===c.id?"Close":"Log"}</button>
            </div>
            {active===c.id&&(
              <div style={{borderTop:"1px solid #6b8fb5",padding:"12px 14px",background:"#f1f5f9",display:"flex",flexDirection:"column",gap:8}}>
                <select style={inp} value={form.out} onChange={e=>setForm(p=>({...p,out:e.target.value}))}>
                  <option value="interested">Interested — followup later</option>
                  <option value="booked">Booked ✓</option>
                  <option value="visit_scheduled">Visit Scheduled</option>
                  <option value="no_response">No Response</option>
                  <option value="switched_off">Switched Off</option>
                  <option value="not_interested">Not Interested</option>
                  <option value="price_issue">Price Issue</option>
                </select>
                {form.out==="not_interested"&&<select style={inp} value={form.reason} onChange={e=>setForm(p=>({...p,reason:e.target.value}))}>
                  <option value="">Why lost? (required)</option><option>Price too high</option><option>Bought competitor brand</option><option>Finance rejected</option><option>Postponed purchase</option><option>Family decision</option><option>Other</option>
                </select>}
                {form.out==="not_interested"&&form.reason==="Bought competitor brand"&&<input style={inp} value={form.competitor} onChange={e=>setForm(p=>({...p,competitor:e.target.value}))} placeholder="Which brand/model? (e.g. Honda Shine)"/>}
                <textarea rows={2} style={{...inp,resize:"none"}} value={form.note} onChange={e=>setForm(p=>({...p,note:e.target.value}))} placeholder="What did customer say?"/>
                <input type="date" style={inp} value={form.nxt} onChange={e=>setForm(p=>({...p,nxt:e.target.value}))}/>
                <div style={{display:"flex",gap:8}}>
                  <button onClick={()=>submit(c)} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),flex:1}}>Save Followup</button>
                  <button onClick={()=>setActive(null)} style={{...btn("#6b8fb5","#8892a4"),flex:1}}>Cancel</button>
                </div>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function CustList({custs,onOpen,initF,showSM}){
  const [q,setQ]=useState("");
  const [st,setSt]=useState(initF||"All");
  useEffect(()=>{if(initF)setSt(initF);},[initF]);
  const [srt,setSrt]=useState("priority");
  const [md,setMd]=useState("All");
  const [smf,setSmf]=useState("All");
  const filtered=useMemo(()=>custs.filter(c=>{
    const ql=q.toLowerCase();
    const hay=(c.name+" "+c.phone+" "+(c.model||"")+" "+(c.modelCode||"")+" "+(c.aadhar||"")+" "+(c.pan||"")+" "+(c.salesman||"")+" "+(c.financeBank||"")+" "+((c.billing&&c.billing.chassis)||"")+" "+((c.billing&&c.billing.engine)||"")+" "+((c.billing&&c.billing.registrationNo)||"")+" "+(c.exchangeRegNo||"")).toLowerCase();
    const ms=!q||hay.includes(ql);
    const mm=md==="All"||(md==="Cash"&&c.finance==="Cash")||(md==="Finance"&&c.finance==="Finance")||(md==="Exchange"&&(c.exchangeAsked||c.exchangeOffered));
    const msm=smf==="All"||c.salesman===smf;
    return ms&&mm&&msm&&(st==="All"||c.status===st);
  }).sort((a,b)=>{
    if(srt==="followup")return new Date(a.followupDate)-new Date(b.followupDate);
    if(srt==="priority"){const o={Hot:0,Warm:1,Cold:2,Booked:3,Billed:4,Lost:5};return(o[a.status]??9)-(o[b.status]??9);}
    if(srt==="name")return a.name.localeCompare(b.name);
    return new Date(b.enquiryDate)-new Date(a.enquiryDate);
  }),[custs,q,st,srt,md,smf]);
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:12}}>Customers <span style={{color:"#94a3b8",fontSize:13,fontWeight:400}}>({filtered.length})</span></div>
      <input placeholder="🔍 Search anything — name, phone, chassis, engine, Aadhar, reg no…" style={{...inp,marginBottom:10,padding:"13px 14px",fontSize:14,borderRadius:13}} value={q} onChange={e=>setQ(e.target.value)}/>
      <div style={{display:"flex",gap:5,marginBottom:8}}>
        {["All","Cash","Finance","Exchange"].map(m=><button key={m} onClick={()=>setMd(m)} style={{flex:1,padding:"6px",borderRadius:9,fontSize:11,fontWeight:700,cursor:"pointer",background:md===m?"#dbeafe":"#6b8fb5",color:md===m?"#60a5fa":"#8892a4",border:"none"}}>{m}</button>)}
      </div>
      {showSM&&<select style={{...inp,marginBottom:8}} value={smf} onChange={e=>setSmf(e.target.value)}><option value="All">All Executives</option>{SM.map(s=><option key={s}>{s}</option>)}</select>}
      <div style={{display:"flex",gap:5,overflowX:"auto",marginBottom:9,paddingBottom:2}}>
        {["All","Hot","Warm","Cold","Booked","Billed","Lost"].map(s=><button key={s} onClick={()=>setSt(s)} style={{padding:"5px 11px",borderRadius:20,fontSize:11,fontWeight:600,cursor:"pointer",flexShrink:0,background:st===s?(ST_C[s]||"#f97316"):"#6b8fb5",color:st===s?"#fff":"#8892a4",border:"none"}}>{s}</button>)}
      </div>
      <div style={{display:"flex",gap:5,marginBottom:12}}>
        {[["followup","Followup"],["priority","Priority"],["name","A-Z"],["recent","Recent"]].map(([k,l])=>(
          <button key={k} onClick={()=>setSrt(k)} style={{padding:"4px 10px",borderRadius:8,fontSize:10,fontWeight:600,cursor:"pointer",background:srt===k?"#dbeafe":"#ffffff",color:srt===k?"#60a5fa":"#5a6478",border:"1px solid "+(srt===k?"#2563eb":"#6b8fb5")}}>{l}</button>
        ))}
      </div>
      {filtered.length===0&&<div className="fu glass" style={{textAlign:"center",padding:"36px 20px",borderRadius:16,background:"#ffffff"}}><div style={{fontSize:36,marginBottom:6}}>👥</div><div style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>No customers here yet</div><div style={{fontSize:12,color:"#94a3b8",marginTop:4}}>Tap the orange + button to add a new enquiry</div></div>}
      {filtered.map(c=><Card key={c.id} c={c} onClick={()=>onOpen(c)} showSM={showSM}/>)}
    </div>
  );
}

function DocGrid({cust,onUpload,docs}){
  return(
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      {docs.map(({key,l,ic})=>(
        <div key={key} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,overflow:"hidden"}}>
          <div style={{padding:"10px 13px",display:"flex",justifyContent:"space-between",alignItems:"center"}}><div style={{display:"flex",alignItems:"center",gap:8}}><span style={{fontSize:18}}>{ic}</span><span style={{fontSize:13,color:"#334155"}}>{l}</span></div>{(cust.photos||{})[key]&&<span style={{fontSize:11,color:"#22c55e",fontWeight:700}}>✓</span>}</div>
          {(cust.photos||{})[key]&&<img src={cust.photos[key]} alt={l} style={{width:"100%",maxHeight:200,objectFit:"contain",background:"#000"}}/>}
          <div style={{padding:"0 13px 12px"}}>
            <div style={{fontSize:11,color:"#64748b",marginBottom:4}}>{(cust.photos||{})[key]?"🔄 Replace photo:":"📷 Choose photo:"}</div>
            <input type="file" accept="image/*" onChange={e=>{if(e.target.files&&e.target.files[0]){onUpload(key,e.target.files[0]);e.target.value="";}}} style={{width:"100%",background:"#c2d6ec",borderRadius:9,padding:8,fontSize:12,color:"#64748b",border:"1px dashed #2a3040"}}/>
          </div>
        </div>
      ))}
    </div>
  );
}
function parseExcel(file,cb,errCb){
  const rd=new FileReader();
  rd.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const data=XLSX.utils.sheet_to_json(ws,{defval:""});
      if(data.length===0){errCb("Empty file");return;}
      cb(data);
    }catch(err){errCb("Could not read file");}
  };
  rd.readAsArrayBuffer(file);
}
function UploadsHub({stockData,statusData,onStockUpload,onStatusUpload,notify}){
  const uploads=[
    {
      id:"stock",ic:"🏍️",title:"Stock Statement",color:"#34d399",bg:"rgba(52,211,153,0.08)",border:"rgba(52,211,153,0.4)",
      desc:"Branch-wise vehicle stock from dealership",
      current:stockData.length>0?stockData.length+" vehicles loaded":null,
      onFile:(file)=>parseExcel(file,d=>{onStockUpload(d);notify("✅ Stock uploaded — "+d.length+" vehicles");},e=>notify("❌ "+e)),
    },
    {
      id:"rcstatus",ic:"📋",title:"RC / HSRP Status",color:"#60a5fa",bg:"rgba(96,165,250,0.08)",border:"rgba(96,165,250,0.4)",
      desc:"RC and HSRP status report from RTO",
      current:statusData.length>0?statusData.length+" records loaded":null,
      onFile:(file)=>parseExcel(file,d=>{onStatusUpload(d);notify("✅ RC/HSRP status uploaded — "+d.length+" records");},e=>notify("❌ "+e)),
    },
  ];
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:4}}>📤 Uploads</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:18}}>Upload Excel files to keep data fresh. Any format, any columns.</div>
      {uploads.map(u=>(
        <div key={u.id} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:16,padding:"16px 16px",marginBottom:14}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:10}}>
            <span style={{fontSize:30}}>{u.ic}</span>
            <div>
              <div style={{fontWeight:700,fontSize:15,color:"#1e293b"}}>{u.title}</div>
              <div style={{fontSize:11,color:"#94a3b8"}}>{u.desc}</div>
            </div>
          </div>
          {u.current&&<div style={{background:"rgba(34,197,94,0.08)",border:"1px solid rgba(34,197,94,0.25)",borderRadius:9,padding:"6px 12px",marginBottom:10,fontSize:12,color:"#22c55e",fontWeight:600}}>✅ {u.current}</div>}
          <label style={{display:"block",background:u.bg,border:"1px dashed "+u.border,borderRadius:11,padding:"12px",cursor:"pointer",textAlign:"center"}}>
            <div style={{fontSize:13,color:u.color,fontWeight:700}}>{u.current?"🔄 Replace":"📂 Choose Excel File"}</div>
            <div style={{fontSize:10,color:"#94a3b8",marginTop:2}}>.xlsx · .xls · .csv supported</div>
            <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files&&e.target.files[0]){u.onFile(e.target.files[0]);e.target.value="";}}}/>
          </label>
        </div>
      ))}
    </div>
  );
}
function findStockCol(keys,words){return keys.find(k=>words.some(w=>k.toLowerCase().includes(w)))||null;}
function cleanBranch(raw){const s=String(raw||"").toLowerCase();if(s.includes("chirkunda"))return"Chirkunda";if(s.includes("saraidhela")||s.includes("saraidela"))return"Saraidhela";if(s.includes("hirak")||s.includes("12967")||s.includes("nkd"))return"Hirak Road";return raw||"";}
function excelDateToAge(raw){const n=typeof raw==="number"?raw:Number(raw);if(n>40000&&n<100000){const dt=new Date(Math.round((n-25569)*86400000));return Math.max(0,Math.floor((Date.now()-dt.getTime())/86400000));}return null;}
function StockView({stockData,billedChassis,role,userBranch,onUpload,notify}){
  const [q,setQ]=useState("");
  const [tab,setTab]=useState("stock"); // "stock" | "ageing"
  const rows=stockData||[];
  const keys=rows.length>0?Object.keys(rows[0]):[];
  const chassisKey=findStockCol(keys,["chassis","frame"]);
  const engineKey=findStockCol(keys,["engine"]);
  const colorKey=findStockCol(keys,["color","colour"]);
  const modelKey=findStockCol(keys,["model","name","variant"]);
  const branchKey=findStockCol(keys,["branch","location","godown","store"]);
  const dateKey=findStockCol(keys,["date","invoice","inward","receipt","received","entry","purchase","billing"]);
  const available=rows.filter(r=>!billedChassis.includes(String(r[chassisKey]||"").trim().toUpperCase()));
  function isMyBranch(row){if(!userBranch||!branchKey)return false;return String(row[branchKey]||"").toLowerCase().includes(userBranch.toLowerCase());}
  function branchSort(a,b){if(isMyBranch(a)&&!isMyBranch(b))return -1;if(!isMyBranch(a)&&isMyBranch(b))return 1;return 0;}

  function getAge(row){
    if(!dateKey||!row[dateKey])return null;
    const rawVal=row[dateKey];
    // Excel serial date (number or numeric string like 46164.5)
    const numVal=typeof rawVal==="number"?rawVal:Number(rawVal);
    if(!isNaN(numVal)&&numVal>40000&&numVal<100000){return excelDateToAge(numVal);}
    // String date formats
    const raw=String(rawVal).trim();
    let d=new Date(raw);
    if(isNaN(d)&&raw.includes("/")){const p=raw.split("/");if(p.length===3)d=new Date(p[2]+"-"+p[1].padStart(2,"0")+"-"+p[0].padStart(2,"0"));}
    if(isNaN(d)&&raw.includes("-")&&raw.length<=10){const p=raw.split("-");if(p.length===3&&p[0].length<=2)d=new Date(p[2]+"-"+p[1].padStart(2,"0")+"-"+p[0].padStart(2,"0"));}
    if(isNaN(d))return null;
    return Math.max(0,Math.floor((Date.now()-d.getTime())/86400000));
  }
  function ageBadge(days){
    if(days===null)return{bg:"#f1f5f9",border:"#6b8fb5",col:"#64748b",label:"No date",tag:"—"};
    if(days<=30)return{bg:"rgba(34,197,94,0.1)",border:"rgba(34,197,94,0.5)",col:"#16a34a",label:"Fresh",tag:days+"d"};
    if(days<=60)return{bg:"rgba(234,179,8,0.1)",border:"rgba(234,179,8,0.55)",col:"#ca8a04",label:"Moderate",tag:days+"d"};
    if(days<=90)return{bg:"rgba(249,115,22,0.1)",border:"rgba(249,115,22,0.55)",col:"#ea580c",label:"Ageing",tag:days+"d"};
    return{bg:"rgba(239,68,68,0.1)",border:"rgba(239,68,68,0.6)",col:"#dc2626",label:"OLD STOCK",tag:days+"d ⚠️"};
  }

  const filtered=useMemo(()=>{
    const base=q.trim().length<2?available:available.filter(r=>keys.some(k=>String(r[k]||"").toLowerCase().includes(q.toLowerCase())));
    return [...base].sort(branchSort);
  },[available,q,keys,userBranch,branchKey]);
  const byModel=useMemo(()=>{
    if(!modelKey)return[];
    const map={};
    filtered.forEach(r=>{const m=String(r[modelKey]||"Other");map[m]=(map[m]||0)+1;});
    return Object.entries(map).sort((a,b)=>b[1]-a[1]);
  },[filtered,modelKey]);

  // Ageing: own branch first, then oldest first
  const ageingRows=useMemo(()=>{
    const withAge=available.map(r=>({...r,__age:getAge(r)}));
    const qL=q.toLowerCase().trim();
    const filt=qL.length<2?withAge:withAge.filter(r=>keys.some(k=>String(r[k]||"").toLowerCase().includes(qL)));
    return filt.sort((a,b)=>{
      const bs=branchSort(a,b);if(bs!==0)return bs;
      if(a.__age===null&&b.__age===null)return 0;
      if(a.__age===null)return 1;
      if(b.__age===null)return -1;
      return b.__age-a.__age;
    });
  },[available,q,keys,userBranch,branchKey]);

  // Ageing summary
  const ageSummary=useMemo(()=>{
    const s={fresh:0,mod:0,ageing:0,old:0,nodate:0};
    available.forEach(r=>{const d=getAge(r);if(d===null)s.nodate++;else if(d<=30)s.fresh++;else if(d<=60)s.mod++;else if(d<=90)s.ageing++;else s.old++;});
    return s;
  },[available]);

  const StockCard=({row,i})=>{
    const age=row.__age!==undefined?row.__age:getAge(row);
    const ab=ageBadge(age);
    const mine=isMyBranch(row);
    const branchName=branchKey?cleanBranch(row[branchKey]):"";
    return(
      <div key={i} style={{background:mine?"rgba(59,130,246,0.04)":age>90?"rgba(239,68,68,0.04)":age>60?"rgba(249,115,22,0.04)":"#ffffff",border:"2px solid "+(mine?"#3b82f6":ab.border),borderRadius:12,padding:"10px 14px",marginBottom:8,position:"relative"}}>
        {/* Branch row */}
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5}}>
          <div style={{display:"flex",alignItems:"center",gap:6}}>
            {branchName&&<span style={{fontSize:11,fontWeight:700,background:mine?"#dbeafe":"#f1f5f9",color:mine?"#1d4ed8":"#475569",padding:"2px 9px",borderRadius:8,border:"1px solid "+(mine?"#93c5fd":"#6b8fb5")}}>{mine?"📍 "+branchName+" ★":"📍 "+branchName}</span>}
            {mine&&<span style={{fontSize:9,fontWeight:800,color:"#1d4ed8",background:"#dbeafe",padding:"1px 6px",borderRadius:6}}>YOUR BRANCH</span>}
          </div>
          {age!==null&&<span style={{fontSize:10,fontWeight:800,background:ab.bg,color:ab.col,padding:"2px 8px",borderRadius:10,border:"1px solid "+ab.border,flexShrink:0}}>{ab.tag}</span>}
        </div>
        {modelKey&&<div style={{fontWeight:700,fontSize:13,color:"#1e293b",marginBottom:5}}>{row[modelKey]}</div>}
        <div style={{display:"flex",flexWrap:"wrap",gap:6}}>
          {chassisKey&&<span style={{fontSize:11,color:"#60a5fa"}}>🔩 {row[chassisKey]}</span>}
          {engineKey&&<span style={{fontSize:11,color:"#64748b"}}>⚙️ {row[engineKey]}</span>}
          {colorKey&&<span style={{fontSize:11,color:"#a78bfa"}}>🎨 {row[colorKey]}</span>}
          {dateKey&&row[dateKey]&&<span style={{fontSize:11,color:"#94a3b8"}}>📅 {String(row[dateKey]).slice(0,12)}</span>}
        </div>
        {age>90&&<div style={{fontSize:10,color:"#dc2626",fontWeight:700,marginTop:5}}>⚠️ Push for sale — in stock {age} days</div>}
      </div>
    );
  };

  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:2}}>🏍️ Stock</div>
      {userBranch&&<div style={{fontSize:11,color:"#1d4ed8",fontWeight:700,marginBottom:4}}>📍 Your branch: {userBranch}</div>}
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:10}}>{rows.length>0?available.length+" available · "+billedChassis.length+" billed":"No stock uploaded yet"}</div>

      {/* Tab toggle */}
      {rows.length>0&&<div style={{display:"flex",gap:8,marginBottom:12}}>
        {[["stock","📦 Stock Search"],["ageing","📊 Ageing Report"]].map(([id,l])=>(
          <button key={id} onClick={()=>setTab(id)} style={{flex:1,padding:"9px",borderRadius:10,fontSize:12,fontWeight:700,cursor:"pointer",border:"1px solid "+(tab===id?"#3b82f6":"#6b8fb5"),background:tab===id?"#dbeafe":"#ffffff",color:tab===id?"#1d4ed8":"#475569"}}>{l}</button>
        ))}
      </div>}

      {rows.length===0&&<div style={{background:"rgba(52,211,153,0.07)",border:"1px solid rgba(52,211,153,0.3)",borderRadius:12,padding:16,textAlign:"center",marginBottom:14}}><div style={{fontSize:28,marginBottom:6}}>📦</div><div style={{fontSize:13,color:"#34d399",fontWeight:700}}>No stock data yet</div><div style={{fontSize:11,color:"#94a3b8",marginTop:4}}>{isPortalRole(role)?"Go to 📤 Uploads tab to upload stock Excel":"Ask Owner/Admin to upload stock Excel"}</div></div>}

      {rows.length>0&&<input placeholder="🔍 Search model, chassis, colour, branch…" style={{...inp,marginBottom:10,padding:"11px 14px",fontSize:13,borderRadius:12}} value={q} onChange={e=>setQ(e.target.value)}/>}

      {/* ── STOCK SEARCH TAB ── */}
      {tab==="stock"&&<>
        {q.trim().length<2&&byModel.length>0&&(
          <div style={{display:"flex",flexWrap:"wrap",gap:6,marginBottom:12}}>
            {byModel.map(([m,cnt])=>(
              <div key={m} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:10,padding:"6px 12px",cursor:"pointer"}} onClick={()=>setQ(m)}>
                <span style={{fontSize:11,color:"#1e293b",fontWeight:600}}>{m}</span>
                <span style={{fontSize:10,color:"#34d399",fontWeight:700,marginLeft:6}}>{cnt}</span>
              </div>
            ))}
          </div>
        )}
        {q.trim().length>=2&&filtered.length===0&&<div style={{textAlign:"center",padding:20,color:"#94a3b8",fontSize:13}}>No stock found for "{q}"</div>}
        {filtered.slice(0,60).map((row,i)=>{
          const isMine=isMyBranch(row);
          const prevMine=i>0?isMyBranch(filtered[i-1]):true;
          const showDivider=userBranch&&branchKey&&!isMine&&prevMine&&i>0;
          return(<Fragment key={i}>
            {showDivider&&<div style={{textAlign:"center",fontSize:11,color:"#94a3b8",margin:"10px 0 8px",display:"flex",alignItems:"center",gap:8}}><div style={{flex:1,height:1,background:"#6b8fb5"}}/><span>Other Branches</span><div style={{flex:1,height:1,background:"#6b8fb5"}}/></div>}
            <StockCard row={row} i={i}/>
          </Fragment>);
        })}
        {filtered.length>60&&<div style={{textAlign:"center",color:"#94a3b8",fontSize:12,padding:8}}>Showing 60 of {filtered.length} — search to narrow down</div>}
      </>}

      {/* ── AGEING REPORT TAB ── */}
      {tab==="ageing"&&<>
        {/* Summary cards */}
        {dateKey&&<div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:7,marginBottom:14}}>
          {[["0–30d","Fresh",ageSummary.fresh,"#16a34a","rgba(34,197,94,0.12)"],["31–60d","Moderate",ageSummary.mod,"#ca8a04","rgba(234,179,8,0.12)"],["61–90d","Ageing",ageSummary.ageing,"#ea580c","rgba(249,115,22,0.12)"],["90d+","Old Stock",ageSummary.old,"#dc2626","rgba(239,68,68,0.12)"]].map(([range,label,cnt,col,bg])=>(
            <div key={range} style={{background:bg,border:"1px solid "+col+"55",borderRadius:10,padding:"8px 6px",textAlign:"center"}}>
              <div style={{fontSize:18,fontWeight:900,color:col}}>{cnt}</div>
              <div style={{fontSize:9,fontWeight:700,color:col}}>{label}</div>
              <div style={{fontSize:9,color:"#94a3b8",marginTop:1}}>{range}</div>
            </div>
          ))}
        </div>}
        {!dateKey&&<div style={{background:"rgba(249,115,22,0.08)",border:"1px solid rgba(249,115,22,0.35)",borderRadius:10,padding:"10px 14px",marginBottom:12,fontSize:12,color:"#ea580c"}}>⚠️ No date column found in stock Excel. Add a column like "Invoice Date" or "Inward Date" to see ageing.</div>}
        {ageingRows.length===0&&<div style={{textAlign:"center",padding:20,color:"#94a3b8",fontSize:13}}>No stock to show</div>}
        {ageingRows.slice(0,80).map((row,i)=>{
          const isMine=isMyBranch(row);
          const prevMine=i>0?isMyBranch(ageingRows[i-1]):true;
          const showDivider=userBranch&&branchKey&&!isMine&&prevMine&&i>0;
          return(<Fragment key={i}>
            {showDivider&&<div style={{textAlign:"center",fontSize:11,color:"#94a3b8",margin:"10px 0 8px",display:"flex",alignItems:"center",gap:8}}><div style={{flex:1,height:1,background:"#6b8fb5"}}/><span>Other Branches</span><div style={{flex:1,height:1,background:"#6b8fb5"}}/></div>}
            <StockCard row={row} i={i}/>
          </Fragment>);
        })}
        {ageingRows.length>80&&<div style={{textAlign:"center",color:"#94a3b8",fontSize:12,padding:8}}>Showing 80 of {ageingRows.length}</div>}
      </>}
    </div>
  );
}
function RCHSRPSearch({statusData,role,onUpload,notify}){
  const [q,setQ]=useState("");
  const rows=statusData||[];
  const cols=rows.length>0?Object.keys(rows[0]):[];
  const filtered=q.trim().length<2?[]:rows.filter(r=>cols.some(c=>String(r[c]||"").toLowerCase().includes(q.toLowerCase())));
  function handleFile(file){
    const rd=new FileReader();
    rd.onload=function(e){
      try{
        const wb=XLSX.read(e.target.result,{type:"array"});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const data=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(data.length===0){notify("❌ Empty Excel file");return;}
        onUpload(data);
        notify("✅ Status data uploaded — "+data.length+" records");
      }catch(err){notify("❌ Could not read file");}
    };
    rd.readAsArrayBuffer(file);
  }
  const STATUS_KEYS=["hsrp","rc","insurance","noc","status","remark","rto"];
  function statusColor(val){
    const v=String(val||"").toLowerCase();
    if(v.includes("done")||v.includes("received")||v.includes("fitted")||v.includes("yes")||v.includes("complete"))return"#22c55e";
    if(v.includes("pending")||v.includes("applied")||v.includes("process"))return"#f97316";
    if(v.includes("no")||v.includes("not"))return"#ef4444";
    return"#8892a4";
  }
  function isStatusCol(c){return STATUS_KEYS.some(k=>c.toLowerCase().includes(k));}
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:4}}>🔍 RC / HSRP Status</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:14}}>{rows.length>0?rows.length+" records loaded":"No data yet — upload Excel below"}</div>
      {rows.length===0&&<div style={{background:"rgba(249,115,22,0.07)",border:"1px solid rgba(249,115,22,0.2)",borderRadius:12,padding:16,textAlign:"center",marginBottom:14}}><div style={{fontSize:28,marginBottom:6}}>📋</div><div style={{fontSize:13,color:"#f97316",fontWeight:700}}>No status data uploaded yet</div><div style={{fontSize:11,color:"#94a3b8",marginTop:4}}>{isPortalRole(role)?"Go to 📤 Uploads tab to upload RC/HSRP Excel":"Ask Owner/Admin to upload the RC/HSRP status Excel"}</div></div>}
      {rows.length>0&&(
        <input placeholder="🔍 Search by name, chassis, reg no, engine, phone…" style={{...inp,marginBottom:12,padding:"13px 14px",fontSize:14,borderRadius:13}} value={q} onChange={e=>setQ(e.target.value)} autoFocus/>
      )}
      {q.trim().length>=2&&filtered.length===0&&<div style={{textAlign:"center",padding:24,color:"#94a3b8",fontSize:13}}>No records found for "{q}"</div>}
      {filtered.map((row,i)=>(
        <div key={i} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:13,padding:"12px 14px",marginBottom:9}}>
          {cols.map(c=>{
            const val=row[c];
            const isSt=isStatusCol(c);
            return val!==""?(
              <div key={c} style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:5,gap:8}}>
                <span style={{fontSize:11,color:"#94a3b8",flexShrink:0}}>{c}</span>
                <span style={{fontSize:12,fontWeight:isSt?700:500,color:isSt?statusColor(val):"#e2e6f0",textAlign:"right"}}>{String(val)}</span>
              </div>
            ):null;
          })}
        </div>
      ))}
      {q.trim().length<2&&rows.length>0&&<div style={{textAlign:"center",color:"#64748b",fontSize:12,padding:16}}>Type at least 2 characters to search</div>}
    </div>
  );
}
function Detail({cust,role,onBack,onUpd,onLog,onBill,onBook,notify,initTab,clearInit,onAddPayment}){
  const [tab,setTab]=useState(initTab||"info");
  useEffect(()=>{if(initTab){setTab(initTab);clearInit&&clearInit();}},[initTab]);
  const [edit,setEdit]=useState(false);
  const [f,setF]=useState({...cust});
  const r=RC[cust.modelCode];
  const lastR=(cust.remarks||"").trim().split("\n").filter(Boolean).slice(-1)[0]||"";

  function saveEdit(){
    if(f.phone&&!/^\d{10}$/.test(f.phone)){notify("⚠️ Phone must be exactly 10 digits","err");return;}
    onUpd(f);setEdit(false);notify("Saved ✓");
  }
  function pickM(code){const m=RC[code];setF(p=>({...p,modelCode:code,model:m?m.n:"",cat:m?m.cat:""}));}
  function uploadPhoto(key,file){compressImg(file,function(dataUrl){
    onUpd({photos:{...(cust.photos||{}),[key]:dataUrl}});notify("Uploading to Drive…");
    uploadToDrive(key+"_"+Date.now()+".jpg",dataUrl,file.type||"image/jpeg",cust.name,key,function(url){
      onUpd({photos:{...(cust.photos||{}),[key]:url}});notify("✅ Saved to Google Drive");
    });
  });}

  const tabs=["info","history","followup","docs",...(cust.billed?["billing"]:[])];

  return(
    <div>
      <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:14}}>
        <button onClick={onBack} style={{background:"#c2d6ec",border:"none",borderRadius:10,width:36,height:36,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",color:"#64748b",fontSize:20}}>←</button>
        <div style={{flex:1}}><div style={{fontWeight:800,fontSize:16,color:"#1e293b"}}>{cust.name}</div><div style={{fontSize:11,color:"#94a3b8"}}>{cust.id} · {cust.modelCode}</div></div>
        <Badge s={cust.status}/>
      </div>

      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:12}}>
        <a href={"tel:"+cust.phone} style={{background:"rgba(34,197,94,0.1)",border:"1px solid rgba(34,197,94,0.3)",borderRadius:13,padding:"11px 6px",display:"flex",flexDirection:"column",alignItems:"center",gap:4,textDecoration:"none"}}><PhIcon s={22}/><span style={{fontSize:11,color:"#22c55e",fontWeight:700}}>Call</span><span style={{fontSize:10,color:"#94a3b8"}}>{cust.phone}</span></a>
        <a href={"https://wa.me/91"+cust.phone+"?text="+encodeURIComponent("Hello "+cust.name+", NKD Bajaj here. Following up on your "+cust.model+" enquiry.")} target="_blank" rel="noreferrer" style={{background:"rgba(37,211,102,0.08)",border:"1px solid rgba(37,211,102,0.3)",borderRadius:13,padding:"11px 6px",display:"flex",flexDirection:"column",alignItems:"center",gap:4,textDecoration:"none"}}><WAIcon s={22}/><span style={{fontSize:11,color:"#25D366",fontWeight:700}}>WhatsApp</span><span style={{fontSize:10,color:"#94a3b8"}}>Message</span></a>
        {!cust.billed&&!cust.stopped?<button onClick={onBill} style={{background:"rgba(52,211,153,0.1)",border:"1px solid rgba(52,211,153,0.3)",borderRadius:13,padding:"11px 6px",display:"flex",flexDirection:"column",alignItems:"center",gap:4,cursor:"pointer"}}><span style={{fontSize:22}}>🏍️</span><span style={{fontSize:11,color:"#34d399",fontWeight:700}}>{cust.billingDraft?"Resume":"Bill"}</span><span style={{fontSize:10,color:cust.billingDraft?"#f59e0b":"#94a3b8"}}>{cust.billingDraft?"Draft Saved":"Vehicle"}</span></button>:<div style={{background:"rgba(52,211,153,0.07)",border:"1px solid rgba(52,211,153,0.25)",borderRadius:13,padding:"11px 6px",display:"flex",flexDirection:"column",alignItems:"center",gap:4}}><span style={{fontSize:22}}>✅</span><span style={{fontSize:11,color:"#34d399",fontWeight:700}}>Billed</span><span style={{fontSize:10,color:"#94a3b8"}}>{fd(cust.billedDate)}</span></div>}
      </div>

      {!cust.billed&&!cust.stopped&&(cust.booking&&role==="salesman"?
        <div style={{width:"100%",background:"rgba(139,92,246,0.07)",border:"1px solid rgba(139,92,246,0.25)",borderRadius:12,padding:"11px",color:"#a78bfa",fontWeight:700,fontSize:13,marginBottom:8,textAlign:"center"}}>📝 Booking: {fc(cust.booking.amt)} on {fd(cust.booking.date)} · {cust.booking.mode}<br/><span style={{fontSize:10,fontWeight:400,color:"#94a3b8"}}>Only manager can edit booking amount</span></div>:
        <button onClick={onBook} style={{width:"100%",background:"rgba(139,92,246,0.1)",border:"1px solid rgba(139,92,246,0.35)",borderRadius:12,padding:"11px",color:"#a78bfa",fontWeight:700,fontSize:13,cursor:"pointer",marginBottom:cust.booking&&cust.booking.receiptHtml?8:12}}>📝 {cust.booking?"Booking: "+fc(cust.booking.amt)+" on "+fd(cust.booking.date)+" — Edit":"Take Booking Amount (without documents)"}</button>
      )}
      {!cust.billed&&!cust.stopped&&<button onClick={()=>{const taking=!cust.testRide;onUpd({testRide:taking?{date:td()}:null,remarks:(cust.remarks||"")+(taking?"\n["+td()+"] TEST RIDE taken":"")});if(taking){setTab("docs");notify("🚦 Test ride recorded — upload the license here");}}} style={{width:"100%",background:cust.testRide?"rgba(34,197,94,0.1)":"rgba(96,165,250,0.08)",border:"1px solid "+(cust.testRide?"rgba(34,197,94,0.35)":"rgba(96,165,250,0.3)"),borderRadius:12,padding:"10px",color:cust.testRide?"#22c55e":"#60a5fa",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:8}}>🏍️ Test Ride {cust.testRide?"✓ Done "+fd(cust.testRide.date):"— Tap when taken (upload license in Docs)"}</button>}
      {!cust.billed&&cust.booking&&<div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:12}}>
        <button onClick={()=>{const doc=makeBookingPdf(cust);sharePdf(doc,"BookingReceipt_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",cust.phone,"Please find your Booking Receipt from NKD Bajaj, Dhanbad.");}} style={{width:"100%",background:"rgba(37,211,102,0.1)",border:"1px solid rgba(37,211,102,0.35)",borderRadius:12,padding:13,color:"#22c55e",fontWeight:700,fontSize:13,cursor:"pointer"}}>📲 Send Booking Receipt PDF → Customer (WhatsApp)</button>
        <button onClick={()=>{const num=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;const doc=makeBookingPdf(cust);sharePdf(doc,"BookingRecord_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",num,"Booking Record for "+cust.name+" — "+fc(cust.booking.amt)+" ("+cust.booking.mode+")");}} style={{width:"100%",background:"rgba(139,92,246,0.1)",border:"1px solid rgba(139,92,246,0.35)",borderRadius:12,padding:11,color:"#a78bfa",fontWeight:700,fontSize:12,cursor:"pointer"}}>🏢 Send Booking Record PDF → Office (WhatsApp)</button>
        <div style={{fontSize:10,color:"#94a3b8"}}>On mobile — tapping Send opens WhatsApp share sheet directly. On desktop — PDF downloads then WhatsApp opens.</div>
      </div>}
      {cust.managerApproval==="rejected"&&!cust.billed&&<div style={{background:"rgba(239,68,68,0.12)",border:"1px solid rgba(239,68,68,0.5)",borderRadius:12,padding:"11px 13px",marginBottom:12}}>
        <div style={{fontSize:12,color:"#ef4444",fontWeight:800,marginBottom:3}}>❌ BILLING REJECTED BY MANAGER</div>
        <div style={{fontSize:12,color:"#1e293b",lineHeight:1.5}}>Correct the calculation sheet and tap 🏍️ Bill again. All uploaded documents are retained.</div>
      </div>}
      {lastR&&<div style={{background:"rgba(249,115,22,0.07)",border:"1px solid rgba(249,115,22,0.22)",borderRadius:11,padding:"10px 12px",marginBottom:12}}><div style={{fontSize:9,color:"#f97316",fontWeight:800,letterSpacing:0.6,marginBottom:3}}>📋 LAST REMARK</div><div style={{fontSize:13,color:"#1e293b",lineHeight:1.5}}>{lastR.replace(/^\[.*?\]\s*/,"").replace(/^[A-Z_]+:\s*/,"")}</div></div>}

      {(()=>{const calls=(cust.callLog||[]).filter(l=>l.duration);const lastCall=calls.length?calls[calls.length-1]:null;
        return(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr 1fr",gap:6,marginBottom:12}}>
          {[["📞 Calls",calls.length],["📋 Logs",(cust.callLog||[]).length],["Last Call",lastCall?fd(lastCall.date):"—"],["Next",cust.billed?"Done":fd(cust.followupDate)]].map(([k,v])=>(
            <div key={k} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:11,padding:"9px 6px",textAlign:"center"}}>
              <div style={{fontSize:13,fontWeight:800,color:"#1e293b"}}>{v}</div>
              <div style={{fontSize:9,color:"#94a3b8",marginTop:2}}>{k}</div>
            </div>
          ))}
        </div>);})()}
      <div style={{display:"flex",gap:4,overflowX:"auto",marginBottom:14,paddingBottom:2}}>
        {tabs.map(t=><button key={t} onClick={()=>setTab(t)} style={{padding:"7px 14px",borderRadius:10,fontSize:12,fontWeight:600,cursor:"pointer",flexShrink:0,background:tab===t?"#f97316":"#6b8fb5",color:tab===t?"#fff":"#8892a4",border:"none",textTransform:"capitalize"}}>{t==="aftersale"?"After Sale":t}</button>)}
      </div>

      {tab==="info"&&(
        <div>
          {edit?(
            <div style={{display:"flex",flexDirection:"column",gap:10}}>
              {[{k:"name",l:"Name"},{k:"phone",l:"Phone"+(cust.billed&&role==="salesman"?" 🔒":""),t:"tel",lock:cust.billed&&role==="salesman"},{k:"fatherName",l:"Father/Mother"},{k:"address",l:"Address"},{k:"dob",l:"DOB",t:"date"},{k:"aadhar",l:"Aadhar"},{k:"pan",l:"PAN"},...(cust.billed?[{k:"nominee",l:"Nominee"},{k:"nomineeRel",l:"Nom. Relation"},{k:"exchangeName",l:"Exchanger Name"},{k:"exchangeAsked",l:"Exchange Bike Model"},{k:"exchangeRegNo",l:"Old Vehicle Reg No"},{k:"exchangeOffered",l:"Exchange Value ₹"}]:[])].map(({k,l,t,lock})=>(
                <div key={k}><label style={lbl}>{l}</label><input type={t||"text"} disabled={lock} style={{...inp,opacity:lock?0.5:1,borderColor:t==="tel"&&f[k]&&f[k].length!==10?"#ef4444":undefined}} value={f[k]||""} onChange={e=>{const v=t==="tel"?e.target.value.replace(/\D/g,"").slice(0,10):e.target.value;setF(p=>({...p,[k]:v}));}}/>{t==="tel"&&f[k]&&f[k].length>0&&f[k].length!==10&&<div style={{fontSize:10,color:"#ef4444",marginTop:2}}>⚠️ Must be 10 digits ({f[k].length} entered)</div>}{lock&&<div style={{fontSize:10,color:"#f59e0b",marginTop:3}}>Phone locked after billing — ask Manager to change</div>}</div>
              ))}
              <div><label style={lbl}>Model Code</label>
                <select style={inp} value={f.modelCode||""} onChange={e=>pickM(e.target.value)}>
                  <option value="">Select…</option>
                  {CATS.map(cat=><optgroup key={cat} label={cat}>{Object.entries(RC).filter(([,m])=>m.cat===cat).map(([code,m])=><option key={code} value={code}>{code} — {m.n}</option>)}</optgroup>)}
                </select>
                {RC[f.modelCode]&&<div style={{fontSize:11,color:"#60a5fa",marginTop:4}}>On-Road: {fc(RC[f.modelCode].onRoad)}</div>}
              </div>
              <div><label style={lbl}>Status</label><select style={inp} value={f.status} onChange={e=>setF(p=>({...p,status:e.target.value}))}>{["Hot","Warm","Cold","Booked","Lost"].map(s=><option key={s}>{s}</option>)}</select></div>
              <div><label style={lbl}>Finance/Cash</label><select style={inp} value={f.finance} onChange={e=>setF(p=>({...p,finance:e.target.value}))}><option>Cash</option><option>Finance</option></select></div>
              <div><label style={lbl}>Next Followup</label><input type="date" style={inp} value={f.followupDate||""} onChange={e=>setF(p=>({...p,followupDate:e.target.value}))}/></div>
              <div style={{display:"flex",gap:8}}><button onClick={saveEdit} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),flex:1}}>Save</button><button onClick={()=>setEdit(false)} style={{...btn("#6b8fb5","#8892a4"),flex:1}}>Cancel</button></div>
            </div>
          ):(
            <div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:10}}>
                {[["Mobile No",cust.phone||"—"],["Model",cust.model||"—"],["Code",cust.modelCode||"—"],["Finance",cust.finance||"—"],["Status",cust.status],["Source",cust.source||"—"],["Referred By",cust.refBy||"—"],["Enquiry",fd(cust.enquiryDate)],["Exp. Purchase",cust.expectedPurchaseDate?fd(cust.expectedPurchaseDate):"—"],["Followup",fd(cust.followupDate)],["Attempts",(cust.attempts||0)+"/3"],...(cust.billed?[["Exchange Bike",cust.exchangeAsked||"—"],["Exchanger",cust.exchangeName||"—"],["Old Reg No",cust.exchangeRegNo||"—"],["Exch. Value",cust.exchangeOffered?fc(cust.exchangeOffered):"—"],["Nominee",cust.nominee||"—"],["Nom. Relation",cust.nomineeRel||"—"],["DOB",fd(cust.dob)],["Father",cust.fatherName||"—"],["Aadhar",cust.aadhar||"—"]]:[])].map(([k,v])=>(
                  <div key={k} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:10,padding:"9px 11px"}}><div style={{fontSize:10,color:"#94a3b8"}}>{k}</div><div style={{fontSize:12,color:"#1e293b",fontWeight:600,marginTop:1,wordBreak:"break-all"}}>{v}</div></div>
                ))}
              </div>
              {cust.finance==="Finance"&&<div style={{background:"rgba(245,158,11,0.06)",border:"1px solid rgba(245,158,11,0.25)",borderRadius:10,padding:"10px 12px",marginBottom:10}}>
                <div style={{fontSize:10,color:"#f59e0b",fontWeight:700,marginBottom:6}}>💰 FINANCE STATUS{cust.financeBank?" — "+cust.financeBank:""}</div>
                <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>
                  {["Pending","File Login","Approved","Disbursed"].map(s=>(
                    <button key={s} onClick={()=>onUpd({financeStatus:s,remarks:(cust.remarks||"")+"\n["+td()+"] FINANCE: "+s})} style={{background:cust.financeStatus===s?"#78350f":"#f1f5f9",border:"1px solid "+(cust.financeStatus===s?"#f59e0b":"#6b8fb5"),borderRadius:14,padding:"5px 11px",fontSize:11,color:cust.financeStatus===s?"#fbbf24":"#8892a4",fontWeight:700,cursor:"pointer"}}>{s}</button>
                  ))}
                </div>
              </div>}
              {cust.expectedPurchaseDate&&<div style={{background:"rgba(139,92,246,0.06)",border:"1px solid rgba(139,92,246,0.3)",borderRadius:10,padding:"10px 12px",marginBottom:10,display:"flex",alignItems:"center",gap:10}}>
                <span style={{fontSize:18}}>📅</span>
                <div><div style={{fontSize:10,color:"#8b5cf6",fontWeight:700}}>EXPECTED DATE OF PURCHASE</div><div style={{fontSize:14,fontWeight:800,color:"#1e293b",marginTop:2}}>{fd(cust.expectedPurchaseDate)}</div></div>
              </div>}
              {r&&<div style={{background:"rgba(96,165,250,0.07)",border:"1px solid rgba(96,165,250,0.2)",borderRadius:10,padding:"10px 12px",marginBottom:10}}><div style={{fontSize:10,color:"#60a5fa",fontWeight:700,marginBottom:5}}>📊 RATE CHART</div>{[["Ex-Showroom",r.ex],["On-Road",r.onRoad],["On-Road+AMC",r.onRoad+(r.amc||0)]].map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"2px 0"}}><span style={{color:"#64748b"}}>{k}</span><span style={{color:"#1e293b",fontWeight:600}}>{fc(v)}</span></div>)}</div>}
              <button onClick={()=>setEdit(true)} style={{width:"100%",background:"#c2d6ec",border:"1px solid #6b8fb5",borderRadius:12,padding:11,color:"#64748b",fontWeight:600,fontSize:13,cursor:"pointer"}}>✏️ Edit Details</button>
            </div>
          )}
        </div>
      )}

      {tab==="history"&&(
        <div>
          <div style={{fontWeight:700,fontSize:13,color:"#1e293b",marginBottom:10}}>Interaction History</div>
          {(cust.remarks||"").trim().split("\n").filter(Boolean).map((l,i,arr)=>{
            const isLast=i===arr.length-1;
            const dm=l.match(/^\[(\d{4}-\d{2}-\d{2})\]/);
            const oc=(l.match(/^\[.*?\]\s*([A-Z_]+):/)||[])[1]||"";
            const txt=l.replace(/^\[.*?\]\s*/,"").replace(/^[A-Z_]+:\s*/,"");
            const oc_c={INTERESTED:"#22c55e",BOOKED:"#a78bfa",NO_RESPONSE:"#ef4444",SWITCHED_OFF:"#ef4444",NOT_INTERESTED:"#6b7280",PRICE_ISSUE:"#f59e0b",VISIT_SCHEDULED:"#60a5fa",REACTIVATED:"#34d399"}[oc]||"#8892a4";
            return(
              <div key={i} style={{background:isLast?"rgba(249,115,22,0.07)":"#ffffff",border:"1px solid "+(isLast?"rgba(249,115,22,0.28)":"#6b8fb5"),borderRadius:10,padding:"10px 12px",marginBottom:7,borderLeft:"3px solid "+oc_c}}>
                <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}>
                  {oc&&<span style={{fontSize:10,color:oc_c,fontWeight:700}}>{({INTERESTED:"👍",BOOKED:"✅",NO_RESPONSE:"📵",SWITCHED_OFF:"🔕",NOT_INTERESTED:"❌",PRICE_ISSUE:"💸",VISIT_SCHEDULED:"🏪",REACTIVATED:"🔄",FINANCE:"💰",MANAGER:"👔",APPROVED:"✅",BILLING:"🧾"}[oc]||"•")+" "+oc.replace(/_/g," ")}</span>}
                  <span style={{fontSize:10,color:"#94a3b8",marginLeft:"auto"}}>{fd(dm?dm[1]:"")}</span>
                </div>
                <div style={{fontSize:12,color:"#334155",lineHeight:1.5}}>{txt}</div>
                {isLast&&<div style={{fontSize:10,color:"#f97316",marginTop:3,fontWeight:600}}>← Most recent</div>}
              </div>
            );
          })}
          {(cust.callLog||[]).filter(l=>l.duration).map((l,i)=>(
            <div key={i} style={{background:"rgba(34,197,94,0.05)",border:"1px solid rgba(34,197,94,0.18)",borderRadius:8,padding:"7px 11px",marginBottom:5,display:"flex",justifyContent:"space-between"}}><span style={{fontSize:11,color:"#64748b"}}>{fd(l.date)}</span><span style={{fontSize:12,color:"#22c55e",fontWeight:600}}>📞 {l.duration}</span></div>
          ))}
        </div>
      )}

      {tab==="followup"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8,marginBottom:14}}>
            <a href={"tel:"+cust.phone} style={{background:"rgba(34,197,94,0.1)",border:"1px solid rgba(34,197,94,0.3)",borderRadius:13,padding:12,display:"flex",alignItems:"center",justifyContent:"center",gap:8,textDecoration:"none",fontWeight:700,fontSize:13}}><PhIcon s={18}/><span style={{color:"#22c55e"}}>Call Now</span></a>
            <a href={"https://wa.me/91"+cust.phone} target="_blank" rel="noreferrer" style={{background:"rgba(37,211,102,0.08)",border:"1px solid rgba(37,211,102,0.3)",borderRadius:13,padding:12,display:"flex",alignItems:"center",justifyContent:"center",gap:8,textDecoration:"none",fontWeight:700,fontSize:13}}><WAIcon s={18}/><span style={{color:"#25D366"}}>WhatsApp</span></a>
          </div>
          {!cust.stopped?<QuickLog cust={cust} onLog={onLog}/>:<div style={{background:"rgba(239,68,68,0.08)",border:"1px solid rgba(239,68,68,0.3)",borderRadius:12,padding:16,textAlign:"center"}}><div style={{fontSize:18,marginBottom:4}}>🚫</div><div style={{color:"#ef4444",fontWeight:700}}>Followup Stopped</div><div style={{color:"#64748b",fontSize:12,marginTop:4}}>3 failed attempts. Manager alerted.</div></div>}
        </div>
      )}

      {tab==="billing"&&cust.billing&&<BillingView billing={cust.billing} cust={cust} onAddPayment={onAddPayment}/>}


      {tab==="docs"&&(
        <div>
          <div style={{fontSize:11,fontWeight:700,color:"#a78bfa",marginBottom:8}}>ANYTIME (before or after billing)</div>
          <DocGrid cust={cust} onUpload={uploadPhoto} docs={[{key:"license",l:"Driving License (test ride)",ic:"🚦"},{key:"aadhar_photo",l:"Aadhar Card Photo",ic:"🪪"},{key:"pan_photo",l:"PAN Card Photo",ic:"🪪"},{key:"booking_proof",l:"Booking Payment Proof",ic:"💳"},{key:"exchange_eval",l:"Exchange Bike Evaluation (old RC + photos)",ic:"🏍️"}]}/>
          <div style={{fontSize:11,fontWeight:700,color:"#34d399",margin:"14px 0 8px"}}>AFTER BILLING</div>
          {!cust.billed&&<div style={{background:"rgba(107,114,128,0.1)",border:"1px dashed #374151",borderRadius:11,padding:"12px",fontSize:12,color:"#94a3b8",marginBottom:10}}>🔒 Unlocks after vehicle is billed</div>}
          {cust.billed&&<DocGrid cust={cust} onUpload={uploadPhoto} docs={[{key:"delivery",l:"Customer Delivery Photo",ic:"📸"},{key:"invoice",l:"Invoice Copy",ic:"🧾"},{key:"insurance",l:"Insurance Document",ic:"🛡️"},{key:"registration",l:"Registration (RC)",ic:"📋"},{key:"moneyreceipt",l:"Money Receipt Copy",ic:"🧾"},{key:"exchange",l:"Exchange Vehicle Handover Docs",ic:"🔄"}]}/>}
          {cust.billed&&cust.finance==="Finance"&&<div style={{fontSize:11,fontWeight:700,color:"#f59e0b",margin:"14px 0 8px"}}>FINANCE FILE</div>}
          {cust.billed&&cust.finance==="Finance"&&<DocGrid cust={cust} onUpload={uploadPhoto} docs={[{key:"mr",l:"MR — Finance Company",ic:"💰"},{key:"do_letter",l:"DO — Finance Company",ic:"📄"},{key:"loan_app",l:"Loan Application",ic:"📝"},{key:"sanction",l:"Sanction Letter",ic:"✅"},{key:"cust_sign",l:"Customer Signature Photo",ic:"✍️"}]}/>}
        </div>
      )}
    </div>
  );
}

function QuickLog({cust,onLog}){
  const [out,setOut]=useState("interested");
  const [note,setNote]=useState("");
  const [nxt,setNxt]=useState("");
  const [reason,setReason]=useState("");
  const [competitor,setCompetitor]=useState("");
  function submit(){
    if(out==="not_interested"&&!reason){alert("Select reason for losing customer");return;}
    onLog(cust,out,note,nxt||null,{reason,competitor});setNote("");setNxt("");setReason("");setCompetitor("");
  }
  return(
    <div style={{display:"flex",flexDirection:"column",gap:10}}>
      <div><label style={lbl}>Outcome</label><select style={inp} value={out} onChange={e=>setOut(e.target.value)}><option value="interested">Interested</option><option value="booked">Booked ✓</option><option value="visit_scheduled">Visit Scheduled</option><option value="no_response">No Response</option><option value="switched_off">Switched Off</option><option value="not_interested">Not Interested</option><option value="price_issue">Price Issue</option></select></div>
      {out==="not_interested"&&<div><label style={lbl}>Why lost? *</label><select style={inp} value={reason} onChange={e=>setReason(e.target.value)}><option value="">Select…</option><option>Price too high</option><option>Bought competitor brand</option><option>Finance rejected</option><option>Postponed purchase</option><option>Family decision</option><option>Other</option></select></div>}
      {out==="not_interested"&&reason==="Bought competitor brand"&&<div><label style={lbl}>Which brand/model?</label><input style={inp} value={competitor} onChange={e=>setCompetitor(e.target.value)}/></div>}
      <div><label style={lbl}>Notes</label><textarea rows={3} style={{...inp,resize:"none"}} value={note} onChange={e=>setNote(e.target.value)} placeholder="What did customer say?"/></div>
      <div><label style={lbl}>Next followup (blank=auto)</label><input type="date" style={inp} value={nxt} onChange={e=>setNxt(e.target.value)}/></div>
      <button onClick={submit} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),padding:13,fontSize:14}}>Save Followup</button>
    </div>
  );
}

function BookingModal({cust,onClose,onSave}){
  const [amt,setAmt]=useState(cust.booking?cust.booking.amt:"");
  const [date,setDate]=useState(cust.booking?cust.booking.date:td());
  const [mode,setMode]=useState(cust.booking?cust.booking.mode||"Cash":"Cash");
  const [note,setNote]=useState("");
  const [proof,setProof]=useState(null);
  function pickProof(file){compressImg(file,setProof);}
  function submit(){
    if(!amt||Number(amt)<=0){alert("Enter booking amount");return;}
    onSave({amt:Number(amt),date,mode,note,proof});
  }
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:150,display:"flex",alignItems:"flex-end"}}>
      <div style={{background:"#ffffff",width:"100%",borderRadius:"20px 20px 0 0",padding:"20px 16px 44px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <div style={{fontWeight:800,fontSize:17,color:"#1e293b"}}>Take Booking</div>
          <button onClick={onClose} style={{background:"#c2d6ec",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",color:"#64748b",fontSize:18}}>✕</button>
        </div>
        <div style={{fontSize:12,color:"#94a3b8",marginBottom:14}}>{cust.name} · {cust.model}</div>
        <div style={{background:"rgba(139,92,246,0.08)",border:"1px solid rgba(139,92,246,0.25)",borderRadius:10,padding:"9px 12px",marginBottom:14,fontSize:11,color:"#a78bfa"}}>Records booking without KYC documents. Amount &amp; date auto-fill in the calculation sheet at billing time.</div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          <div><label style={lbl}>Booking Amount ₹ *</label><input type="number" style={inp} value={amt} onChange={e=>setAmt(e.target.value)}/></div>
          <div><label style={lbl}>Booking Date *</label><input type="date" style={inp} value={date} onChange={e=>setDate(e.target.value)}/></div>
          <div><label style={lbl}>Mode</label><div style={{display:"flex",gap:7}}>{["Cash","UPI","Cheque"].map(m=><button key={m} onClick={()=>setMode(m)} style={{flex:1,background:mode===m?"#dbeafe":"#6b8fb5",border:"1px solid "+(mode===m?"#3b82f6":"#6b8fb5"),borderRadius:10,padding:10,color:mode===m?"#60a5fa":"#5a6478",fontWeight:700,cursor:"pointer",fontSize:12}}>{m}</button>)}</div></div>
          <div><label style={lbl}>Notes</label><input style={inp} value={note} onChange={e=>setNote(e.target.value)} placeholder="e.g. balance on delivery"/></div>
          <div><label style={lbl}>Payment Proof (cheque / UPI screenshot)</label>
            {proof&&<img src={proof} alt="proof" style={{width:"100%",maxHeight:140,objectFit:"contain",background:"#000",borderRadius:9,marginBottom:6}}/>}
            <input type="file" accept="image/*" onChange={e=>{if(e.target.files&&e.target.files[0]){pickProof(e.target.files[0]);e.target.value="";}}} style={{width:"100%",background:"#c2d6ec",borderRadius:9,padding:8,fontSize:12,color:"#64748b",border:"1px dashed #2a3040"}}/>
          </div>
          <div style={{display:"flex",gap:8}}>
            <button onClick={onClose} style={{...btn("#6b8fb5","#8892a4"),flex:1,padding:14,borderRadius:13}}>← Go Back</button>
            <button onClick={submit} style={{...btn("linear-gradient(135deg,#8b5cf6,#6d28d9)"),flex:2,padding:14,fontSize:15,borderRadius:13}}>Save Booking</button>
          </div>
        </div>
      </div>
    </div>
  );
}

function AddModal({onClose,onSave,curUser,role,existing}){
  const [f,setF]=useState({name:"",phone:"",fatherName:"",address:"",dob:"",aadhar:"",pan:"",modelCode:"",model:"",cat:"",enquiryDate:td(),status:"Hot",salesman:curUser,finance:"Cash",exchangeAsked:"",exchangeOffered:"",remarks:"",followupDate:"",nominee:"",nomineeRel:""});
  const [mSearch,setMSearch]=useState("");
  function cap(v){return String(v||"").toUpperCase();}
  function capBlur(k){return{onBlur:e=>setF(p=>({...p,[k]:cap(e.target.value)}))};}
  function pickM(code){const m=RC[code];setF(p=>({...p,modelCode:code,model:m?m.n:"",cat:m?m.cat:""}));setMSearch(code?(code+" — "+(RC[code]?RC[code].n:"")):"");}
  function submit(){
    if(!f.name||!f.phone){alert("Name & phone required");return;}
    if(!/^\d{10}$/.test(f.phone)){alert("Phone must be exactly 10 digits");return;}
    if(f.enquiryDate>td()){alert("Enquiry date cannot be in the future");return;}
    const dup=(existing||[]).find(c=>c.phone===f.phone);
    if(dup){alert("⚠️ Duplicate! "+dup.name+" already exists with this number (handled by "+dup.salesman+")");return;}
    const st=f.status||"Hot";
    onSave({...f,followupDate:f.followupDate||aD(f.enquiryDate||td(),FU[st]||3),callLog:[]});
  }
  // DOB age check
  const dobAge=f.dob?(Math.floor((new Date()-new Date(f.dob))/31557600000)):null;
  const allModels=Object.entries(RC);
  const filteredModels=mSearch?allModels.filter(([code,m])=>code.toLowerCase().includes(mSearch.toLowerCase())||m.n.toLowerCase().includes(mSearch.toLowerCase())):allModels;
  const r=RC[f.modelCode];
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.88)",zIndex:150,display:"flex",alignItems:"flex-end"}}>
      <div style={{background:"#ffffff",width:"100%",borderRadius:"20px 20px 0 0",maxHeight:"94vh",overflowY:"auto",padding:"20px 16px 44px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}><div style={{fontWeight:800,fontSize:17,color:"#1e293b"}}>New Enquiry</div><button onClick={onClose} style={{background:"#c2d6ec",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",color:"#64748b",fontSize:18}}>✕</button></div>
        <div style={{display:"flex",flexDirection:"column",gap:10}}>
          {[{k:"name",l:"Customer Name *"},{k:"phone",l:"Phone *",t:"tel"},{k:"address",l:"Address"}].map(({k,l,t})=>(
            <div key={k}><label style={lbl}>{l}</label><input type={t||"text"} style={t!=="tel"?{...inp,textTransform:"uppercase"}:{...inp,borderColor:f[k]&&f[k].length!==10?"#ef4444":undefined}} value={f[k]||""} onChange={e=>{const v=t==="tel"?e.target.value.replace(/\D/g,"").slice(0,10):e.target.value;setF(p=>({...p,[k]:v}));}} {...(t!=="tel"?capBlur(k):{})}/>{t==="tel"&&f[k]&&f[k].length>0&&f[k].length!==10&&<div style={{fontSize:10,color:"#ef4444",marginTop:2}}>⚠️ Must be 10 digits ({f[k].length} entered)</div>}</div>
          ))}
          <div><label style={lbl}>Model — search by code or name</label>
            <input style={inp} value={mSearch} onChange={e=>{setMSearch(e.target.value);if(!e.target.value){pickM("");}}} placeholder="Type model code or name…" list="model-list"/>
            <datalist id="model-list">{filteredModels.map(([code,m])=><option key={code} value={code+" — "+m.n}/>)}</datalist>
            {mSearch&&filteredModels.length===1&&filteredModels[0][0]!==f.modelCode&&(()=>{const [code]=filteredModels[0];setTimeout(()=>pickM(code),0);return null;})()}
            {/* also support selecting from the datalist */}
            {mSearch&&(()=>{const match=allModels.find(([c,m])=>mSearch===c+" — "+m.n||mSearch===c);if(match&&match[0]!==f.modelCode)pickM(match[0]);return null;})()}
            {r&&<div style={{background:"rgba(96,165,250,0.07)",border:"1px solid rgba(96,165,250,0.2)",borderRadius:8,padding:"8px 10px",marginTop:5,fontSize:11}}><span style={{color:"#60a5fa",fontWeight:700}}>{f.modelCode} — {r.n}</span><br/><span style={{color:"#64748b"}}>Ex-Showroom: </span><b style={{color:"#1e293b"}}>{fc(r.ex)}</b><span style={{color:"#64748b"}}> | On-Road: </span><b style={{color:"#34d399"}}>{fc(r.onRoad)}</b></div>}
          </div>
          <div><label style={lbl}>DOB</label>
            <input type="date" style={inp} value={f.dob||""} onChange={e=>setF(p=>({...p,dob:e.target.value}))} max={td()}/>
            {dobAge!==null&&dobAge<18&&<div style={{background:"rgba(239,68,68,0.1)",border:"1px solid #ef4444",borderRadius:7,padding:"5px 10px",marginTop:4,fontSize:11,color:"#ef4444",fontWeight:700}}>⚠️ Under 18 years — customer is a minor ({dobAge} yrs)</div>}
          </div>
          <div><label style={lbl}>Lead Source</label><select style={inp} value={f.source||""} onChange={e=>setF(p=>({...p,source:e.target.value}))}><option value="">Select…</option><option>Walk-in</option><option>Via Phone Call</option><option>Reference</option><option>Old Customer</option><option>JustDial</option><option>Instagram/Facebook</option><option>Hoarding/Newspaper</option><option>Other</option></select></div>
          {f.source==="Reference"&&<div><label style={lbl}>Referred By (name / phone)</label><input style={{...inp,textTransform:"uppercase"}} value={f.refBy||""} onChange={e=>setF(p=>({...p,refBy:e.target.value}))} {...capBlur("refBy")}/></div>}
          <div><label style={lbl}>Enquiry Date</label><input type="date" style={inp} value={f.enquiryDate} max={td()} onChange={e=>setF(p=>({...p,enquiryDate:e.target.value}))}/></div>
          <div><label style={lbl}>Temperature</label><div style={{display:"flex",gap:7}}>{["Hot","Warm","Cold"].map(s=><button key={s} onClick={()=>setF(p=>({...p,status:s}))} style={{flex:1,background:f.status===s?ST_C[s]:"#6b8fb5",border:"none",borderRadius:10,padding:10,color:f.status===s?"#fff":"#5a6478",fontWeight:700,cursor:"pointer",fontSize:12}}>{s}</button>)}</div></div>
          <div><label style={lbl}>Mode</label><div style={{display:"flex",gap:7}}>{["Cash","Finance"].map(s=><button key={s} onClick={()=>setF(p=>({...p,finance:s}))} style={{flex:1,background:f.finance===s?"#dbeafe":"#6b8fb5",border:"1px solid "+(f.finance===s?"#3b82f6":"#6b8fb5"),borderRadius:10,padding:10,color:f.finance===s?"#60a5fa":"#5a6478",fontWeight:700,cursor:"pointer",fontSize:12}}>{s}</button>)}</div></div>
          <div><label style={lbl}>Expected Date of Purchase</label><input type="date" style={inp} value={f.expectedPurchaseDate||""} onChange={e=>setF(p=>({...p,expectedPurchaseDate:e.target.value}))}/></div>
          <div style={{background:"rgba(107,114,128,0.08)",border:"1px dashed #374151",borderRadius:10,padding:"9px 12px",fontSize:11,color:"#94a3b8"}}>🔒 Nominee &amp; Exchange details are entered at billing time</div>
          {role!=="salesman"&&<div><label style={lbl}>Assign to</label><select style={inp} value={f.salesman} onChange={e=>setF(p=>({...p,salesman:e.target.value}))}>{SM.map(s=><option key={s}>{s}</option>)}</select></div>}
          <div><label style={lbl}>Remarks</label><textarea rows={2} style={{...inp,resize:"none",textTransform:"uppercase"}} value={f.remarks} onChange={e=>setF(p=>({...p,remarks:e.target.value}))} onBlur={e=>setF(p=>({...p,remarks:cap(e.target.value)}))}/></div>
          <div><label style={lbl}>Followup Date (blank=auto)</label><input type="date" style={inp} value={f.followupDate} onChange={e=>setF(p=>({...p,followupDate:e.target.value}))}/></div>
          <button onClick={submit} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),padding:15,fontSize:15,borderRadius:14,marginTop:4}}>Add Customer</button>
        </div>
      </div>
    </div>
  );
}

function Row({label,val,auto}){
  return(
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"4px 0",borderBottom:"1px solid #131820"}}>
      <div style={{display:"flex",alignItems:"center",gap:5}}><span style={{fontSize:12,color:"#64748b"}}>{label}</span>{auto&&<span style={{fontSize:9,background:"#c2d6ec",color:"#94a3b8",padding:"1px 5px",borderRadius:4}}>AUTO</span>}</div>
      <span style={{fontSize:13,color:"#1e293b",fontWeight:600}}>{fc(val)}</span>
    </div>
  );
}
function Inp({label,k,f,setF}){
  return(
    <div style={{display:"flex",alignItems:"center",padding:"4px 0",borderBottom:"1px solid #131820",gap:8}}>
      <span style={{fontSize:12,color:"#64748b",flex:1}}>{label}</span>
      <input type="number" inputMode="numeric" min="0" value={f[k]||""} onChange={e=>{const v=e.target.value.replace(/[^0-9]/g,"");setF(p=>({...p,[k]:v}));}} placeholder="0" style={{width:88,background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 8px",fontSize:12,color:"#1e293b",textAlign:"right"}}/>
    </div>
  );
}
function Tot({label,val,col}){
  return(
    <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",background:"rgba(255,255,255,0.04)",borderRadius:8,padding:"8px 10px",marginTop:3}}>
      <span style={{fontSize:12,color:"#1e293b",fontWeight:700}}>{label}</span>
      <span style={{fontSize:17,color:col||"#f97316",fontWeight:900}}>{fc(val)}</span>
    </div>
  );
}
function BillingModal({cust,onClose,onSave,onDraft,notify,role,stockData,billedChassis}){
  const r=RC[cust.modelCode]||{};
  const isFin=cust.finance==="Finance";
  const eb=cust.billing||cust.billingDraft||{};
  const [f,setF]=useState({...eb,billName:eb.billName||cust.name,exchName:cust.exchangeName||"",exchPhone:cust.exchangePhone||eb.details?.exchangePhone||"",exchModel:cust.exchangeAsked||"",exchRegNo:cust.exchangeRegNo||"",bkDate:(cust.booking&&cust.booking.date)||td(),fatherName:cust.fatherName||"",dob:cust.dob||"",aadhar:cust.aadhar||"",pan:cust.pan||"",nominee:cust.nominee||"",nomineeRel:cust.nomineeRel||"",hdl:eb.hdl!==undefined?eb.hdl:(r.hdl||600),ins:eb.ins!==undefined?eb.ins:(r.ins||0),reg:eb.reg!==undefined?eb.reg:(r.reg||0),acc:0,tef:isFin?500:0,hyp:isFin?500:0,addAmc:false,cof:0,sdis:0,corp:0,bk:(cust.booking&&cust.booking.amt)||0,exv:eb.exv!==undefined?eb.exv:0,loan:0,payments:eb.payments&&eb.payments.length?eb.payments:(eb.paid||eb.payMode?[{mode:eb.payMode||"Cash",amt:Number(eb.paid||0),ref:""}]:[{mode:"Cash",amt:0,ref:""}]),chassis:"",engine:"",color:"",deliveryDate:td(),financeBank:"",registrationNo:"",insuranceNo:""});
  const c=calcB(f,r);
  const [chk,setChk]=useState(eb.checklist||{pdi:false,helmet:false,docs:false,service:false});
  const VER_ALL=[["nameV","Customer name verified"],["fatherV","Father name verified"],["aadharV","Aadhar number verified"],["nomineeV","Nominee & relation added"],["chassisV","Chassis number verified"],["engineV","Engine number verified"],["colorV","Colour verified"]];
  const [ver,setVer]=useState(eb.verify||{});
  const [chassisCustom,setChassisCustom]=useState("");
  useEffect(()=>{setVer(p=>({...p,
    nameV:p.nameV||!!f.billName,fatherV:p.fatherV||!!f.fatherName,
    aadharV:p.aadharV||!!f.aadhar,nomineeV:p.nomineeV||!!(f.nominee&&f.nomineeRel),
    chassisV:p.chassisV||!!f.chassis,engineV:p.engineV||!!f.engine,colorV:p.colorV||!!f.color}));},
    [f.billName,f.fatherName,f.aadhar,f.nominee,f.nomineeRel,f.chassis,f.engine,f.color]);
  const sRows=stockData||[];
  const sKeys=sRows.length>0?Object.keys(sRows[0]):[];
  const sChassisKey=findStockCol(sKeys,["chassis","frame"]);
  const sEngineKey=findStockCol(sKeys,["engine"]);
  const sColorKey=findStockCol(sKeys,["color","colour"]);
  const sModelKey=findStockCol(sKeys,["model","name","variant","description","item","product","vehicle"]);
  const sAgeKey=findStockCol(sKeys,["age","days","ageing","aging"]);
  const sDateKey=findStockCol(sKeys,["date","invoice","inward","receipt","stock"]);
  const sBranchKey=findStockCol(sKeys,["branch","location","godown","showroom","store","place"]);
  const modelStr=(cust.model||"").toLowerCase();
  const modelCode=(cust.modelCode||"").toLowerCase();
  function rowMatchesModel(row){
    const allVals=Object.values(row).map(v=>String(v||"").toLowerCase()).join(" ");
    const words=modelStr.split(" ").filter(w=>w.length>2);
    return allVals.includes(modelCode)||words.filter(w=>allVals.includes(w)).length>=2;
  }
  const availableForModel=sRows.filter(row=>{
    return rowMatchesModel(row)&&!(billedChassis||[]).includes(String(row[sChassisKey]||"").trim().toUpperCase());
  }).sort((a,b)=>{
    function getAge(row){
      if(sAgeKey&&row[sAgeKey]!==undefined&&row[sAgeKey]!=="")return Number(row[sAgeKey])||0;
      if(sDateKey&&row[sDateKey]){const raw=row[sDateKey];let dt=typeof raw==="number"&&raw>40000?new Date(Math.round((raw-25569)*86400000)):new Date(raw);if(dt&&!isNaN(dt))return Math.floor((Date.now()-dt)/86400000);}
      return 0;
    }
    return getAge(b)-getAge(a);
  });
  function pickChassis(chassisVal){
    const row=availableForModel.find(r=>String(r[sChassisKey]||"")===chassisVal);
    setF(p=>({...p,chassis:chassisVal,engine:row&&sEngineKey?String(row[sEngineKey]||""):p.engine,color:row&&sColorKey?String(row[sColorKey]||""):p.color}));
  }

  function buildReceipt(){
    var rows="<div class=row><span>Date</span><span class=v>"+td()+"</span></div>";
    rows+="<div class=row><span>Customer</span><span class=v>"+(f.billName||cust.name)+"</span></div>";
    rows+="<div class=row><span>Father/Mother</span><span class=v>"+(f.fatherName||"—")+"</span></div>";rows+="<div class=row><span>Nominee</span><span class=v>"+(f.nominee||"—")+" ("+(f.nomineeRel||"—")+")</span></div>";rows+="<div class=row><span>Aadhar</span><span class=v>"+(f.aadhar||"—")+"</span></div>";
    rows+="<div class=row><span>Phone</span><span class=v>"+cust.phone+"</span></div>";
    rows+="<div class=row><span>Model</span><span class=v>"+cust.model+"</span></div>";
    rows+="<div class=row><span>Code</span><span class=v>"+cust.modelCode+"</span></div>";
    rows+="<div class=row><span>Colour</span><span class=v>"+f.color+"</span></div>";
    rows+="<div class=row><span>Chassis</span><span class=v>"+f.chassis+"</span></div>";
    rows+="<div class=row><span>Engine</span><span class=v>"+f.engine+"</span></div>";
    rows+="<div class=row><span>Delivery</span><span class=v>"+fd(f.deliveryDate)+"</span></div>";
    rows+="<div class=row><span>Mode</span><span class=v>"+cust.finance+"</span></div>";
    if(f.financeBank)rows+="<div class=row><span>Finance Bank</span><span class=v>"+f.financeBank+"</span></div>";
    if(f.registrationNo)rows+="<div class=row><span>Reg No</span><span class=v>"+f.registrationNo+"</span></div>";
    rows+="<hr/>";
    if(f.mrNo)rows+="<div class=row><span>MR No.</span><span class=v>"+f.mrNo+"</span></div>";
    var activePmts=(f.payments||[]).filter(p=>Number(p.amt||0)>0);
    if(activePmts.length>1){activePmts.forEach(p=>{rows+="<div class=row><span>"+p.mode+(p.ref?" ("+p.ref+")":"")+"</span><span class=v>"+fc(Number(p.amt))+"</span></div>";});}else if(activePmts.length===1){rows+="<div class=row><span>Payment Mode</span><span class=v>"+activePmts[0].mode+(activePmts[0].ref?" ("+activePmts[0].ref+")":"")+"</span></div>";}
    var CALCROWS="";
    CALCROWS+="<div class=row><span>Ex-Showroom (A)</span><span class=v>"+fc(c.ex)+"</span></div>";
    if(c.ca)CALCROWS+="<div class=row><span>+ Comp.Acc</span><span>"+fc(c.ca)+"</span></div>";
    CALCROWS+="<div class=row><span>+ Handling</span><span>"+fc(c.hdl)+"</span></div>";
    CALCROWS+="<div class=row><span>+ Insurance 5yr</span><span>"+fc(c.ins)+"</span></div>";
    if(c.reg)CALCROWS+="<div class=row><span>+ Registration</span><span>"+fc(c.reg)+"</span></div>";
    if(c.acc)CALCROWS+="<div class=row><span>+ Accessories</span><span>"+fc(c.acc)+"</span></div>";
    if(c.tef)CALCROWS+="<div class=row><span>+ Teflon</span><span>"+fc(c.tef)+"</span></div>";
    if(c.hyp)CALCROWS+="<div class=row><span>+ Hypothication</span><span>"+fc(c.hyp)+"</span></div>";
    if(c.amcV)CALCROWS+="<div class=row><span>+ AMC</span><span>"+fc(c.amcV)+"</span></div>";
    CALCROWS+="<div class=row><b>Total C=A+B</b><b>"+fc(c.C)+"</b></div>";
    if(c.cof)CALCROWS+="<div class=row><span>- Consumer Offer</span><span>"+fc(c.cof)+"</span></div>";
    if(c.sdis)CALCROWS+="<div class=row><span>- Special Discount</span><span>"+fc(c.sdis)+"</span></div>";
    CALCROWS+="<div class=row><b>Deal Price E</b><b>"+fc(c.E)+"</b></div>";
    if(c.bk)rows+="<div class=row><span>- Booking ("+fd(f.bkDate)+")</span><span>"+fc(c.bk)+"</span></div>";
    if(c.exv)rows+="<div class=row><span>- Exchange: "+(f.exchModel||"")+" "+(f.exchRegNo||"")+(f.exchName?" ("+f.exchName+")":"")+"</span><span>"+fc(c.exv)+"</span></div>";
    if(c.loan)rows+="<div class=row><span>- Loan</span><span>"+fc(c.loan)+"</span></div>";
    rows+="<div class=total><span>AMOUNT RECEIVED</span><span>"+fc(c.paid)+"</span></div>";
    rows+="<div class=row><span>Balance</span><span style='color:"+(c.K>0?"red":"green")+"'>"+fc(Math.max(c.K,0))+"</span></div>";
    window.__CG=CALCROWS;
    return "<!DOCTYPE html><html><head><title>Money Receipt</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial,sans-serif;font-size:13px;padding:20px;color:#111}.logo{font-size:24px;font-weight:900;letter-spacing:2px;text-align:center}.hdr{text-align:center;border-bottom:2px solid #000;padding-bottom:10px;margin-bottom:14px}.hdr p{font-size:11px;color:#444;margin-top:2px}h2{text-align:center;font-size:17px;margin:10px 0 14px}.row{display:flex;justify-content:space-between;padding:4px 0;border-bottom:1px solid #eee}.v{font-weight:700}.total{display:flex;justify-content:space-between;padding:10px 0;border-top:2px solid #000;font-size:16px;font-weight:900}.sigs{display:flex;justify-content:space-between;margin-top:40px}.sigs div{text-align:center;font-size:11px}</style></head><body><div class=hdr><div class=logo>NKD BAJAJ</div><p>Authorised Main Dealer — Bajaj Auto Ltd.</p><p>Hirak Road, Near Kamal Katesaria School, Dhanbad</p><p>Ph: 7033099006 | info@nkdbajaj.com</p></div><h2>MONEY RECEIPT</h2>"+rows+"<div class=sigs><div>____________________<br/>Customer Sign</div><div>____________________<br/>For NKD Bajaj</div></div><p style='text-align:center;font-size:10px;margin-top:16px;color:#666'>Subject to realization of Cheque/Draft</p></body></html>";
  }
  function openReceipt(html){
    var w=window.open("","_blank","width=420,height=700");
    if(!w){notify("Popup blocked — receipt saved in Billing tab","warn");return;}
    w.document.write(html);w.document.close();setTimeout(function(){w.print();},400);
  }

  const [busy,setBusy]=useState(false);
  function saveDraft(){
    onDraft({...f,payMode:(f.payments||[]).filter(p=>Number(p.amt||0)>0).map(p=>p.mode).join(" + ")||"—",paid:c.paid,calc:c,checklist:chk,verify:ver,details:{name:f.billName||cust.name,exchangeName:f.exchName,exchangePhone:f.exchPhone,exchangeAsked:f.exchModel,exchangeRegNo:f.exchRegNo,exchangeOffered:String(f.exv||""),fatherName:f.fatherName,address:cust.address,dob:f.dob,nominee:f.nominee,nomineeRel:f.nomineeRel,aadhar:f.aadhar,pan:f.pan}});
    notify("📋 Draft saved — reopen billing to continue");
    onClose();
  }
  function submit(){
    if(busy)return;
    if(!f.chassis){notify("Enter chassis number","err");return;}
    if(!f.aadhar||!f.fatherName||!f.nominee||!f.nomineeRel){notify("Fill KYC: Aadhar, Father name, Nominee & Relation","err");return;}
    if(c.C<0||c.E<0||c.G<0||c.I<0){alert("⚠️ Calculation error — a total has gone NEGATIVE.\nCheck discounts/booking/exchange amounts. No value can exceed the price above it.");return;}
    if(c.K<0){alert("⚠️ Payment Received ("+fc(c.paid)+") is MORE than balance due ("+fc(c.I)+").\nCorrect the Payment Received amount.");return;}
    const missing=VER_ALL.filter(([k])=>!ver[k]);
    if(missing.length>0){alert("⚠️ Cannot submit — verify these first:\n\n"+missing.map(([,l])=>"☐ "+l).join("\n"));return;}
    var html=buildReceipt();
    var CALC_G=window.__CG||"";
    var activePmtsS=(f.payments||[]).filter(p=>Number(p.amt||0)>0);
    var payModeSummary=activePmtsS.map(p=>p.mode+(p.ref?" ("+p.ref+")":"")).join(" + ")||"—";
    var payModeRows=activePmtsS.map(p=>"<div class=row><span>"+p.mode+(p.ref?" ("+p.ref+")":"")+"</span><span class=v>"+fc(Number(p.amt))+"</span></div>").join("");
    var calcHtml=html.replace("MONEY RECEIPT","CALCULATION SHEET (INTERNAL)").replace("</h2>","</h2>"+[payModeRows||("<div class=row><span>Payment Mode</span><span class=v>—</span></div>"),"<div class=row><span>MR No.</span><span class=v>"+(f.mrNo||"—")+"</span></div>","<div class=row><span>Financed By</span><span class=v>"+(f.financeBank||"Cash")+"</span></div>"].join("")+CALC_G);
    onSave({...f,payMode:payModeSummary,paid:c.paid,calc:c,calcHtml:calcHtml,checklist:chk,verify:ver,verifyList:VER_ALL.map(([k,l])=>[k,l]),receiptHtml:html,details:{name:f.billName||cust.name,exchangeName:f.exchName,exchangePhone:f.exchPhone,exchangeAsked:f.exchModel,exchangeRegNo:f.exchRegNo,exchangeOffered:String(f.exv||""),fatherName:f.fatherName,address:cust.address,dob:f.dob,nominee:f.nominee,nomineeRel:f.nomineeRel,aadhar:f.aadhar,pan:f.pan}});
    setBusy(true);
    notify(role==="salesman"?"✅ Sent to Manager for approval — receipt saved in Billing tab":"✅ Billed & approved — receipt saved in Billing tab");
  }

  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.93)",zIndex:150,display:"flex",alignItems:"flex-end"}}>
      <div style={{background:"#ffffff",width:"100%",borderRadius:"20px 20px 0 0",maxHeight:"97vh",overflowY:"auto",padding:"20px 16px 44px"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}><div style={{fontWeight:800,fontSize:16,color:"#1e293b"}}>Bill Vehicle</div><button onClick={onClose} style={{background:"#c2d6ec",border:"none",borderRadius:8,width:32,height:32,cursor:"pointer",color:"#64748b",fontSize:18}}>✕</button></div>
        <div style={{fontSize:12,color:"#94a3b8",marginBottom:14}}>{cust.name} · {cust.model} · <span style={{color:"#60a5fa",fontWeight:600}}>{cust.modelCode}</span></div>
        <div style={{background:"rgba(96,165,250,0.07)",border:"1px solid rgba(96,165,250,0.22)",borderRadius:11,padding:"9px 12px",marginBottom:14,fontSize:11,color:"#64748b",lineHeight:1.7}}>✓ <b style={{color:"#60a5fa"}}>Auto-filled:</b> Ex-Showroom {fc(r.ex)}, Insurance {fc(r.ins||0)}, Registration {fc(r.reg||0)}. Fill <b style={{color:"#f97316"}}>Teflon/Accessories/Hypo</b> manually.</div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#475569",letterSpacing:0.8,marginBottom:6}}>CUSTOMER KYC (required for billing)</div>
          <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              {[{k:"billName",l:"Customer Name (as on bill) *"},{k:"fatherName",l:"Father/Mother Name *"},{k:"dob",l:"Date of Birth",t:"date"},{k:"aadhar",l:"Aadhar No *"},{k:"pan",l:"PAN No"},{k:"nominee",l:"Nominee Name *"},{k:"nomineeRel",l:"Nominee Relation *"}].map(({k,l,t})=>{
                const age=k==="dob"&&f.dob?Math.floor((new Date()-new Date(f.dob))/31557600000):null;
                return(<div key={k}><label style={{...lbl,fontSize:10}}>{l}</label><input type={t||"text"} value={f[k]||""} max={t==="date"?td():undefined} onChange={e=>setF(p=>({...p,[k]:e.target.value}))} onBlur={t!=="date"?e=>setF(p=>({...p,[k]:String(e.target.value).toUpperCase()})):undefined} style={{...inp,fontSize:12,padding:"8px 10px",textTransform:t==="date"?"none":"uppercase"}}/>
                {age!==null&&age<18&&<div style={{background:"rgba(239,68,68,0.1)",border:"1px solid #ef4444",borderRadius:7,padding:"4px 8px",marginTop:3,fontSize:11,color:"#ef4444",fontWeight:700}}>⚠️ Under 18 — minor ({age} yrs)</div>}</div>);
              })}
            </div>
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#f59e0b",letterSpacing:0.8,marginBottom:6}}>EXCHANGE DETAILS (if old bike exchanged)</div>
          <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              {[{k:"exchName",l:"Exchanger Name"},{k:"exchPhone",l:"Exchanger Mobile",t:"tel"},{k:"exchModel",l:"Old Bike Model"},{k:"exchRegNo",l:"Old Vehicle Reg No"}].map(({k,l,t})=>(
                <div key={k}><label style={{...lbl,fontSize:10}}>{l}</label><input type={t||"text"} inputMode={t==="tel"?"numeric":undefined} value={f[k]||""} onChange={e=>{const v=t==="tel"?e.target.value.replace(/\D/g,"").slice(0,10):e.target.value;setF(p=>({...p,[k]:v}));}} onBlur={t!=="tel"?e=>setF(p=>({...p,[k]:String(e.target.value).toUpperCase()})):undefined} style={{...inp,fontSize:12,padding:"8px 10px",textTransform:t==="tel"?"none":"uppercase",borderColor:t==="tel"&&f[k]&&f[k].length!==10?"#ef4444":undefined}}/>{t==="tel"&&f[k]&&f[k].length>0&&f[k].length!==10&&<div style={{fontSize:10,color:"#ef4444",marginTop:2}}>⚠️ Must be 10 digits</div>}</div>
              ))}
              <div><label style={{...lbl,fontSize:10}}>Exchange Value ₹</label><input type="number" inputMode="numeric" value={f.exv||""} onChange={e=>setF(p=>({...p,exv:e.target.value}))} style={{...inp,fontSize:12,padding:"8px 10px"}}/></div>
            </div>
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#475569",letterSpacing:0.8,marginBottom:6}}>VEHICLE DETAILS</div>
          <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12}}>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
              <div style={{gridColumn:"1/-1"}}>
                <label style={{...lbl,fontSize:10}}>Chassis No * {availableForModel.length>0&&<span style={{color:"#34d399",fontWeight:700}}>({availableForModel.length} in stock)</span>}{availableForModel.length===0&&sRows.length>0&&<span style={{color:"#f97316",fontWeight:700}}>(no stock for this model)</span>}</label>
                <input list="chassis-list" value={f.chassis||""} onChange={e=>pickChassis(e.target.value)} placeholder="Type or search chassis no…" style={{...inp,fontSize:12,padding:"8px 10px",textTransform:"uppercase"}} onBlur={e=>setF(p=>({...p,chassis:String(e.target.value).toUpperCase()}))}/>
                <datalist id="chassis-list">{availableForModel.map((row,i)=>{
                  const ch=String(row[sChassisKey]||"");
                  const col=sColorKey?String(row[sColorKey]||""):"";
                  let age="";
                  if(sAgeKey&&row[sAgeKey]!==undefined&&row[sAgeKey]!==null&&row[sAgeKey]!==""){age=String(Math.round(Number(row[sAgeKey])))+" days";}
                  else if(sDateKey&&row[sDateKey]){
                    const raw=row[sDateKey];
                    let dt=null;
                    if(typeof raw==="number"&&raw>40000){dt=new Date(Math.round((raw-25569)*86400000));}// Excel serial
                    else{dt=new Date(raw);}
                    if(dt&&!isNaN(dt)){age=Math.max(0,Math.floor((Date.now()-dt)/86400000))+" days";}
                  }
                  const br=sBranchKey?cleanBranch(row[sBranchKey]):"";
                  const label=[col,age,br].filter(Boolean).join(" · ");
                  return(<option key={i} value={ch}>{label}</option>);
                })}</datalist>
              </div>
              {[{k:"engine",l:"Engine No"},{k:"color",l:"Colour"},{k:"deliveryDate",l:"Delivery Date",t:"date"},...(isFin?[{k:"financeBank",l:"Finance Bank"}]:[]),{k:"registrationNo",l:"Reg No"},{k:"insuranceNo",l:"Insurance No"},{k:"mrNo",l:"MR No."}].map(({k,l,t})=>(
                <div key={k}><label style={{...lbl,fontSize:10}}>{l}</label><input type={t||"text"} value={f[k]||""} onChange={e=>setF(p=>({...p,[k]:e.target.value}))} onBlur={t!=="date"?e=>setF(p=>({...p,[k]:String(e.target.value).toUpperCase()})):undefined} style={{...inp,fontSize:12,padding:"8px 10px",textTransform:t==="date"?"none":"uppercase"}}/></div>
              ))}
            </div>
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#f97316",letterSpacing:0.8,marginBottom:6}}>CALCULATION SHEET</div>
          <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12}}>
            <Row label="(A) Ex-Showroom" val={c.ex} auto={true}/>
            <Row label="+ Comp. Accessories" val={c.ca} auto={true}/>
            <Inp label="+ Handling" k="hdl" f={f} setF={setF}/>
            <Inp label="+ Insurance (5yr)" k="ins" f={f} setF={setF}/>
            <Inp label="+ Registration" k="reg" f={f} setF={setF}/>
            <div style={{background:"rgba(249,115,22,0.06)",borderRadius:7,padding:"6px 8px",margin:"6px 0 2px"}}><div style={{fontSize:9,color:"#f97316",fontWeight:700,marginBottom:4}}>MANUAL ENTRY</div>
              <Inp label="+ Accessories" k="acc" f={f} setF={setF}/>
              <Inp label="+ Teflon Coating" k="tef" f={f} setF={setF}/>
              <Inp label="+ Hypothication" k="hyp" f={f} setF={setF}/>
            </div>
            <div style={{display:"flex",alignItems:"center",gap:8,padding:"6px 0",borderBottom:"1px solid #131820"}}><input type="checkbox" id="amc" checked={f.addAmc} onChange={e=>setF(p=>({...p,addAmc:e.target.checked}))}/><label htmlFor="amc" style={{fontSize:12,color:"#64748b",cursor:"pointer"}}>Add AMC ({fc(r.amc||0)}/yr)</label></div>
            <Tot label="TOTAL C = A + B" val={c.C} col="#f97316"/>
            <div style={{height:8}}/>
            <Inp label="− Consumer Offer" k="cof" f={f} setF={setF}/>
            <Inp label="− Special Discount" k="sdis" f={f} setF={setF}/>
            <Inp label="− Corporate Scheme" k="corp" f={f} setF={setF}/>
            <Tot label="Deal Price E = C − D" val={c.E} col="#a78bfa"/>
            <div style={{height:8}}/>
            <Inp label="− Booking Amount" k="bk" f={f} setF={setF}/>
            <div style={{display:"flex",alignItems:"center",padding:"4px 0",borderBottom:"1px solid #131820",gap:8}}><span style={{fontSize:12,color:"#64748b",flex:1}}>Booking Date</span><input type="date" value={f.bkDate||""} onChange={e=>setF(p=>({...p,bkDate:e.target.value}))} style={{width:130,background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 8px",fontSize:11,color:"#1e293b"}}/></div>
            <Inp label="− Exchange Value" k="exv" f={f} setF={setF}/>
            <Tot label="Net G = E − F" val={c.G} col="#60a5fa"/>
            {isFin&&<Inp label="− Loan / Disbursal" k="loan" f={f} setF={setF}/>}
            <div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",borderBottom:"1px solid #131820",fontSize:11}}><span style={{color:"#64748b"}}>Balance from Customer (I)</span><span style={{color:"#1e293b",fontWeight:700}}>{fc(c.I)}</span></div>
            <div style={{borderTop:"1px solid #6b8fb5",paddingTop:8,marginTop:4}}>
              <div style={{fontSize:10,color:"#f97316",fontWeight:700,marginBottom:6}}>PAYMENT RECEIVED (J)</div>
              {(f.payments||[]).map((p,i)=>(
                <div key={i} style={{display:"flex",gap:6,marginBottom:6,alignItems:"center"}}>
                  <select value={p.mode} onChange={e=>{const px=[...f.payments];px[i]={...px[i],mode:e.target.value};setF(q=>({...q,payments:px}));}} style={{background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 6px",fontSize:11,color:"#1e293b",width:90,flexShrink:0}}>
                    {["Cash","Cheque","UPI","RTGS","Finance"].map(m=><option key={m}>{m}</option>)}
                  </select>
                  <input type="number" value={p.amt||""} placeholder="Amount" onChange={e=>{const px=[...f.payments];px[i]={...px[i],amt:e.target.value};setF(q=>({...q,payments:px}));}} style={{flex:1,background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 8px",fontSize:11,color:"#1e293b"}}/>
                  <input value={p.ref||""} placeholder="Cheque No / Ref" onChange={e=>{const px=[...f.payments];px[i]={...px[i],ref:e.target.value};setF(q=>({...q,payments:px}));}} style={{flex:1,background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 8px",fontSize:11,color:"#1e293b"}}/>
                  {(f.payments||[]).length>1&&<button onClick={()=>setF(q=>({...q,payments:q.payments.filter((_,j)=>j!==i)}))} style={{background:"rgba(239,68,68,0.15)",border:"none",borderRadius:6,padding:"4px 8px",color:"#ef4444",cursor:"pointer",fontSize:13,flexShrink:0}}>✕</button>}
                </div>
              ))}
              <button onClick={()=>setF(q=>({...q,payments:[...(q.payments||[]),{mode:"Cash",amt:"",ref:""}]}))} style={{width:"100%",background:"rgba(96,165,250,0.07)",border:"1px dashed rgba(96,165,250,0.3)",borderRadius:8,padding:"6px",color:"#60a5fa",fontSize:11,cursor:"pointer",marginBottom:4}}>+ Add Payment Entry</button>
              <div style={{display:"flex",justifyContent:"space-between",padding:"5px 0",fontSize:12,fontWeight:700}}><span style={{color:"#64748b"}}>Total Received (J)</span><span style={{color:"#1e293b"}}>{fc(c.paid)}</span></div>
            </div>
            <div style={{background:c.K===0?"rgba(34,197,94,0.12)":"rgba(239,68,68,0.12)",border:"1px solid "+(c.K===0?"#22c55e":"#ef4444"),borderRadius:10,padding:"11px 12px",marginTop:6}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>Difference K = I − J</span><span style={{fontWeight:900,fontSize:20,color:c.K===0?"#22c55e":"#ef4444"}}>{fc(c.K)}</span></div>
              {c.K===0&&<div style={{fontSize:11,color:"#22c55e",marginTop:2}}>✓ Fully settled</div>}
            </div>
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#34d399",letterSpacing:0.8,marginBottom:6}}>DELIVERY CHECKLIST</div>
          <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12,display:"flex",flexDirection:"column",gap:8}}>
            {[["pdi","PDI (Pre-Delivery Inspection) done"],["helmet","Helmet given to customer"],["docs","All documents handed over"],["service","Service schedule explained (1st free @ 45 days)"]].map(([k,l])=>(
              <label key={k} style={{display:"flex",alignItems:"center",gap:9,fontSize:12,color:chk[k]?"#22c55e":"#8892a4",cursor:"pointer"}}>
                <input type="checkbox" checked={!!chk[k]} onChange={e=>setChk(p=>({...p,[k]:e.target.checked}))}/>{l}
              </label>
            ))}
          </div>
        </div>

        <div style={{marginBottom:14}}>
          <div style={{fontSize:10,fontWeight:700,color:"#ef4444",letterSpacing:0.8,marginBottom:6}}>⚠️ MANDATORY VERIFICATION (all required to submit)</div>
          <div style={{background:"#ffffff",border:"1px solid rgba(239,68,68,0.25)",borderRadius:12,padding:12,display:"flex",flexDirection:"column",gap:8}}>
            {VER_ALL.map(([k,l])=>(
              <label key={k} style={{display:"flex",alignItems:"center",gap:9,fontSize:12,color:ver[k]?"#22c55e":"#e2e6f0",cursor:"pointer"}}>
                <input type="checkbox" checked={!!ver[k]} onChange={e=>setVer(p=>({...p,[k]:e.target.checked}))}/>{ver[k]?"✅":"☐"} {l}
              </label>
            ))}
            <div style={{fontSize:11,color:"#94a3b8",borderTop:"1px solid #6b8fb5",paddingTop:8}}>{VER_ALL.filter(([k])=>ver[k]).length}/{VER_ALL.length} verified · Documents (insurance, invoice, receipt) can be uploaded after billing — not required here</div>
          </div>
        </div>

        <button onClick={saveDraft} style={{...btn("linear-gradient(135deg,#64748b,#475569)"),width:"100%",padding:13,fontSize:14,borderRadius:12,marginBottom:10}}>💾 Save Draft (continue later)</button>
        <div style={{background:role==="salesman"?"rgba(249,115,22,0.08)":"rgba(139,92,246,0.08)",border:"1px solid "+(role==="salesman"?"rgba(249,115,22,0.3)":"rgba(139,92,246,0.3)"),borderRadius:10,padding:"10px 12px",marginBottom:14,fontSize:12,color:role==="salesman"?"#f97316":"#a78bfa",fontWeight:600}}>{role==="salesman"?"⏳ Sent to Manager for approval":"✅ You are "+role+" — auto-approved"}</div>
        <button onClick={submit} style={{...btn("linear-gradient(135deg,#059669,#10b981)"),width:"100%",padding:17,fontSize:16,borderRadius:16,boxShadow:"0 6px 24px rgba(16,185,129,0.35)"}}>{busy?"Saving…":"✅ Confirm Billing & Send MR on WhatsApp"}</button>
      </div>
    </div>
  );
}

function BillingPayBox({K,custId,onAddPayment}){
  const [mode,setMode]=useState("Cash");
  const [amt,setAmt]=useState("");
  const [ref,setRef]=useState("");
  function save(){
    const a=Number(amt);
    if(!a||a<=0){alert("Enter a valid amount");return;}
    if(a>K){alert("Amount ₹"+a+" exceeds balance due ₹"+K);return;}
    onAddPayment(custId,{mode,amt:a,ref,date:td()});
    setAmt("");setRef("");
  }
  return(
    <div style={{background:"rgba(34,197,94,0.07)",border:"2px solid rgba(34,197,94,0.4)",borderRadius:14,padding:"14px 16px",marginBottom:14}}>
      <div style={{fontSize:12,fontWeight:800,color:"#22c55e",marginBottom:10}}>💳 Receive Payment — Balance: {fc(K)}</div>
      <div style={{display:"flex",gap:8,marginBottom:8}}>
        <select value={mode} onChange={e=>setMode(e.target.value)} style={{...inp,flex:"0 0 110px",padding:"9px 10px",fontSize:13,borderRadius:10}}>
          {["Cash","Cheque","UPI","RTGS","Finance"].map(m=><option key={m}>{m}</option>)}
        </select>
        <input type="number" value={amt} onChange={e=>setAmt(e.target.value)} placeholder={"Amount (max "+fc(K)+")"} style={{...inp,flex:1,padding:"9px 10px",fontSize:13,borderRadius:10}}/>
      </div>
      <input value={ref} onChange={e=>setRef(e.target.value)} placeholder="Cheque no / UPI ref (optional)" style={{...inp,width:"100%",marginBottom:10,padding:"9px 10px",fontSize:13,borderRadius:10,boxSizing:"border-box"}}/>
      <button onClick={save} style={{...btn("linear-gradient(135deg,#22c55e,#16a34a)"),width:"100%",padding:"11px",borderRadius:10,fontSize:14}}>✅ Save Receipt</button>
    </div>
  );
}
function BillingView({billing:b,cust,onAddPayment}){
  const r=RC[cust.modelCode]||{};
  const c=b.calc||calcB(b,r);
  const [showR,setShowR]=useState(null);
  const ifr=useRef(null);
  const calcText=["NKD BAJAJ - CALCULATION SHEET","Date: "+td(),"Customer: "+cust.name,"Phone: "+cust.phone,"Model: "+(cust.model||"")+" ("+(cust.modelCode||"")+")","Chassis: "+(b.chassis||"-"),"","(A) Ex-Showroom: "+fc(c.ex),"+ Comp.Acc: "+fc(c.ca),"+ Handling: "+fc(c.hdl),"+ Insurance: "+fc(c.ins),"+ Registration: "+fc(c.reg),"+ Accessories: "+fc(c.acc),"+ Teflon: "+fc(c.tef),"+ Hypo: "+fc(c.hyp),"+ AMC: "+fc(c.amcV),"TOTAL C: "+fc(c.C),"- Discounts: "+fc(c.D),"DEAL PRICE E: "+fc(c.E),"- Booking: "+fc(c.bk),"- Exchange: "+fc(c.exv)+(cust.exchangeRegNo?" ("+cust.exchangeRegNo+")":""),"- Loan: "+fc(c.loan),"Balance I: "+fc(c.I),"PAID J: "+fc(c.paid),"DIFFERENCE K: "+fc(c.K)].join("\n");
  const custText=encodeURIComponent(["NKD BAJAJ - MONEY RECEIPT","Date: "+td(),"Customer: "+cust.name,"Model: "+(cust.model||"")+" ("+(cust.modelCode||"")+")","Amount Received: "+fc(c.paid),"Balance: "+fc(Math.max(c.K,0)),"","Thank you for choosing NKD Bajaj, Dhanbad!"].join("\n"));
  const officeText=encodeURIComponent(calcText);
  function sendOffice(){
    var num=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;
    window.open("https://wa.me/91"+num+"?text="+officeText,"_blank");
  }
  return(
    <div>
      {/* ── BALANCE BANNER ── */}
      <div style={{background:c.K===0?"rgba(34,197,94,0.12)":"rgba(239,68,68,0.12)",border:"2px solid "+(c.K===0?"#22c55e":"#ef4444"),borderRadius:14,padding:"14px 16px",marginBottom:14,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <div>
          <div style={{fontSize:11,fontWeight:700,color:c.K===0?"#22c55e":"#ef4444"}}>{c.K===0?"✅ FULLY PAID":"⚠️ BALANCE DUE"}</div>
          <div style={{fontSize:11,color:"#64748b",marginTop:2}}>Paid: {fc(c.paid||0)} · Total: {fc(c.I)}</div>
        </div>
        <div style={{fontSize:26,fontWeight:900,color:c.K===0?"#22c55e":"#ef4444"}}>{fc(Math.max(c.K,0))}</div>
      </div>
      {c.K>0&&onAddPayment&&<BillingPayBox K={c.K} custId={cust.id} onAddPayment={onAddPayment}/>}
      {<div style={{marginBottom:12}}>
        <button onClick={()=>{const doc=makeMRDoc(cust,b,c);sharePdf(doc,"MR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",cust.phone,"Please find your Money Receipt from NKD Bajaj, Dhanbad.");}} style={{width:"100%",background:"rgba(37,211,102,0.1)",border:"1px solid rgba(37,211,102,0.35)",borderRadius:12,padding:13,color:"#22c55e",fontWeight:700,fontSize:13,cursor:"pointer",marginBottom:8}}>📲 Send Money Receipt PDF → Customer (WhatsApp)</button>
        <button onClick={()=>{const doc=makeCombinedDoc(cust,b,c);const num=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;sharePdf(doc,"CalcSheet_MR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",num,"Calculation Sheet + Money Receipt for "+cust.name+" ("+cust.model+")");}} style={{width:"100%",background:"rgba(249,115,22,0.1)",border:"1px solid rgba(249,115,22,0.35)",borderRadius:12,padding:13,color:"#f97316",fontWeight:700,fontSize:13,cursor:"pointer",marginBottom:8}}>🏢 Send Calc Sheet + MR (2 pages) PDF → Office</button>
        <div style={{display:"flex",gap:8}}>
          <button onClick={()=>setShowR(b.receiptHtml)} style={{flex:1,background:"rgba(96,165,250,0.08)",border:"1px solid rgba(96,165,250,0.25)",borderRadius:10,padding:10,color:"#60a5fa",fontWeight:700,fontSize:11,cursor:"pointer"}}>🧾 Preview MR</button>
          <button onClick={()=>b.calcHtml&&setShowR(b.calcHtml)} style={{flex:1,background:"rgba(245,158,11,0.08)",border:"1px solid rgba(245,158,11,0.25)",borderRadius:10,padding:10,color:"#f59e0b",fontWeight:700,fontSize:11,cursor:"pointer"}}>📊 Preview Calc</button>
        </div>
        <div style={{fontSize:10,color:"#94a3b8",marginTop:6}}>On mobile — tapping Send opens WhatsApp share sheet directly. On desktop — PDF downloads then WhatsApp opens.</div>
      </div>}
      {showR&&(
        <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.92)",zIndex:200,display:"flex",flexDirection:"column",padding:12}}>
          <div style={{display:"flex",gap:8,marginBottom:8}}>
            <button onClick={()=>dlFile(showR,showR===b.calcHtml?"CalcSheet_":"MoneyReceipt_"+cust.name.replace(/ /g,"_")+".html")} style={{flex:1,background:"#dbeafe",border:"none",borderRadius:10,padding:11,color:"#60a5fa",fontWeight:700,fontSize:12,cursor:"pointer"}}>⬇️ Download</button>
            <button onClick={()=>{try{ifr.current.contentWindow.print();}catch(e){alert("Print blocked here — use Download, open the file, then print/save as PDF");}}} style={{flex:1,background:"#c2d6ec",border:"none",borderRadius:10,padding:11,color:"#64748b",fontWeight:700,fontSize:12,cursor:"pointer"}}>🖨️ Print</button>
            <button onClick={()=>setShowR(false)} style={{flex:1,background:"#c2d6ec",border:"none",borderRadius:10,padding:11,color:"#64748b",fontWeight:700,fontSize:13,cursor:"pointer"}}>✕ Close</button>
          </div>
          <iframe ref={ifr} srcDoc={showR} title="receipt" style={{flex:1,background:"#fff",borderRadius:12,border:"none",width:"100%"}}/>
        </div>
      )}
      <div style={{background:cust.managerApproval==="approved"?"rgba(34,197,94,0.1)":"rgba(249,115,22,0.08)",border:"1px solid "+(cust.managerApproval==="approved"?"#22c55e":"rgba(249,115,22,0.4)"),borderRadius:11,padding:"11px 13px",marginBottom:12,fontSize:12,color:cust.managerApproval==="approved"?"#22c55e":"#f97316",fontWeight:600}}>{cust.managerApproval==="approved"?"✅ Billing Approved by Manager — record locked":cust.managerApproval==="rejected"?"❌ Rejected by Manager — correct the sheet and bill again (documents are retained)":"⏳ Awaiting Manager Approval"}</div>
      {b.checklist&&<div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12,marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#34d399",marginBottom:8}}>DELIVERY CHECKLIST</div>
        {[["pdi","PDI done"],["helmet","Helmet given"],["docs","Documents handed"],["service","Service explained"]].map(([k,l])=><div key={k} style={{fontSize:12,padding:"3px 0",color:b.checklist[k]?"#22c55e":"#ef4444"}}>{b.checklist[k]?"✅":"❌"} {l}</div>)}
      </div>}
      <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12,marginBottom:10}}>
        <div style={{fontSize:10,fontWeight:700,color:"#475569",marginBottom:8}}>VEHICLE</div>
        {[["Chassis",b.chassis],["Engine",b.engine],["Colour",b.color],["Delivery",fd(b.deliveryDate)],["Finance Bank",b.financeBank||"—"],["Reg No",b.registrationNo||"—"]].map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",padding:"4px 0",borderBottom:"1px solid #6b8fb5",fontSize:12}}><span style={{color:"#94a3b8"}}>{k}</span><span style={{color:"#1e293b",fontWeight:600}}>{v}</span></div>)}
      </div>
      <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12}}>
        <div style={{fontSize:10,fontWeight:700,color:"#475569",marginBottom:8}}>AMOUNTS</div>
        {[["Ex-Showroom",fc(c.ex)],["Add-ons B",fc(c.B)],["Total C",fc(c.C)],["Less D","−"+fc(c.D)],["Deal Price E",fc(c.E)],["Less F","−"+fc(c.F)],["Net G",fc(c.G)],["Loan H","−"+fc(c.loan||0)],["Balance I",fc(c.I)],["Received J",fc(c.paid||0)]].map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",padding:"3px 0",fontSize:12,borderBottom:"1px solid #131820"}}><span style={{color:"#94a3b8"}}>{k}</span><span style={{color:"#1e293b",fontWeight:600}}>{v}</span></div>)}
        <div style={{display:"flex",justifyContent:"space-between",padding:"10px 0 4px",fontWeight:800}}><span style={{color:"#1e293b"}}>Difference K</span><span style={{color:c.K===0?"#22c55e":"#ef4444",fontSize:18}}>{fc(c.K)}</span></div>
      </div>
    </div>
  );
}

function Approvals({custs,onApprove,onOpen,onEditCalc,allC,canApprove}){
  const [rem,setRem]=useState({});
  const rejected=(allC||[]).filter(c=>c.managerApproval==="rejected"&&!c.billed);
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:4}}>Approvals &amp; Actions</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:14}}>Review calculation before approving</div>
      {custs.length===0&&<div className="fu glass" style={{textAlign:"center",padding:"28px 20px",borderRadius:16,background:"#ffffff",marginBottom:14}}><div style={{fontSize:32,marginBottom:6}}>✅</div><div style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>No pending approvals</div></div>}
      {custs.map(c=>{
        const r=RC[c.modelCode]||{};const cl=c.billing?.calc||calcB(c.billing||{},r);
        return(
          <div key={c.id} className="fu glass" style={{background:"#ffffff",border:"1px solid rgba(139,92,246,0.35)",borderRadius:18,marginBottom:16,overflow:"hidden"}}>
            <div style={{padding:"12px 14px",cursor:"pointer"}} onClick={()=>onOpen(c)}>
              <div style={{fontWeight:700,fontSize:15,color:"#1e293b"}}>{c.name}</div>
              <div style={{fontSize:12,color:"#64748b"}}>{c.model} · {c.modelCode}</div>
              <div style={{fontSize:11,color:"#94a3b8",marginTop:2}}>{c.salesman} · {fd(c.billing?.deliveryDate)}</div>
            </div>
            <div style={{padding:"10px 14px",background:"#f1f5f9"}}>
              <div style={{fontSize:10,fontWeight:700,color:"#a78bfa",marginBottom:6}}>FULL CALCULATION SHEET</div>
              {[["(A) Ex-Showroom",fc(cl.ex)],["+ Comp.Acc",fc(cl.ca)],["+ Handling",fc(cl.hdl)],["+ Insurance",fc(cl.ins)],["+ Registration",fc(cl.reg)],["+ Accessories",fc(cl.acc)],["+ Teflon",fc(cl.tef)],["+ Hypo",fc(cl.hyp)],["+ AMC",fc(cl.amcV)],["TOTAL C",fc(cl.C)],["− Discounts (D)",fc(cl.D)],["DEAL PRICE E",fc(cl.E)],["− Booking",fc(cl.bk)],["− Exchange",fc(cl.exv)],["− Loan",fc(cl.loan)],["Balance (I)",fc(cl.I)],["Paid (J)",fc(cl.paid)]].map(([k,v])=><div key={k} style={{display:"flex",justifyContent:"space-between",fontSize:12,padding:"3px 0",borderBottom:"1px solid #131820"}}><span style={{color:k.includes("TOTAL")||k.includes("DEAL")?"#e2e6f0":"#5a6478",fontWeight:k.includes("TOTAL")||k.includes("DEAL")?700:400}}>{k}</span><span style={{color:"#1e293b",fontWeight:600}}>{v}</span></div>)}
              <div style={{display:"flex",justifyContent:"space-between",fontSize:14,padding:"8px 0 2px",fontWeight:800}}><span style={{color:"#1e293b"}}>Difference K</span><span style={{color:cl.K===0?"#22c55e":"#ef4444"}}>{fc(cl.K)}</span></div>
              {c.billing&&c.billing.verify&&<div>
                <div style={{fontSize:10,fontWeight:700,color:"#ef4444",margin:"10px 0 6px"}}>SALESMAN VERIFICATION</div>
                <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:4,marginBottom:8}}>
                  {(c.billing.verifyList||[]).map(([k,l])=>{
                    const done=c.billing.verify[k];
                    return(<div key={k} style={{fontSize:10,padding:"4px 7px",borderRadius:7,background:done?"rgba(34,197,94,0.08)":"rgba(239,68,68,0.1)",border:"1px solid "+(done?"rgba(34,197,94,0.3)":"rgba(239,68,68,0.4)"),color:done?"#22c55e":"#ef4444",fontWeight:600}}>{done?"✅":"❌"} {l}</div>);
                  })}
                </div>
              </div>}
              <div style={{fontSize:10,fontWeight:700,color:"#a78bfa",margin:"10px 0 6px"}}>DOCUMENT CHECKLIST</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5,marginBottom:8}}>
                {[["aadhar_photo","Aadhar"],["pan_photo","PAN"],["booking_proof","Booking Proof"],["invoice","Invoice"],["insurance","Insurance"],["moneyreceipt","Money Receipt"],["mr","MR"],["do_letter","DO"]].map(([k,l])=>{
                  const has=(c.photos||{})[k];
                  return(<div key={k} style={{display:"flex",alignItems:"center",gap:6,background:has?"rgba(34,197,94,0.08)":"rgba(107,114,128,0.08)",border:"1px solid "+(has?"rgba(34,197,94,0.3)":"#6b8fb5"),borderRadius:8,padding:"5px 8px"}}>
                    {has?<img src={c.photos[k]} alt={l} style={{width:26,height:26,objectFit:"cover",borderRadius:5}}/>:<span style={{fontSize:14}}>⬜</span>}
                    <span style={{fontSize:10,color:has?"#22c55e":"#5a6478",fontWeight:600}}>{l}</span>
                  </div>);
                })}
              </div>
              <button onClick={()=>onOpen(c)} style={{width:"100%",background:"#c2d6ec",border:"1px solid #6b8fb5",borderRadius:9,padding:8,color:"#64748b",fontSize:11,fontWeight:600,cursor:"pointer"}}>📂 View Documents</button>
            </div>
            {cl.K!==0&&<div style={{margin:"0 14px 10px",background:"rgba(239,68,68,0.1)",border:"1px solid rgba(239,68,68,0.35)",borderRadius:9,padding:"8px 11px",fontSize:11,color:"#ef4444",fontWeight:700}}>⛔ Difference K = {fc(cl.K)} — approval blocked until fully settled (K must be ₹0)</div>}
            <div style={{padding:"0 14px 8px"}}><input placeholder="Manager remark (optional — saved to history)" value={rem[c.id]||""} onChange={e=>setRem(p=>({...p,[c.id]:e.target.value}))} style={{background:"#f8fafc",border:"1px solid #6b8fb5",borderRadius:10,padding:"9px 12px",fontSize:12,color:"#1e293b",width:"100%",boxSizing:"border-box",outline:"none"}}/></div>
            {cl.K!==0&&<div style={{margin:"0 14px 8px",fontSize:11,color:"#f59e0b",fontWeight:700}}>👉 Tap EDIT to correct the sheet yourself, then it auto-approves</div>}
            {canApprove&&(<div style={{display:"flex",gap:6,padding:"10px 14px",borderTop:"1px solid #6b8fb5"}}>
              <button onClick={()=>cl.K===0&&onApprove(c.id,true,rem[c.id]||"")} disabled={cl.K!==0} style={{...btn(cl.K===0?"rgba(34,197,94,0.12)":"rgba(107,114,128,0.1)",cl.K===0?"#22c55e":"#374151"),flex:1,border:"1px solid "+(cl.K===0?"rgba(34,197,94,0.4)":"#6b8fb5"),cursor:cl.K===0?"pointer":"not-allowed",fontSize:12,padding:"11px 4px"}}>✅ Approve</button>
              <button onClick={()=>onEditCalc(c)} style={{...btn("rgba(245,158,11,0.15)","#f59e0b"),flex:1,border:"1px solid rgba(245,158,11,0.5)",fontSize:12,padding:"11px 4px"}}>✏️ EDIT</button>
              <button onClick={()=>onApprove(c.id,false,rem[c.id]||"")} style={{...btn("rgba(239,68,68,0.1)","#ef4444"),flex:1,border:"1px solid rgba(239,68,68,0.3)",fontSize:12,padding:"11px 4px"}}>❌ Reject</button>
            </div>)}
          </div>
        );
      })}
      {rejected.length>0&&<div style={{marginTop:14}}>
        <div style={{fontSize:12,fontWeight:800,color:"#ef4444",marginBottom:8}}>❌ REJECTED — AWAITING CORRECTION ({rejected.length})</div>
        {rejected.map(c=><div key={c.id} onClick={()=>onOpen(c)} className="glass" style={{background:"rgba(239,68,68,0.06)",borderRadius:12,padding:"11px 13px",marginBottom:7,cursor:"pointer"}}>
          <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{c.name}</div>
          <div style={{fontSize:11,color:"#64748b"}}>{c.model} · {c.salesman} · waiting for re-bill</div>
        </div>)}
      </div>}
      <div style={{marginTop:16}}>
        <div style={{fontSize:12,fontWeight:800,color:"#f59e0b",marginBottom:8}}>🔑 EXECUTIVE PASSWORD RESETS</div>
        <div className="glass" style={{background:"#ffffff",borderRadius:14,padding:12}}>
          {SM.map(s=>{const pws=ld("nkd_pw",{});const rec=pws[s]||{};
            return(<div key={s} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:"1px solid #1a1f2e"}}>
              <span style={{fontSize:12,color:"#1e293b"}}>{s} {rec.locked&&<span style={{color:"#ef4444",fontWeight:700,fontSize:10}}>🔒 LOCKED</span>}{rec.must&&!rec.locked&&<span style={{color:"#94a3b8",fontSize:10}}> (not set yet)</span>}</span>
              <button onClick={()=>{const p2=ld("nkd_pw",{});p2[s]={pw:"1111",must:true,fails:0,locked:false};sv("nkd_pw",p2);_dbSet("passwords",p2);alert(s+" reset to 1111 — they will set a new password on next login");}} style={{background:"#c2d6ec",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 10px",fontSize:10,color:"#f59e0b",fontWeight:700,cursor:"pointer"}}>Reset → 1111</button>
            </div>);})}
        </div>
      </div>
    </div>
  );
}

function Revival({items,onRevive}){
  const [sel,setSel]=useState([]);
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:4}}>Cold Pool Revival</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:12}}>{items.length} leads dormant 6+ months</div>
      <div style={{display:"flex",gap:7,marginBottom:12}}>
        <button onClick={()=>setSel(items.map(c=>c.id))} style={{...btn("#6b8fb5","#8892a4"),flex:1,border:"1px solid #6b8fb5"}}>All</button>
        <button onClick={()=>setSel([])} style={{...btn("#6b8fb5","#8892a4"),flex:1,border:"1px solid #6b8fb5"}}>Clear</button>
        <button onClick={()=>{onRevive(sel);setSel([]);}} disabled={sel.length===0} style={{...btn(sel.length>0?"#dbeafe":"#f4f7fb",sel.length>0?"#60a5fa":"#374151"),flex:2,border:"1px solid "+(sel.length>0?"#3b82f6":"#6b8fb5")}}>Revive {sel.length>0?"("+sel.length+")":""}</button>
      </div>
      {items.map(c=>(
        <div key={c.id} onClick={()=>setSel(p=>p.includes(c.id)?p.filter(x=>x!==c.id):[...p,c.id])} style={{background:sel.includes(c.id)?"rgba(96,165,250,0.07)":"#ffffff",border:"1px solid "+(sel.includes(c.id)?"rgba(96,165,250,0.4)":"#6b8fb5"),borderRadius:12,padding:"12px 14px",marginBottom:7,cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <div><div style={{fontWeight:600,fontSize:13,color:"#1e293b"}}>{c.name}</div><div style={{fontSize:11,color:"#94a3b8"}}>{c.phone} · {c.model}</div><div style={{fontSize:10,color:"#64748b"}}>Enquired: {fd(c.enquiryDate)}</div></div>
          <div style={{width:22,height:22,borderRadius:6,border:"2px solid "+(sel.includes(c.id)?"#60a5fa":"#6b8fb5"),background:sel.includes(c.id)?"#60a5fa":"transparent",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}>{sel.includes(c.id)&&<span style={{color:"#1e293b",fontSize:14,fontWeight:700}}>✓</span>}</div>
        </div>
      ))}
    </div>
  );
}

function Reports({custs,onImportCust}){
  const allC=custs;
  const [brF,setBrF]=useState("All");
  const [repMonth,setRepMonth]=useState(new Date().toISOString().slice(0,7));
  custs=brF==="All"?custs:custs.filter(c=>(c.branch||SM_BRANCH[c.salesman])===brF);
  const billed=custs.filter(c=>c.billed);
  const billedForMonth=allC.filter(c=>c.billed&&(c.billedDate||"").startsWith(repMonth));
  const conv=custs.length>0?((billed.length/custs.length)*100).toFixed(1):0;
  const cats={};custs.forEach(c=>{if(c.cat)cats[c.cat]=(cats[c.cat]||0)+1;});
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:12}}>Owner Reports</div>
      <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:13,padding:"12px 14px",marginBottom:12}}>
        <div style={{fontSize:11,color:"#f97316",fontWeight:700,marginBottom:6}}>📅 SELECT MONTH FOR REPORTS</div>
        <input type="month" value={repMonth} onChange={e=>setRepMonth(e.target.value)} style={{...inp,fontSize:14,padding:"10px 12px"}}/>
        <div style={{fontSize:11,color:"#94a3b8",marginTop:6}}>{billedForMonth.length} billed customers in {repMonth}</div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:8,marginBottom:14}}>
        <button onClick={()=>{
          const H=["Bill Date","Customer","Phone","Address","Father Name","Aadhar","PAN","Model","Code","Chassis","Engine","Colour","Delivery Date","MR No","Pay Mode","Financed By","Reg No","Salesman","Branch","Last Modified"];
          const rows=billedForMonth.filter(c=>c.billing).map(c=>{const b=c.billing;
            return[c.billedDate||"",c.name||"",c.phone||"",c.address||"",c.fatherName||"",c.aadhar||"",c.pan||"",c.model||"",c.modelCode||"",b.chassis||"",b.engine||"",b.color||"",b.deliveryDate||"",b.mrNo||"",b.payMode||"",b.financeBank||"Cash",b.registrationNo||"",c.salesman||"",c.branch||"",c.updatedAt||c.billedDate||""];
          });
          if(rows.length===0){alert("No billed customers for "+repMonth);return;}
          const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);
          ws["!cols"]=H.map(()=>({wch:16}));XLSX.utils.book_append_sheet(wb,ws,"Billing Team");
          XLSX.writeFile(wb,"NKD_BillingTeam_"+repMonth+".xlsx");
        }} style={{width:"100%",background:"rgba(52,211,153,0.1)",border:"1px solid rgba(52,211,153,0.35)",borderRadius:11,padding:"13px",color:"#34d399",fontWeight:700,fontSize:13,cursor:"pointer"}}>🧾 Export Billing Team Report — {repMonth}</button>
        <button onClick={()=>{
          const H=["Bill Date","Customer","Phone","Address","Father Name","Model","Code","Chassis","Engine","Delivery Date","MR No","Pay Mode","Financed By","Reg No","Ex-Showroom","Comp Acc","Handling","Insurance","Registration","Accessories","Teflon","Hypo","AMC","TOTAL ON-ROAD","Consumer Offer","Special Disc","Corporate","DEAL PRICE","Booking Amt","Exchange Vehicle","Exchange Value","NET AMT","Loan","BALANCE","PAID","DIFF","Salesman","Branch","Approved By","Enquiry Date","Last Modified"];
          const rows=billedForMonth.filter(c=>c.billing&&c.billing.calc).map(c=>{const b=c.billing,k=b.calc;
            return[c.billedDate||"",c.name||"",c.phone||"",c.address||"",c.fatherName||"",c.model||"",c.modelCode||"",b.chassis||"",b.engine||"",b.deliveryDate||"",b.mrNo||"",b.payMode||"",b.financeBank||"Cash",b.registrationNo||"",k.ex,k.ca,k.hdl,k.ins,k.reg,k.acc,k.tef,k.hyp,k.amcV,k.C,k.cof,k.sdis,k.corp,k.E,k.bk,c.exchangeAsked||"",k.exv,k.G,k.loan,k.I,k.paid,k.K,c.salesman||"",c.branch||"",c.approvedBy||"",c.enquiryDate||"",c.updatedAt||c.billedDate||""];
          });
          if(rows.length===0){alert("No billed customers for "+repMonth);return;}
          const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);
          ws["!cols"]=H.map(()=>({wch:16}));XLSX.utils.book_append_sheet(wb,ws,"Accounts Team");
          XLSX.writeFile(wb,"NKD_AccountsTeam_"+repMonth+".xlsx");
        }} style={{width:"100%",background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.35)",borderRadius:11,padding:"13px",color:"#60a5fa",fontWeight:700,fontSize:13,cursor:"pointer"}}>💰 Export Accounts Team Report — {repMonth}</button>
      </div>
      <div style={{display:"flex",gap:5,marginBottom:10,overflowX:"auto"}}>
        {["All",...BRANCHES].map(b=><button key={b} onClick={()=>setBrF(b)} style={{padding:"6px 12px",borderRadius:10,fontSize:11,fontWeight:700,cursor:"pointer",flexShrink:0,background:brF===b?"#dbeafe":"#6b8fb5",color:brF===b?"#60a5fa":"#8892a4",border:"none"}}>{b}</button>)}
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:14}}>
        {BRANCHES.map(b=>{const bc=allC.filter(c=>(c.branch||SM_BRANCH[c.salesman])===b);const bb=bc.filter(c=>c.billed);
          const rev=bb.reduce((s,c)=>s+((c.billing&&c.billing.calc&&c.billing.calc.E)||0),0);
          return(<div key={b} className="glass" style={{background:"#ffffff",borderRadius:14,padding:"11px 8px",textAlign:"center"}}>
            <div style={{fontSize:9,color:"#94a3b8",fontWeight:700}}>{b.toUpperCase()}</div>
            <div style={{fontSize:17,fontWeight:900,color:"#34d399",marginTop:3}}>{bb.length}</div>
            <div style={{fontSize:9,color:"#94a3b8"}}>sold · ₹{(rev/100000).toFixed(1)}L</div>
          </div>);})}
      </div>
      <button onClick={()=>{
        const rows=[["Salesman","Branch","Total Leads","Bookings","Test Rides","Docs Collected","Sold","Revenue"]];
        SM.forEach(s=>{const m=allC.filter(c=>c.salesman===s);
          rows.push([s,SM_BRANCH[s],m.length,m.filter(c=>c.booking).length,m.filter(c=>c.testRide).length,m.filter(c=>c.photos&&Object.keys(c.photos).some(k=>c.photos[k])).length,m.filter(c=>c.billed).length,m.filter(c=>c.billed).reduce((x,c)=>x+((c.billing&&c.billing.calc&&c.billing.calc.E)||0),0)]);});
        dlFile(rows.map(r=>r.join(",")).join("\n"),"NKD_Branch_Report_"+td()+".csv","text/csv");
      }} style={{width:"100%",background:"rgba(52,211,153,0.1)",border:"1px solid rgba(52,211,153,0.35)",borderRadius:12,padding:11,color:"#34d399",fontWeight:700,fontSize:12,cursor:"pointer",marginBottom:10}}>📊 Export Branch Report (Excel/CSV)</button>
      <div className="glass" style={{background:"#ffffff",borderRadius:14,padding:12,marginBottom:14}}>
        <div style={{fontSize:11,fontWeight:800,color:"#60a5fa",marginBottom:8}}>⬆️ BULK UPLOADS (save Excel as .CSV first: File → Save As → CSV)</div>
        <div style={{fontSize:10,color:"#94a3b8",marginBottom:4}}>1. New Rate Chart — columns: code,name,cat,ex,cAcc,hdl,ins,reg,onRoad,amc</div>
        <input type="file" accept=".csv" onChange={e=>{const fl=e.target.files[0];if(!fl)return;const rd=new FileReader();rd.onload=ev=>{
          try{const lines=ev.target.result.split(/\r?\n/).filter(x=>x.trim());const hd=lines[0].toLowerCase().split(",").map(x=>x.trim());
          const gi=n=>hd.indexOf(n);const obj={};
          for(let i=1;i<lines.length;i++){const c2=lines[i].split(",");const cd=(c2[gi("code")]||"").trim().toUpperCase();if(!cd)continue;
            obj[cd]={n:c2[gi("name")]||cd,cat:c2[gi("cat")]||"OTHER",ex:+c2[gi("ex")]||0,cAcc:+c2[gi("cacc")]||0,hdl:+c2[gi("hdl")]||600,ins:+c2[gi("ins")]||0,reg:+c2[gi("reg")]||0,onRoad:+c2[gi("onroad")]||0,amc:+c2[gi("amc")]||0};}
          sv("nkd_rc",obj);_dbSet("rate_chart",obj);Object.assign(RC,obj);alert("✅ Rate chart updated: "+Object.keys(obj).length+" models. New rates apply to all new billings.");}catch(err){alert("Could not read file — check column names");}
        };rd.readAsText(fl);e.target.value="";}} style={{width:"100%",background:"#c2d6ec",borderRadius:9,padding:8,fontSize:11,color:"#64748b",border:"1px dashed #2a3040",marginBottom:10}}/>
        <div style={{fontSize:10,color:"#94a3b8",marginBottom:4}}>2. Old Customers (10,000+) — columns: name,phone,model,salesman,enquirydate,status,remarks</div>
        <input type="file" accept=".csv" onChange={e=>{const fl=e.target.files[0];if(!fl)return;const rd=new FileReader();rd.onload=ev=>{
          try{const lines=ev.target.result.split(/\r?\n/).filter(x=>x.trim());const hd=lines[0].toLowerCase().split(",").map(x=>x.trim());
          const gi=n=>hd.indexOf(n);const rows=[];const perDay={};
          for(let i=1;i<lines.length;i++){const c2=lines[i].split(",");const nm=(c2[gi("name")]||"").trim();const ph=(c2[gi("phone")]||"").replace(/\D/g,"").slice(-10);if(!nm||ph.length!==10)continue;
            const sm2=(c2[gi("salesman")]||"").trim()||SM[i%SM.length];
            perDay[sm2]=(perDay[sm2]||0)+1;const off=Math.floor((perDay[sm2]-1)/80);
            rows.push({id:"OLD"+Date.now()+"_"+i,name:nm,phone:ph,model:(c2[gi("model")]||"").trim(),modelCode:"",cat:"",address:"",enquiryDate:(c2[gi("enquirydate")]||td()).trim()||td(),status:(c2[gi("status")]||"Cold").trim()||"Cold",salesman:sm2,branch:SM_BRANCH[sm2]||BRANCHES[0],finance:"Cash",remarks:"["+td()+"] IMPORTED: old customer data. "+((c2[gi("remarks")]||"").trim()),followupDate:aD(td(),off),attempts:0,stopped:false,billed:false,billedDate:null,photos:{},billing:null,managerApproval:null,callLog:[]});}
          onImportCust(rows);alert("✅ Imported "+rows.length+" customers — spread at max 80 calls/day per executive for followup");}catch(err){alert("Could not read file — check column names");}
        };rd.readAsText(fl);e.target.value="";}} style={{width:"100%",background:"#c2d6ec",borderRadius:9,padding:8,fontSize:11,color:"#64748b",border:"1px dashed #2a3040"}}/>
      </div>
      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:18}}>
        {[["Total",custs.length,"#60a5fa"],["Billed",billed.length,"#22c55e"],["Conv%",conv+"%","#f59e0b"],["Hot",custs.filter(c=>c.status==="Hot"&&!c.billed).length,"#ef4444"],["Lost",custs.filter(c=>c.status==="Lost").length,"#6b7280"],["Stopped",custs.filter(c=>c.stopped).length,"#f97316"]].map(([l,v,c])=>(
          <div key={l} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:"12px 8px",textAlign:"center"}}><div style={{fontSize:22,fontWeight:900,color:c,lineHeight:1}}>{v}</div><div style={{fontSize:10,color:"#94a3b8",marginTop:4}}>{l}</div></div>
        ))}
      </div>
      {(()=>{const fin=custs.filter(c=>c.finance==="Finance");const finB=fin.filter(c=>c.billed);
        const exB=billed.filter(c=>c.billing&&Number(c.billing.exv)>0);
        const avgF=billed.length?(billed.reduce((s,c)=>s+((c.callLog||[]).length),0)/billed.length).toFixed(1):0;
        return(<div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:8,marginBottom:16}}>
          {[["Finance Conv",fin.length?Math.round(finB.length/fin.length*100)+"%":"—","#f59e0b"],["Exchange %",billed.length?Math.round(exB.length/billed.length*100)+"%":"—","#a78bfa"],["Avg Followups",avgF,"#60a5fa"]].map(([l,v,c2])=>(
            <div key={l} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:"12px 8px",textAlign:"center"}}>
              <div style={{fontSize:20,fontWeight:900,color:c2,lineHeight:1}}>{v}</div>
              <div style={{fontSize:9,color:"#94a3b8",marginTop:4}}>{l}</div>
            </div>
          ))}
        </div>);})()}
      {(()=>{const src={};custs.forEach(c=>{if(c.source)src[c.source]=(src[c.source]||0)+1;});const lost={};custs.forEach(c=>{if(c.lostReason)lost[c.lostReason]=(lost[c.lostReason]||0)+1;});
        return(<div style={{marginBottom:16}}>
          {Object.keys(src).length>0&&<div><div style={{fontSize:12,fontWeight:700,color:"#64748b",marginBottom:8}}>LEAD SOURCES</div>
          {Object.entries(src).sort((a,b)=>b[1]-a[1]).map(([k,n])=><div key={k} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:10,padding:"8px 13px",marginBottom:5,display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:"#1e293b"}}>{k}</span><span style={{color:"#60a5fa",fontWeight:700}}>{n}</span></div>)}</div>}
          {Object.keys(lost).length>0&&<div style={{marginTop:12}}><div style={{fontSize:12,fontWeight:700,color:"#ef4444",marginBottom:8}}>WHY WE LOSE CUSTOMERS</div>
          {Object.entries(lost).sort((a,b)=>b[1]-a[1]).map(([k,n])=><div key={k} style={{background:"rgba(239,68,68,0.06)",border:"1px solid rgba(239,68,68,0.2)",borderRadius:10,padding:"8px 13px",marginBottom:5,display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:"#1e293b"}}>{k}</span><span style={{color:"#ef4444",fontWeight:700}}>{n}</span></div>)}</div>}
        </div>);})()}
      <div style={{fontSize:12,fontWeight:700,color:"#f59e0b",marginBottom:8}}>🔑 EXECUTIVE PASSWORDS</div>
      <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:12,marginBottom:16}}>
        {SM.map(s=>{const pws=ld("nkd_pw",{});const rec=pws[s]||{};
          return(<div key={s} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"6px 0",borderBottom:"1px solid #1a1f2e"}}>
            <span style={{fontSize:12,color:"#1e293b"}}>{s} {rec.locked&&<span style={{color:"#ef4444",fontWeight:700,fontSize:10}}>🔒 LOCKED</span>}</span>
            <button onClick={()=>{const pws2=ld("nkd_pw",{});pws2[s]={pw:"1111",must:true,fails:0,locked:false};sv("nkd_pw",pws2);_dbSet("passwords",pws2);alert(s+" reset to 1111 — they must change it on next login");}} style={{background:"#c2d6ec",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 10px",fontSize:10,color:"#f59e0b",fontWeight:700,cursor:"pointer"}}>Reset → 1111</button>
          </div>);})}
      </div>
      <div style={{fontSize:12,fontWeight:700,color:"#64748b",marginBottom:8}}>SALESMAN SCOREBOARD</div>
      {SM.map(s=>{
        const m=custs.filter(c=>c.salesman===s);const b=m.filter(c=>c.billed).length;const pct=m.length>0?Math.round(b/m.length*100):0;
        return(
          <div key={s} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:12,padding:"12px 14px",marginBottom:8}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:6}}><div><div style={{fontWeight:600,fontSize:13,color:"#1e293b"}}>{s}</div><div style={{fontSize:11,color:"#94a3b8"}}>{m.length} leads</div></div><div style={{textAlign:"right"}}><div style={{fontSize:20,fontWeight:900,color:"#34d399"}}>{b}</div><div style={{fontSize:9,color:"#94a3b8"}}>BILLED</div></div></div>
            <div style={{background:"#c2d6ec",borderRadius:5,height:5}}><div style={{background:"linear-gradient(90deg,#f97316,#34d399)",borderRadius:5,height:5,width:pct+"%",transition:"width 0.4s"}}/></div>
            <div style={{fontSize:10,color:"#94a3b8",marginTop:3}}>{pct}% conversion</div>
          </div>
        );
      })}
    </div>
  );
}

function DocVault({custs,onImport}){
  const [open,setOpen]=useState(null);
  const [flt,setFlt]=useState("approved");
  const [repMonth,setRepMonth]=useState(new Date().toISOString().slice(0,7));
  const billedForMonth=custs.filter(c=>c.billed&&(c.billedDate||"").startsWith(repMonth));
  const withDocs=custs.filter(c=>{
    const hasDocs=c.photos&&Object.keys(c.photos).filter(k=>c.photos[k]).length>0;
    if(flt==="approved")return hasDocs&&c.billed&&c.managerApproval==="approved";
    if(flt==="billed")return hasDocs&&c.billed;
    return hasDocs;
  });
  function exportDB(){
    var blob=new Blob([JSON.stringify(custs,null,1)],{type:"application/json"});
    var a=document.createElement("a");a.href=URL.createObjectURL(blob);
    a.download="NKD_Bajaj_Database_"+td()+".json";a.click();
  }
  function importDB(file){
    var rd=new FileReader();
    rd.onload=function(e){try{var d=JSON.parse(e.target.result);if(Array.isArray(d))onImport(d);else alert("Invalid database file");}catch(err){alert("Could not read file");}};
    rd.readAsText(file);
  }
  const LABELS={aadhar_photo:"Aadhar",pan_photo:"PAN",booking_proof:"Booking Proof",exchange_eval:"Exchange Evaluation",delivery:"Delivery Photo",invoice:"Invoice",insurance:"Insurance",registration:"RC",moneyreceipt:"Money Receipt",exchange:"Exchange Handover",mr:"MR",do_letter:"DO",loan_app:"Loan Application",sanction:"Sanction Letter",cust_sign:"Customer Sign"};
  return(
    <div>
      <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:4}}>📁 Document Vault</div>
      <div style={{fontSize:11,color:"#94a3b8",marginBottom:10}}>{withDocs.length} parties · tap to expand · download files</div>
      <div style={{display:"flex",gap:6,marginBottom:10}}>
        {[["approved","✅ Approved"],["billed","Billed"],["all","All"]].map(([k,l])=><button key={k} onClick={()=>setFlt(k)} style={{flex:1,padding:"7px",borderRadius:10,fontSize:11,fontWeight:700,cursor:"pointer",background:flt===k?"#dbeafe":"#6b8fb5",color:flt===k?"#60a5fa":"#8892a4",border:"1px solid "+(flt===k?"#3b82f6":"#6b8fb5")}}>{l}</button>)}
      </div>
      <div style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:13,padding:"12px 14px",marginBottom:12}}>
        <div style={{fontSize:11,color:"#f97316",fontWeight:700,marginBottom:6}}>📅 SELECT MONTH FOR REPORTS</div>
        <input type="month" value={repMonth} onChange={e=>setRepMonth(e.target.value)} style={{...inp,fontSize:14,padding:"10px 12px"}}/>
        <div style={{fontSize:11,color:"#94a3b8",marginTop:6}}>{billedForMonth.length} billed customers in {repMonth}</div>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:7,marginBottom:14}}>
        <button onClick={()=>{
          const H=["Bill Date","Customer","Phone","Address","Father Name","Aadhar","PAN","Model","Code","Chassis","Engine","Colour","Delivery Date","MR No","Pay Mode","Financed By","Reg No","Salesman","Branch","Last Modified"];
          const rows=billedForMonth.filter(c=>c.billing).map(c=>{const b=c.billing;
            return[c.billedDate||"",c.name||"",c.phone||"",c.address||"",c.fatherName||"",c.aadhar||"",c.pan||"",c.model||"",c.modelCode||"",b.chassis||"",b.engine||"",b.color||"",b.deliveryDate||"",b.mrNo||"",b.payMode||"",b.financeBank||"Cash",b.registrationNo||"",c.salesman||"",c.branch||"",c.updatedAt||c.billedDate||""];
          });
          if(rows.length===0){alert("No billed customers for "+repMonth);return;}
          const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);
          ws["!cols"]=H.map(()=>({wch:16}));XLSX.utils.book_append_sheet(wb,ws,"Billing Team");
          XLSX.writeFile(wb,"NKD_BillingTeam_"+repMonth+".xlsx");
        }} style={{width:"100%",background:"rgba(52,211,153,0.1)",border:"1px solid rgba(52,211,153,0.35)",borderRadius:11,padding:"11px",color:"#34d399",fontWeight:700,fontSize:12,cursor:"pointer"}}>🧾 Billing Team — {repMonth}</button>
        <button onClick={()=>{
          const H=["Bill Date","Customer","Phone","Address","Father Name","Model","Code","Chassis","Engine","Delivery Date","MR No","Pay Mode","Financed By","Reg No","Ex-Showroom","Comp Acc","Handling","Insurance","Registration","Accessories","Teflon","Hypo","AMC","TOTAL ON-ROAD","Consumer Offer","Special Disc","Corporate","DEAL PRICE","Booking Amt","Exchange Vehicle","Exchange Value","NET AMT","Loan","BALANCE","PAID","DIFF","Salesman","Branch","Approved By","Enquiry Date","Last Modified"];
          const rows=billedForMonth.filter(c=>c.billing&&c.billing.calc).map(c=>{const b=c.billing,k=b.calc;
            return[c.billedDate||"",c.name||"",c.phone||"",c.address||"",c.fatherName||"",c.model||"",c.modelCode||"",b.chassis||"",b.engine||"",b.deliveryDate||"",b.mrNo||"",b.payMode||"",b.financeBank||"Cash",b.registrationNo||"",k.ex,k.ca,k.hdl,k.ins,k.reg,k.acc,k.tef,k.hyp,k.amcV,k.C,k.cof,k.sdis,k.corp,k.E,k.bk,c.exchangeAsked||"",k.exv,k.G,k.loan,k.I,k.paid,k.K,c.salesman||"",c.branch||"",c.approvedBy||"",c.enquiryDate||"",c.updatedAt||c.billedDate||""];
          });
          if(rows.length===0){alert("No billed customers for "+repMonth);return;}
          const wb=XLSX.utils.book_new();const ws=XLSX.utils.aoa_to_sheet([H,...rows]);
          ws["!cols"]=H.map(()=>({wch:16}));XLSX.utils.book_append_sheet(wb,ws,"Accounts Team");
          XLSX.writeFile(wb,"NKD_AccountsTeam_"+repMonth+".xlsx");
        }} style={{width:"100%",background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.35)",borderRadius:11,padding:"11px",color:"#60a5fa",fontWeight:700,fontSize:12,cursor:"pointer"}}>💰 Accounts Team — {repMonth}</button>
        <div style={{display:"flex",gap:7}}>
          <button onClick={exportDB} style={{flex:1,background:"rgba(52,211,153,0.1)",border:"1px solid rgba(52,211,153,0.35)",borderRadius:11,padding:"11px",color:"#34d399",fontWeight:700,fontSize:12,cursor:"pointer"}}>⬇️ Export Full Database</button>
          <label style={{flex:1,background:"rgba(139,92,246,0.1)",border:"1px solid rgba(139,92,246,0.35)",borderRadius:11,padding:"11px",color:"#a78bfa",fontWeight:700,fontSize:12,cursor:"pointer",textAlign:"center"}}>⬆️ Import Database<input type="file" accept=".json" style={{display:"none"}} onChange={e=>e.target.files[0]&&importDB(e.target.files[0])}/></label>
        </div>
      </div>
      {withDocs.length===0&&<div style={{color:"#64748b",textAlign:"center",padding:32,fontSize:13}}>{flt==="approved"?"No approved parties with documents yet":"No documents uploaded yet"}</div>}
      {withDocs.map(c=>{
        const keys=Object.keys(c.photos).filter(k=>c.photos[k]);
        const isOpen=open===c.id;
        return(
          <div key={c.id} style={{background:"#ffffff",border:"1px solid #6b8fb5",borderRadius:13,marginBottom:10,overflow:"hidden"}}>
            <div onClick={()=>setOpen(isOpen?null:c.id)} style={{padding:"12px 14px",cursor:"pointer",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div><div style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>{c.name}</div><div style={{fontSize:11,color:"#94a3b8"}}>{c.phone} · {c.model} · {c.salesman}</div></div>
              <div style={{textAlign:"right"}}><span style={{fontSize:12,color:"#60a5fa",fontWeight:700}}>{keys.length} docs</span><span style={{color:"#94a3b8",marginLeft:6}}>{isOpen?"▲":"▼"}</span></div>
            </div>
            {isOpen&&(
              <div style={{borderTop:"1px solid #6b8fb5",padding:"12px 14px",display:"flex",flexDirection:"column",gap:10}}>
                {c.billing&&c.billing.receiptHtml&&(
                  <button onClick={()=>dlFile(c.billing.receiptHtml,"MoneyReceipt_"+c.name.replace(/ /g,"_")+".html")} style={{background:"rgba(96,165,250,0.1)",border:"1px solid rgba(96,165,250,0.35)",borderRadius:10,padding:"10px",color:"#60a5fa",fontWeight:700,fontSize:12,cursor:"pointer"}}>⬇️ Download Money Receipt (open → print → save PDF)</button>
                )}
                {keys.map(k=>(
                  <div key={k} style={{background:"#f1f5f9",border:"1px solid #6b8fb5",borderRadius:10,overflow:"hidden"}}>
                    <div style={{padding:"8px 12px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                      <span style={{fontSize:12,color:"#334155",fontWeight:600}}>{LABELS[k]||k}</span>
                      <a href={c.photos[k]} download={(LABELS[k]||k).replace(/ /g,"_")+"_"+c.name.replace(/ /g,"_")+".jpg"} style={{background:"#dbeafe",borderRadius:8,padding:"5px 12px",color:"#60a5fa",fontSize:11,fontWeight:700,textDecoration:"none"}}>⬇️ Download</a>
                    </div>
                    <img src={c.photos[k]} alt={k} style={{width:"100%",maxHeight:180,objectFit:"contain",background:"#000"}}/>
                  </div>
                ))}
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

function UserMgmt({nkdUsers,onSave,notify}){
  const [users,setUsers]=useState(()=>({manager:[...(nkdUsers.manager||[])],owner:[...(nkdUsers.owner||[])],admin:[...(nkdUsers.admin||[])]}));
  const [newRole,setNewRole]=useState("manager");
  const [newName,setNewName]=useState("");
  const [newPin,setNewPin]=useState("");
  const ROLE_COLOR={manager:"#3b82f6",owner:"#f97316",admin:"#8b5cf6",tech:"#0ea5e9"};
  const ROLE_LABEL={manager:"Manager",owner:"Owner",admin:"Admin",tech:"Tech"};
  function addUser(){
    if(!newName.trim()){notify("❌ Enter a name");return;}
    if(newPin.length<4){notify("❌ PIN must be at least 4 digits");return;}
    const already=(users[newRole]||[]).find(u=>u.name.toLowerCase()===newName.trim().toLowerCase());
    if(already){notify("❌ Name already exists for this role");return;}
    const updated={...users,[newRole]:[...(users[newRole]||[]),{name:newName.trim(),pin:newPin}]};
    setUsers(updated);onSave(updated);setNewName("");setNewPin("");notify("✅ "+ROLE_LABEL[newRole]+" account added");
  }
  function removeUser(role,name){
    if((users[role]||[]).length<=1){notify("❌ Must keep at least one account per role");return;}
    const updated={...users,[role]:(users[role]||[]).filter(u=>u.name!==name)};
    setUsers(updated);onSave(updated);notify("Removed "+name);
  }
  function changePin(role,name,pin){
    if(pin.length<4){notify("❌ PIN min 4 digits");return;}
    const updated={...users,[role]:(users[role]||[]).map(u=>u.name===name?{...u,pin}:u)};
    setUsers(updated);onSave(updated);notify("✅ PIN updated for "+name);
  }
  return(
    <div style={{maxWidth:700}}>
      <div style={{fontWeight:800,fontSize:20,color:"#1e293b",marginBottom:4}}>👤 User Accounts</div>
      <div style={{fontSize:12,color:"#94a3b8",marginBottom:20}}>Manage login credentials for Manager, Owner, and Admin roles.</div>
      {/* Existing users by role */}
      {["manager","owner","admin","tech"].map(r=>(
        <div key={r} style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,padding:"16px 18px",marginBottom:16}}>
          <div style={{fontWeight:800,fontSize:13,color:ROLE_COLOR[r],marginBottom:10,textTransform:"uppercase",letterSpacing:1}}>{ROLE_LABEL[r]} Accounts</div>
          {(users[r]||[]).map((u,i)=>(
            <UserPinRow key={i} u={u} color={ROLE_COLOR[r]} onRemove={()=>removeUser(r,u.name)} onPin={(p)=>changePin(r,u.name,p)}/>
          ))}
        </div>
      ))}
      {/* Add new user */}
      <div style={{background:"rgba(249,115,22,0.05)",border:"2px dashed #f97316",borderRadius:14,padding:"16px 18px"}}>
        <div style={{fontWeight:700,fontSize:13,color:"#f97316",marginBottom:12}}>➕ Add New Account</div>
        <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr auto",gap:10,alignItems:"end"}}>
          <div><div style={{fontSize:11,color:"#64748b",fontWeight:600,marginBottom:4}}>Role</div>
            <select style={{...inp,margin:0}} value={newRole} onChange={e=>setNewRole(e.target.value)}>
              <option value="manager">Manager</option><option value="owner">Owner</option><option value="admin">Admin</option><option value="tech">Tech</option>
            </select>
          </div>
          <div><div style={{fontSize:11,color:"#64748b",fontWeight:600,marginBottom:4}}>Username</div>
            <input style={{...inp,margin:0}} value={newName} onChange={e=>setNewName(e.target.value)} placeholder="Full name"/>
          </div>
          <div><div style={{fontSize:11,color:"#64748b",fontWeight:600,marginBottom:4}}>PIN (min 4 digits)</div>
            <input type="password" style={{...inp,margin:0}} value={newPin} onChange={e=>setNewPin(e.target.value)} placeholder="••••"/>
          </div>
          <button onClick={addUser} style={{...btn("linear-gradient(135deg,#f97316,#ef4444)"),padding:"10px 18px",borderRadius:10,whiteSpace:"nowrap"}}>Add</button>
        </div>
      </div>
    </div>
  );
}
function UserPinRow({u,color,onRemove,onPin}){
  const [editing,setEditing]=useState(false);
  const [p,setP]=useState("");
  return(
    <div style={{display:"flex",alignItems:"center",gap:10,padding:"8px 0",borderBottom:"1px solid #e8eef8"}}>
      <div style={{flex:1,fontWeight:600,fontSize:13,color:"#1e293b"}}>{u.name}</div>
      <div style={{fontSize:11,color:"#94a3b8",background:"#f1f5f9",borderRadius:6,padding:"2px 8px"}}>PIN: {"•".repeat(u.pin.length)}</div>
      {editing?(
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <input type="password" placeholder="New PIN" value={p} onChange={e=>setP(e.target.value)} style={{...inp,margin:0,width:90,padding:"5px 8px",fontSize:12}}/>
          <button onClick={()=>{onPin(p);setEditing(false);setP("");}} style={{...btn(color),padding:"5px 10px",borderRadius:7,fontSize:11}}>Save</button>
          <button onClick={()=>{setEditing(false);setP("");}} style={{padding:"5px 8px",borderRadius:7,border:"1px solid #6b8fb5",background:"transparent",cursor:"pointer",fontSize:11,color:"#475569"}}>✕</button>
        </div>
      ):(
        <button onClick={()=>setEditing(true)} style={{padding:"5px 10px",borderRadius:7,border:"1px solid #6b8fb5",background:"transparent",cursor:"pointer",fontSize:11,color:"#475569",fontWeight:600}}>Change PIN</button>
      )}
      <button onClick={onRemove} style={{padding:"5px 8px",borderRadius:7,border:"1px solid #ef444455",background:"rgba(239,68,68,0.07)",cursor:"pointer",fontSize:11,color:"#ef4444",fontWeight:700}}>Remove</button>
    </div>
  );
}

function PaymentNotifPopup({notifs,onDismiss}){
  const unread=(notifs||[]).filter(n=>!n.read);
  if(!unread.length)return null;
  return(
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.55)",zIndex:2000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:"#fff",borderRadius:20,padding:24,maxWidth:420,width:"100%",boxShadow:"0 24px 80px rgba(0,0,0,0.35)",fontFamily:"'Sora','Inter',sans-serif"}}>
        <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:4}}>
          <div style={{fontSize:28}}>💳</div>
          <div>
            <div style={{fontSize:17,fontWeight:800,color:"#1e293b"}}>Payment Received</div>
            <div style={{fontSize:11,color:"#64748b"}}>{unread.length} new payment{unread.length>1?"s":""} recorded</div>
          </div>
        </div>
        <div style={{maxHeight:320,overflowY:"auto",margin:"14px 0"}}>
          {unread.map(n=>(
            <div key={n.id} style={{background:"rgba(34,197,94,0.07)",border:"1px solid rgba(34,197,94,0.35)",borderRadius:12,padding:"12px 14px",marginBottom:8}}>
              <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{n.custName} <span style={{color:"#64748b",fontWeight:400}}>· {n.model}</span></div>
              <div style={{fontSize:13,color:"#22c55e",fontWeight:800,margin:"3px 0"}}>₹{Number(n.amt).toLocaleString("en-IN")} <span style={{fontWeight:500,color:"#475569"}}>via {n.mode}</span></div>
              {n.balance>0&&<div style={{fontSize:11,color:"#ef4444",fontWeight:600}}>⚠️ Balance still due: ₹{Number(n.balance).toLocaleString("en-IN")}</div>}
              {n.balance===0&&<div style={{fontSize:11,color:"#22c55e",fontWeight:600}}>✅ Fully settled</div>}
              <div style={{fontSize:10,color:"#94a3b8",marginTop:4}}>By {n.salesman} · {n.date} {n.time}</div>
            </div>
          ))}
        </div>
        <button onClick={onDismiss} style={{width:"100%",padding:"12px",background:"linear-gradient(135deg,#1d4ed8,#3b82f6)",border:"none",borderRadius:11,color:"#fff",fontWeight:700,fontSize:14,cursor:"pointer"}}>✓ Mark All as Read</button>
      </div>
    </div>
  );
}
function makeExchMRDoc(exchName,entries,date){
  const doc=new jsPDF({unit:"mm",format:"a4"});
  const W=210,pad=18;
  let y=18;
  function line(text,x,yy,size,style,align){doc.setFontSize(size||11);doc.setFont("helvetica",style||"normal");doc.text(String(text),x,yy,{align:align||"left"});}
  function hline(yy){doc.setDrawColor(200);doc.line(pad,yy,W-pad,yy);}
  line("NKD BAJAJ",W/2,y,18,"bold","center");y+=7;
  line("Authorised Bajaj Dealer | Dhanbad",W/2,y,9,"normal","center");y+=5;
  hline(y);y+=6;
  line("EXCHANGE SETTLEMENT MR",W/2,y,14,"bold","center");y+=8;
  hline(y);y+=6;
  line("Exchanger: "+exchName,pad,y,11,"bold");line("Date: "+fd(date),W-pad,y,10,"normal","right");y+=10;
  hline(y);y+=5;
  // Table header
  const cols=[pad,60,100,130,158,185];
  const hdrs=["Customer","Vehicle / Reg No","Exch Value","Amt Rec'd","Commission","Discount"];
  hdrs.forEach((h,i)=>{doc.setFontSize(9);doc.setFont("helvetica","bold");doc.text(h,cols[i],y);});
  y+=5;hline(y);y+=5;
  let totExv=0,totRec=0,totComm=0,totDisc=0;
  entries.forEach(e=>{
    doc.setFontSize(9);doc.setFont("helvetica","normal");
    doc.text(String(e.name||"").substring(0,20),cols[0],y);
    doc.text((String(e.model||"").substring(0,14)+"\n"+(e.regNo||"")).trim(),cols[1],y);
    doc.text(fc(e.exv),cols[2],y,{align:"left"});
    doc.text(fc(e.amtRec),cols[3],y,{align:"left"});
    doc.text(fc(e.comm),cols[4],y,{align:"left"});
    doc.text(fc(e.disc),cols[5],y,{align:"left"});
    totExv+=Number(e.exv||0);totRec+=Number(e.amtRec||0);totComm+=Number(e.comm||0);totDisc+=Number(e.disc||0);
    y+=8;if(y>270){doc.addPage();y=20;}
  });
  hline(y);y+=5;
  doc.setFontSize(10);doc.setFont("helvetica","bold");
  doc.text("TOTALS",cols[0],y);
  doc.text(fc(totExv),cols[2],y);doc.text(fc(totRec),cols[3],y);doc.text(fc(totComm),cols[4],y);doc.text(fc(totDisc),cols[5],y);
  y+=10;hline(y);y+=10;
  line("Exchanger Signature",pad,y+16,9);line("Authorised Signatory",W-pad,y+16,9,"normal","right");
  line("______________________",pad,y+14,9);line("______________________",W-pad,y+14,9,"normal","right");
  line("NKD Bajaj, Dhanbad",W-pad,y+22,8,"italic","right");
  return doc;
}
function ExchangerDue({custs,onUpd,notify}){
  const exchCusts=custs.filter(c=>c.billed&&c.billing&&Number(c.billing.exv||c.billing.calc?.exv||0)>0&&!c.exchMrIssued);
  const [edits,setEdits]=useState({});
  const [collapsed,setCollapsed]=useState({});
  function getExchName(c){return c.billing?.details?.exchangeName||c.exchangeName||"Unknown Exchanger";}
  function getExv(c){return Number(c.billing?.exv||c.billing?.calc?.exv||0);}
  const grouped={};
  exchCusts.forEach(c=>{const n=getExchName(c);if(!grouped[n])grouped[n]=[];grouped[n].push(c);});
  function setE(id,field,val){setEdits(p=>({...p,[id]:{...(p[id]||{}),[ field]:val}}));}
  function getE(c,field){return edits[c.id]?.[field]??c[field]??"";}
  function issueMR(exchName,entries){
    const num=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;
    const rows=entries.map(c=>({name:c.name,model:c.model||"",regNo:c.billing?.details?.exchangeRegNo||c.exchangeRegNo||"",exv:getExv(c),amtRec:Number(getE(c,"exchAmtRec")||0),comm:Number(getE(c,"exchComm")||0),disc:Number(getE(c,"exchDisc")||0)}));
    const doc=makeExchMRDoc(exchName,rows,td());
    entries.forEach(c=>{
      onUpd(c.id,{exchAmtRec:Number(getE(c,"exchAmtRec")||0),exchComm:Number(getE(c,"exchComm")||0),exchDisc:Number(getE(c,"exchDisc")||0),exchMrIssued:true,exchMrDate:td()});
    });
    const efname="ExchMR_"+exchName.replace(/ /g,"_")+"_"+td()+".pdf";
    sharePdf(doc,efname,num,"Exchange Settlement MR for "+exchName+" — "+entries.length+" vehicle(s)");
    savePdfToDrive(doc,efname,exchName,"ExchMR");
    entries.forEach(c=>savePdfToDrive(doc,efname,c.name,"ExchMR"));
    notify("✅ MR issued & sent to office for "+exchName);
  }
  if(Object.keys(grouped).length===0)return(<div style={{maxWidth:900}}><div style={{fontWeight:800,fontSize:18,color:"#1e293b",marginBottom:16}}>🔄 Exchanger Due</div><div style={{textAlign:"center",color:"#94a3b8",padding:60,fontSize:14}}>No pending exchange settlements</div></div>);
  return(
    <div style={{maxWidth:960}}>
      <div style={{fontWeight:800,fontSize:18,color:"#1e293b",marginBottom:4}}>🔄 Exchanger Due</div>
      <div style={{color:"#64748b",fontSize:12,marginBottom:18}}>Exchange customers pending settlement. Amounts are editable — update and issue MR when settled.</div>
      {Object.entries(grouped).map(([exchName,entries])=>{
        const isOpen=!collapsed[exchName];
        const totExv=entries.reduce((s,c)=>s+getExv(c),0);
        const totRec=entries.reduce((s,c)=>s+Number(getE(c,"exchAmtRec")||0),0);
        return(
          <div key={exchName} style={{background:"#fff",border:"1px solid #6b8fb5",borderRadius:14,marginBottom:18,overflow:"hidden"}}>
            {/* Exchanger header */}
            <div onClick={()=>setCollapsed(p=>({...p,[exchName]:isOpen}))} style={{background:"rgba(245,158,11,0.08)",borderBottom:isOpen?"1px solid #6b8fb5":"none",padding:"13px 18px",display:"flex",justifyContent:"space-between",alignItems:"center",cursor:"pointer"}}>
              <div>
                <div style={{fontWeight:800,fontSize:15,color:"#92400e"}}>🏪 {exchName}{(()=>{const ph=entries[0]?.billing?.details?.exchangePhone||entries[0]?.exchangePhone||"";return ph?<span style={{fontSize:12,fontWeight:500,color:"#78716c",marginLeft:10}}>📞 {ph}</span>:null;})()}</div>
                <div style={{fontSize:11,color:"#78716c",marginTop:2}}>{entries.length} vehicle{entries.length>1?"s":""} · Exchange Due: <b>{fc(totExv)}</b>{totRec>0?<> · Received: <b style={{color:"#16a34a"}}>{fc(totRec)}</b></>:null}</div>
              </div>
              <div style={{display:"flex",gap:10,alignItems:"center"}}>
                <button onClick={e=>{e.stopPropagation();issueMR(exchName,entries);}} style={{padding:"9px 18px",background:"linear-gradient(135deg,#f59e0b,#d97706)",border:"none",borderRadius:10,color:"#fff",fontWeight:700,fontSize:12,cursor:"pointer"}}>📄 Issue MR → Office</button>
                <span style={{fontSize:16,color:"#94a3b8"}}>{isOpen?"▲":"▼"}</span>
              </div>
            </div>
            {/* Customer rows */}
            {isOpen&&(
              <div style={{overflowX:"auto"}}>
                <table style={{width:"100%",borderCollapse:"collapse",minWidth:700}}>
                  <thead><tr style={{background:"#f8fafc"}}>
                    {["Customer","Vehicle","Reg No","Exch Value","Amt Rec'd","Commission","Disc Allowed","📞"].map(h=><th key={h} style={{padding:"9px 12px",fontSize:11,color:"#64748b",fontWeight:700,textAlign:"left",borderBottom:"1px solid #6b8fb5",whiteSpace:"nowrap"}}>{h}</th>)}
                  </tr></thead>
                  <tbody>{entries.map(c=>{
                    const exv=getExv(c);
                    const due=exv-Number(getE(c,"exchAmtRec")||0);
                    return(
                      <tr key={c.id} style={{borderBottom:"1px solid #f1f5f9"}}>
                        <td style={{padding:"10px 12px"}}>
                          <div style={{fontWeight:700,fontSize:13,color:"#1e293b"}}>{c.name}</div>
                          <div style={{fontSize:10,color:"#94a3b8"}}>{c.billing?.deliveryDate?fd(c.billing.deliveryDate):fd(c.billedDate)}</div>
                        </td>
                        <td style={{padding:"10px 12px",fontSize:12,color:"#475569"}}>{c.billing?.details?.exchangeAsked||c.exchangeAsked||c.model||"—"}</td>
                        <td style={{padding:"10px 12px",fontSize:12,color:"#475569"}}>{c.billing?.details?.exchangeRegNo||c.exchangeRegNo||"—"}</td>
                        <td style={{padding:"10px 12px",fontSize:13,fontWeight:800,color:"#1e293b"}}>{fc(exv)}</td>
                        <td style={{padding:"8px 10px"}}>
                          <input type="number" inputMode="numeric" value={getE(c,"exchAmtRec")} onChange={e=>setE(c.id,"exchAmtRec",e.target.value)} onBlur={e=>onUpd(c.id,{exchAmtRec:Number(e.target.value||0)})} style={{width:90,padding:"6px 8px",border:"1px solid #6b8fb5",borderRadius:7,fontSize:12,fontWeight:700,color:due<=0?"#16a34a":"#ef4444"}} placeholder="0"/>
                          {due>0&&<div style={{fontSize:9,color:"#ef4444",marginTop:2}}>Due: {fc(due)}</div>}
                          {due<=0&&Number(getE(c,"exchAmtRec")||0)>0&&<div style={{fontSize:9,color:"#16a34a",marginTop:2}}>✓ Settled</div>}
                        </td>
                        <td style={{padding:"8px 10px"}}>
                          <input type="number" inputMode="numeric" value={getE(c,"exchComm")} onChange={e=>setE(c.id,"exchComm",e.target.value)} onBlur={e=>onUpd(c.id,{exchComm:Number(e.target.value||0)})} style={{width:90,padding:"6px 8px",border:"1px solid #6b8fb5",borderRadius:7,fontSize:12}} placeholder="0"/>
                        </td>
                        <td style={{padding:"8px 10px"}}>
                          <input type="number" inputMode="numeric" value={getE(c,"exchDisc")} onChange={e=>setE(c.id,"exchDisc",e.target.value)} onBlur={e=>onUpd(c.id,{exchDisc:Number(e.target.value||0)})} style={{width:90,padding:"6px 8px",border:"1px solid #6b8fb5",borderRadius:7,fontSize:12}} placeholder="0"/>
                        </td>
                        <td style={{padding:"8px 12px"}}>
                          {(()=>{const ph=c.billing?.details?.exchangePhone||c.exchangePhone||"";return ph?(<div style={{display:"flex",flexDirection:"column",gap:4}}><div style={{fontSize:11,fontWeight:600,color:"#1e293b"}}>{ph}</div><a href={"https://wa.me/91"+ph.replace(/\D/g,"")} target="_blank" rel="noreferrer" style={{fontSize:10,color:"#22c55e",fontWeight:700,textDecoration:"none"}}>💬 WhatsApp</a></div>):(<span style={{fontSize:11,color:"#cbd5e1"}}>—</span>);})()}
                        </td>
                      </tr>
                    );
                  })}</tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}
function CashBook({custs}){
  const [date,setDate]=useState(td());
  const [branch,setBranch]=useState("All");
  // Collect all payments for the selected date
  const entries=[];
  custs.forEach(c=>{
    const br=SM_BRANCH[c.salesman]||c.branch||"—";
    // Booking payment
    if(c.booking&&c.booking.date===date){
      entries.push({type:"Booking",name:c.name,model:c.model||"",amt:Number(c.booking.amt||0),mode:c.booking.mode||"Cash",branch:br,salesman:c.salesman||"",time:""});
    }
    // Received payments (from billing)
    if(c.billing&&c.billing.payments){
      c.billing.payments.forEach(p=>{
        if(p.date===date&&Number(p.amt||0)>0){
          entries.push({type:"Payment",name:c.name,model:c.model||"",amt:Number(p.amt),mode:p.mode||"Cash",branch:br,salesman:c.salesman||"",time:p.time||""});
        }
      });
    }
  });
  const filtered=branch==="All"?entries:entries.filter(e=>e.branch===branch);
  const byBranch={};
  filtered.forEach(e=>{if(!byBranch[e.branch])byBranch[e.branch]=[];byBranch[e.branch].push(e);});
  const total=filtered.reduce((s,e)=>s+e.amt,0);
  const modeTotal={};filtered.forEach(e=>{modeTotal[e.mode]=(modeTotal[e.mode]||0)+e.amt;});
  return(
    <div style={{maxWidth:900}}>
      <div style={{fontWeight:800,fontSize:18,color:"#1e293b",marginBottom:16}}>💰 Cash Book</div>
      <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
        <input type="date" value={date} max={td()} onChange={e=>setDate(e.target.value)} style={{...inp,width:"auto",padding:"8px 12px",fontSize:13}}/>
        <div style={{display:"flex",gap:6}}>
          {["All",...BRANCHES].map(b=><button key={b} onClick={()=>setBranch(b)} style={{padding:"8px 14px",borderRadius:10,border:"1px solid "+(branch===b?"#3b82f6":"#6b8fb5"),background:branch===b?"#dbeafe":"#f8fafc",color:branch===b?"#1d4ed8":"#64748b",fontWeight:700,fontSize:12,cursor:"pointer"}}>{b}</button>)}
        </div>
      </div>
      {/* Summary cards */}
      <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
        <div style={{background:"linear-gradient(135deg,#22c55e,#16a34a)",borderRadius:14,padding:"14px 20px",color:"#fff",minWidth:140}}>
          <div style={{fontSize:11,fontWeight:600,opacity:0.85}}>TOTAL RECEIVED</div>
          <div style={{fontSize:22,fontWeight:900}}>{fc(total)}</div>
          <div style={{fontSize:11,opacity:0.8}}>{filtered.length} transactions</div>
        </div>
        {Object.entries(modeTotal).map(([m,a])=>(
          <div key={m} style={{background:"#fff",border:"1px solid #6b8fb5",borderRadius:14,padding:"14px 20px",minWidth:120}}>
            <div style={{fontSize:11,color:"#94a3b8",fontWeight:600}}>{m.toUpperCase()}</div>
            <div style={{fontSize:18,fontWeight:800,color:"#1e293b"}}>{fc(a)}</div>
          </div>
        ))}
      </div>
      {filtered.length===0?<div style={{textAlign:"center",color:"#94a3b8",padding:40,fontSize:14}}>No payments recorded for {fd(date)}</div>:
      Object.entries(byBranch).map(([br,rows])=>(
        <div key={br} style={{background:"#fff",border:"1px solid #6b8fb5",borderRadius:14,marginBottom:16,overflow:"hidden"}}>
          <div style={{background:"rgba(96,165,250,0.08)",borderBottom:"1px solid #6b8fb5",padding:"10px 16px",display:"flex",justifyContent:"space-between",alignItems:"center"}}>
            <div style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>🏢 {br}</div>
            <div style={{fontWeight:800,color:"#22c55e",fontSize:14}}>{fc(rows.reduce((s,r)=>s+r.amt,0))}</div>
          </div>
          <table style={{width:"100%",borderCollapse:"collapse"}}>
            <thead><tr style={{background:"#f8fafc"}}>
              {["Type","Customer","Model","Amount","Mode","By","Time"].map(h=><th key={h} style={{padding:"8px 12px",fontSize:11,color:"#64748b",fontWeight:700,textAlign:"left",borderBottom:"1px solid #6b8fb5"}}>{h}</th>)}
            </tr></thead>
            <tbody>{rows.map((r,i)=>(
              <tr key={i} style={{borderBottom:"1px solid #f1f5f9"}}>
                <td style={{padding:"9px 12px",fontSize:12}}><span style={{background:r.type==="Booking"?"rgba(139,92,246,0.1)":"rgba(34,197,94,0.1)",color:r.type==="Booking"?"#8b5cf6":"#16a34a",borderRadius:6,padding:"2px 8px",fontWeight:700,fontSize:11}}>{r.type}</span></td>
                <td style={{padding:"9px 12px",fontSize:12,fontWeight:600,color:"#1e293b"}}>{r.name}</td>
                <td style={{padding:"9px 12px",fontSize:11,color:"#64748b"}}>{r.model}</td>
                <td style={{padding:"9px 12px",fontSize:13,fontWeight:800,color:"#22c55e"}}>{fc(r.amt)}</td>
                <td style={{padding:"9px 12px",fontSize:12,color:"#475569"}}>{r.mode}</td>
                <td style={{padding:"9px 12px",fontSize:11,color:"#64748b"}}>{r.salesman}</td>
                <td style={{padding:"9px 12px",fontSize:11,color:"#94a3b8"}}>{r.time}</td>
              </tr>
            ))}</tbody>
          </table>
        </div>
      ))}
    </div>
  );
}
function OwnerPortal({custs,stockData,billedChassis,statusData,role,user,mBr,saveStockData,saveStatusData,nkdUsers,onSaveUsers,notify,onUpd,onLogout,onMobile}){
  const [view,setView]=useState(role==="admin"?"uploads":"dashboard");
  const billed=custs.filter(c=>c.billed);
  const thisM=td().slice(0,7);
  const billedThisM=billed.filter(c=>(c.billedDate||"").startsWith(thisM));
  const revThisM=billedThisM.reduce((s,c)=>s+((c.billing&&c.billing.calc&&c.billing.calc.E)||0),0);
  const branches=BRANCHES.map(b=>({name:b,enq:custs.filter(c=>(c.branch||SM_BRANCH[c.salesman])===b).length,book:custs.filter(c=>(c.branch||SM_BRANCH[c.salesman])===b&&c.booking).length,bill:billed.filter(c=>(c.branch||SM_BRANCH[c.salesman])===b).length,rev:billed.filter(c=>(c.branch||SM_BRANCH[c.salesman])===b).reduce((s,c)=>s+((c.billing&&c.billing.calc&&c.billing.calc.E)||0),0)}));
  const modelMap={};billed.forEach(c=>{const m=c.model||"Unknown";modelMap[m]=(modelMap[m]||0)+1;});
  const modelWise=Object.entries(modelMap).sort((a,b)=>b[1]-a[1]);
  const smMap={};custs.forEach(c=>{if(!smMap[c.salesman])smMap[c.salesman]={enq:0,book:0,bill:0,rev:0,branch:SM_BRANCH[c.salesman]||""};smMap[c.salesman].enq++;if(c.booking)smMap[c.salesman].book++;if(c.billed){smMap[c.salesman].bill++;smMap[c.salesman].rev+=((c.billing&&c.billing.calc&&c.billing.calc.E)||0);}});
  const smPerf=Object.entries(smMap).sort((a,b)=>b[1].bill-a[1].bill);
  const navItems=role==="admin"
    ?[{id:"uploads",l:"Uploads & Data",ic:"📤"},{id:"vault",l:"Document Vault",ic:"📁"}]
    :[{id:"dashboard",l:"Dashboard",ic:"📊"},{id:"customers",l:"All Customers",ic:"👥"},{id:"team",l:"Team Performance",ic:"👔"},{id:"cashbook",l:"Cash Book",ic:"💰"},{id:"exchdue",l:"Exchanger Due",ic:"🔄"},{id:"stock",l:"Stock & Ageing",ic:"🏍️"},{id:"uploads",l:"Uploads",ic:"📤"},{id:"rcstatus",l:"RC / HSRP",ic:"📋"},{id:"reports",l:"Reports",ic:"📄"},{id:"users",l:"User Accounts",ic:"👤"},{id:"vault",l:"Document Vault",ic:"📁"}];
  // tech = full owner powers
  const SB=({label})=>(<th style={{fontSize:11,color:"#64748b",fontWeight:700,textAlign:"left",padding:"7px 12px",borderBottom:"2px solid #6b8fb5",background:"#f8fafc"}}>{label}</th>);
  const TD=({v,col,bold})=>(<td style={{padding:"8px 12px",fontSize:13,color:col||"#1e293b",fontWeight:bold?700:400,borderBottom:"1px solid #e8eef8"}}>{v}</td>);
  return(
    <div style={{display:"flex",minHeight:"100vh",background:"linear-gradient(160deg,#f0f7ff,#f8fafc)",fontFamily:"'Sora','Inter',sans-serif",color:"#1e293b"}}>
      <style>{`@import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&display=swap');*{font-family:'Sora','Inter',sans-serif!important}body{background:#f0f7ff}::-webkit-scrollbar{width:6px;height:6px}::-webkit-scrollbar-track{background:#f1f5f9}::-webkit-scrollbar-thumb{background:#6b8fb5;border-radius:3px}`}</style>
      {/* ─── SIDEBAR ─── */}
      <div style={{width:220,background:"#ffffff",borderRight:"2px solid #6b8fb5",display:"flex",flexDirection:"column",position:"fixed",top:0,bottom:0,left:0,zIndex:50,boxShadow:"2px 0 12px rgba(15,23,42,.08)"}}>
        <div style={{padding:"18px 16px 14px",borderBottom:"2px solid #6b8fb5"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:8}}>
            <div style={{width:40,height:40,borderRadius:12,background:"linear-gradient(135deg,#f97316,#ef4444)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:18,color:"#fff"}}>B</div>
            <div><div style={{fontWeight:800,fontSize:14,color:"#1e293b"}}>NKD BAJAJ</div><div style={{fontSize:10,color:"#64748b",marginTop:1}}>{role==="owner"?"Owner Portal":role==="tech"?"Tech Portal":"Admin Portal"}</div></div>
          </div>
          <div style={{fontSize:11,color:"#94a3b8",background:"#f1f5f9",borderRadius:7,padding:"4px 8px"}}>{user} · {role}</div>
        </div>
        <div style={{flex:1,padding:"12px 10px",overflowY:"auto"}}>
          {navItems.map(n=>(
            <button key={n.id} onClick={()=>setView(n.id)} style={{width:"100%",display:"flex",alignItems:"center",gap:10,padding:"10px 12px",borderRadius:10,border:"none",cursor:"pointer",marginBottom:3,background:view===n.id?"#dbeafe":"transparent",color:view===n.id?"#1d4ed8":"#475569",fontWeight:view===n.id?700:500,fontSize:13,textAlign:"left",transition:"background .15s"}}>
              <span style={{fontSize:18,minWidth:22}}>{n.ic}</span>{n.l}
            </button>
          ))}
        </div>
        <div style={{padding:"12px 10px",borderTop:"2px solid #6b8fb5",display:"flex",flexDirection:"column",gap:7}}>
          <button onClick={onMobile} style={{padding:"8px 12px",borderRadius:8,border:"1px solid #6b8fb5",background:"transparent",color:"#475569",fontSize:12,cursor:"pointer",fontWeight:600}}>📱 Mobile View</button>
          <button onClick={onLogout} style={{padding:"8px 12px",borderRadius:8,border:"1px solid #ef4444",background:"transparent",color:"#ef4444",fontSize:12,cursor:"pointer",fontWeight:700}}>🚪 Logout</button>
        </div>
      </div>
      {/* ─── MAIN CONTENT ─── */}
      <div style={{marginLeft:220,flex:1,padding:"28px 32px",overflowY:"auto",minHeight:"100vh"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:24}}>
          <div style={{fontWeight:800,fontSize:24,color:"#1e293b"}}>{navItems.find(n=>n.id===view)?.ic} {navItems.find(n=>n.id===view)?.l}</div>
          <div style={{fontSize:12,color:"#94a3b8",background:"#fff",border:"1px solid #6b8fb5",borderRadius:8,padding:"5px 12px"}}>{new Date().toLocaleDateString("en-IN",{weekday:"long",day:"2-digit",month:"long",year:"numeric"})}</div>
        </div>

        {/* ── DASHBOARD ── */}
        {view==="dashboard"&&<>
          <div style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:16,marginBottom:24}}>
            {[[" Total Enquiries",custs.length,"#3b82f6","👥"],[" Bookings",custs.filter(c=>c.booking).length,"#8b5cf6","📝"],[" Total Billed",billed.length,"#22c55e","✅"],[" Total Revenue",fc(billed.reduce((s,c)=>s+((c.billing&&c.billing.calc&&c.billing.calc.E)||0),0)),"#f97316","💰"]].map(([l,v,col,ic])=>(
              <div key={l} style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:16,padding:"20px 22px",boxShadow:"0 2px 14px rgba(15,23,42,.06)"}}>
                <div style={{fontSize:28,marginBottom:8}}>{ic}</div>
                <div style={{fontSize:28,fontWeight:900,color:col,lineHeight:1}}>{v}</div>
                <div style={{fontSize:12,color:"#64748b",marginTop:6,fontWeight:600}}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:24}}>
            {[["Billed This Month",billedThisM.length,"#22c55e"],["Revenue This Month",fc(revThisM),"#f97316"],["Hot + Warm Leads",custs.filter(c=>!c.billed&&!c.stopped&&["Hot","Warm"].includes(c.status)).length,"#8b5cf6"]].map(([l,v,col])=>(
              <div key={l} style={{background:"#fff",border:"2px solid "+col+"50",borderRadius:14,padding:"16px 20px"}}>
                <div style={{fontSize:22,fontWeight:900,color:col}}>{v}</div>
                <div style={{fontSize:12,color:"#64748b",marginTop:4,fontWeight:600}}>{l}</div>
              </div>
            ))}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:20,marginBottom:24}}>
            <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,padding:"18px 20px"}}>
              <div style={{fontWeight:800,fontSize:15,color:"#1e293b",marginBottom:14}}>🏪 Branch Performance</div>
              <table style={{width:"100%",borderCollapse:"collapse"}}>
                <thead><tr><SB label="Branch"/><SB label="Enquiries"/><SB label="Bookings"/><SB label="Billed"/><SB label="Revenue"/></tr></thead>
                <tbody>{branches.map(b=><tr key={b.name}><TD v={b.name} bold/><TD v={b.enq} col="#3b82f6"/><TD v={b.book} col="#8b5cf6"/><TD v={b.bill} col="#22c55e" bold/><TD v={fc(b.rev)} col="#f97316" bold/></tr>)}</tbody>
              </table>
            </div>
            <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,padding:"18px 20px"}}>
              <div style={{fontWeight:800,fontSize:15,color:"#1e293b",marginBottom:14}}>🏍️ Top Models Billed</div>
              {modelWise.slice(0,10).map(([m,cnt])=>(
                <div key={m} style={{display:"flex",alignItems:"center",gap:10,marginBottom:9}}>
                  <div style={{flex:1,fontSize:12,color:"#1e293b",fontWeight:600,minWidth:0,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m}</div>
                  <div style={{width:120,height:8,background:"#f1f5f9",borderRadius:4,flexShrink:0}}><div style={{height:8,background:"linear-gradient(90deg,#f97316,#ef4444)",borderRadius:4,width:(cnt/(modelWise[0]?.[1]||1)*100)+"%"}}/></div>
                  <div style={{fontSize:12,fontWeight:800,color:"#f97316",minWidth:20,textAlign:"right"}}>{cnt}</div>
                </div>
              ))}
            </div>
          </div>
          <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,padding:"18px 20px"}}>
            <div style={{fontWeight:800,fontSize:15,color:"#1e293b",marginBottom:14}}>👔 Team Performance</div>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <thead><tr><SB label="Salesman"/><SB label="Branch"/><SB label="Enquiries"/><SB label="Bookings"/><SB label="Billed"/><SB label="Revenue"/></tr></thead>
              <tbody>{smPerf.map(([s,d])=><tr key={s}><TD v={s} bold/><TD v={d.branch} col="#64748b"/><TD v={d.enq} col="#3b82f6"/><TD v={d.book} col="#8b5cf6"/><TD v={d.bill} col="#22c55e" bold/><TD v={fc(d.rev)} col="#f97316" bold/></tr>)}</tbody>
            </table>
          </div>
        </>}

        {/* ── ALL CUSTOMERS TABLE ── */}
        {view==="customers"&&(()=>{
          const [cq,setCq]=useState("");
          const filt=cq.trim().length<2?custs:custs.filter(c=>(c.name+" "+c.phone+" "+(c.model||"")+" "+(c.salesman||"")).toLowerCase().includes(cq.toLowerCase()));
          return(<>
            <input placeholder="🔍 Search name, phone, model, salesman…" value={cq} onChange={e=>setCq(e.target.value)} style={{width:"100%",padding:"11px 16px",border:"2px solid #6b8fb5",borderRadius:10,fontSize:14,marginBottom:16,boxSizing:"border-box",outline:"none"}}/>
            <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,overflow:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",minWidth:900}}>
                <thead><tr><SB label="Name"/><SB label="Phone"/><SB label="Model"/><SB label="Status"/><SB label="Salesman"/><SB label="Branch"/><SB label="Enquiry"/><SB label="Billed"/><SB label="Revenue"/></tr></thead>
                <tbody>{filt.slice(0,200).map(c=>{const b=c.billing;const cl=b&&b.calc?b.calc:null;return(<tr key={c.id} style={{cursor:"pointer"}} onClick={()=>{}}><TD v={c.name} bold/><TD v={c.phone} col="#64748b"/><TD v={c.model||"—"}/><TD v={<span style={{fontSize:10,fontWeight:800,background:({"Hot":"#ef4444","Warm":"#f97316","Billed":"#22c55e","Booked":"#8b5cf6","Cold":"#3b82f6","Lost":"#64748b"}[c.status]||"#374151")+"20",color:({"Hot":"#ef4444","Warm":"#f97316","Billed":"#22c55e","Booked":"#8b5cf6","Cold":"#3b82f6","Lost":"#64748b"}[c.status]||"#374151"),padding:"2px 8px",borderRadius:8}}>{c.status}</span>}/><TD v={c.salesman||"—"} col="#64748b"/><TD v={(c.branch||SM_BRANCH[c.salesman]||"—")} col="#64748b"/><TD v={fd(c.enquiryDate)} col="#94a3b8"/><TD v={c.billed?fd(c.billedDate):"—"} col={c.billed?"#22c55e":"#94a3b8"}/><TD v={cl?fc(cl.E):"—"} col="#f97316" bold/></tr>);})}</tbody>
              </table>
              {filt.length>200&&<div style={{padding:"10px 16px",fontSize:12,color:"#94a3b8"}}>Showing 200 of {filt.length} — refine search</div>}
            </div>
          </>);
        })()}

        {/* ── TEAM VIEW ── */}
        {view==="team"&&<div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:16}}>
          {smPerf.map(([s,d])=>(
            <div key={s} style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:14,padding:"18px 20px"}}>
              <div style={{fontWeight:800,fontSize:16,color:"#1e293b",marginBottom:2}}>{s}</div>
              <div style={{fontSize:12,color:"#64748b",marginBottom:14}}>📍 {d.branch}</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                {[["Enquiries",d.enq,"#3b82f6"],["Bookings",d.book,"#8b5cf6"],["Billed",d.bill,"#22c55e"],["Revenue",fc(d.rev),"#f97316"]].map(([l,v,col])=>(
                  <div key={l} style={{background:"#f8fafc",border:"1px solid #6b8fb5",borderRadius:10,padding:"10px 12px"}}>
                    <div style={{fontSize:18,fontWeight:900,color:col}}>{v}</div>
                    <div style={{fontSize:10,color:"#64748b",marginTop:2}}>{l}</div>
                  </div>
                ))}
              </div>
              <div style={{marginTop:12,background:"#f8fafc",borderRadius:8,height:8}}><div style={{height:8,background:"linear-gradient(90deg,#22c55e,#3b82f6)",borderRadius:8,width:(d.bill?Math.min((d.bill/Math.max(...smPerf.map(([,x])=>x.bill)))*100,100):0)+"%"}}/></div>
              <div style={{fontSize:10,color:"#94a3b8",marginTop:4}}>Billing progress vs team</div>
            </div>
          ))}
        </div>}

        {/* ── STOCK ── */}
        {view==="stock"&&<div style={{maxWidth:900}}><StockView stockData={stockData} billedChassis={billedChassis} role={role} userBranch={null} notify={notify}/></div>}

        {/* ── UPLOADS ── */}
        {view==="uploads"&&(role==="admin"?(
          <div style={{display:"flex",flexDirection:"column",gap:24,maxWidth:900}}>
            {/* ── STOCK SECTION ── */}
            <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:16,padding:"20px 22px"}}>
              <div style={{fontWeight:800,fontSize:16,color:"#1e293b",marginBottom:4}}>🏍️ Stock Statement</div>
              <div style={{fontSize:11,color:"#94a3b8",marginBottom:12}}>Branch-wise vehicle stock from dealership</div>
              {/* Upload row */}
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
                {stockData.length>0&&<span style={{fontSize:12,color:"#22c55e",fontWeight:600,background:"rgba(34,197,94,0.08)",border:"1px solid rgba(34,197,94,0.25)",borderRadius:8,padding:"4px 10px"}}>✅ {stockData.length} vehicles loaded</span>}
                <label style={{background:"rgba(52,211,153,0.08)",border:"1px dashed rgba(52,211,153,0.5)",borderRadius:10,padding:"8px 16px",cursor:"pointer",fontSize:12,color:"#34d399",fontWeight:700,whiteSpace:"nowrap"}}>
                  {stockData.length>0?"🔄 Replace Excel":"📂 Choose Excel File"}
                  <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files&&e.target.files[0]){parseExcel(e.target.files[0],d=>{saveStockData(d);notify("✅ Stock uploaded — "+d.length+" vehicles");},err=>notify("❌ "+err));e.target.value="";}}}/>
                </label>
                <span style={{fontSize:10,color:"#94a3b8"}}>.xlsx · .xls · .csv</span>
              </div>
              {/* Data view */}
              <StockView stockData={stockData} billedChassis={billedChassis} role={role} userBranch={null} notify={notify}/>
            </div>
            {/* ── RC/HSRP SECTION ── */}
            <div style={{background:"#fff",border:"2px solid #6b8fb5",borderRadius:16,padding:"20px 22px"}}>
              <div style={{fontWeight:800,fontSize:16,color:"#1e293b",marginBottom:4}}>📋 RC / HSRP Status</div>
              <div style={{fontSize:11,color:"#94a3b8",marginBottom:12}}>RC and HSRP status report from RTO</div>
              {/* Upload row */}
              <div style={{display:"flex",alignItems:"center",gap:12,marginBottom:16,flexWrap:"wrap"}}>
                {statusData.length>0&&<span style={{fontSize:12,color:"#22c55e",fontWeight:600,background:"rgba(34,197,94,0.08)",border:"1px solid rgba(34,197,94,0.25)",borderRadius:8,padding:"4px 10px"}}>✅ {statusData.length} records loaded</span>}
                <label style={{background:"rgba(96,165,250,0.08)",border:"1px dashed rgba(96,165,250,0.5)",borderRadius:10,padding:"8px 16px",cursor:"pointer",fontSize:12,color:"#60a5fa",fontWeight:700,whiteSpace:"nowrap"}}>
                  {statusData.length>0?"🔄 Replace Excel":"📂 Choose Excel File"}
                  <input type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={e=>{if(e.target.files&&e.target.files[0]){parseExcel(e.target.files[0],d=>{saveStatusData(d);notify("✅ RC/HSRP uploaded — "+d.length+" records");},err=>notify("❌ "+err));e.target.value="";}}}/>
                </label>
                <span style={{fontSize:10,color:"#94a3b8"}}>.xlsx · .xls · .csv</span>
              </div>
              {/* Data view */}
              <RCHSRPSearch statusData={statusData} role={role} onUpload={saveStatusData} notify={notify}/>
            </div>
          </div>
        ):<div style={{maxWidth:700}}><UploadsHub stockData={stockData} statusData={statusData} onStockUpload={saveStockData} onStatusUpload={saveStatusData} notify={notify}/></div>)}

        {/* ── RC/HSRP ── */}
        {view==="rcstatus"&&<div style={{maxWidth:900}}><RCHSRPSearch statusData={statusData} role={role} onUpload={saveStatusData} notify={notify}/></div>}

        {/* ── REPORTS ── */}
        {view==="reports"&&<div style={{maxWidth:800}}><Reports custs={custs} onImportCust={()=>{}}/></div>}

        {/* ── CASH BOOK ── */}
        {view==="cashbook"&&<CashBook custs={custs}/>}
        {view==="exchdue"&&<ExchangerDue custs={custs} onUpd={onUpd} notify={notify}/>}

        {/* ── USERS ── */}
        {view==="users"&&isOwner(role)&&<UserMgmt nkdUsers={nkdUsers||DEFAULT_USERS} onSave={onSaveUsers} notify={notify}/>}

        {/* ── VAULT ── */}
        {view==="vault"&&<DocVault custs={custs} onImport={()=>{}}/>}
      </div>
    </div>
  );
}

export default function App(){
  const [role,setRole]=useState(ld("nkd_r","salesman"));
  const [user,setUser]=useState(ld("nkd_u",SM[0]));
  const [custs,setCusts]=useState(seedData);
  const [fbReady,setFbReady]=useState(false);
  const [view,setView]=useState("dashboard");
  const [sel,setSel]=useState(null);
  const [toast,setToast]=useState(null);
  const [li,setLi]=useState(ld("nkd_li",false));
  const [addOpen,setAddOpen]=useState(false);
  const [billOpen,setBillOpen]=useState(false);
  const [bookOpen,setBookOpen]=useState(false);
  const [custF,setCustF]=useState("All");
  const [dtab,setDtab]=useState(null);
  const [fSM,setFSM]=useState("All");
  const [statusData,setStatusData]=useState(()=>ld("nkd_rcstatus",[]));
  function saveStatusData(data){setStatusData(data);sv("nkd_rcstatus",data);_dbSet("nkd_rcstatus",data);}
  const [stockData,setStockData]=useState(()=>ld("nkd_stock",[]));
  function saveStockData(data){setStockData(data);sv("nkd_stock",data);_dbSet("nkd_stock",data);}
  const [nkdUsers,setNkdUsers]=useState(()=>ld("nkd_users",DEFAULT_USERS));
  function saveUsers(data){setNkdUsers(data);sv("nkd_users",data);_dbSet("nkd_users",data);}
  const [payNotifs,setPayNotifs]=useState(()=>ld("nkd_pnotifs",[]));
  const billedChassis=useMemo(()=>custs.filter(c=>c.billed&&c.billing&&c.billing.chassis).map(c=>String(c.billing.chassis).trim().toUpperCase()),[custs]);
  const stack=useRef([]);
  function nav(v){if(v!==view){stack.current.push(view);setView(v);}}
  function goBack(){const pv=stack.current.pop();setView(pv||"dashboard");}
  useEffect(()=>{
    try{
      window.history.pushState(null,"");
      const h=()=>{goBack();try{window.history.pushState(null,"");}catch(e){}};
      window.addEventListener("popstate",h);
      return()=>window.removeEventListener("popstate",h);
    }catch(e){}
  },[]);
  const alerts=custs.filter(c=>c.stopped&&c.status==="Lost"&&!c.alertDismissed);

  useEffect(()=>{
    Promise.all([
      _dbGet("custs").then(d=>{if(d&&d.length){sv("nkd6",d);setCusts(d);}}),
      _dbGet("passwords").then(d=>{if(d)sv("nkd_pw",d);}),
      _dbGet("rate_chart").then(d=>{if(d){sv("nkd_rc",d);try{Object.assign(RC,d);}catch(e){}}}),
      _dbGet("office_wa").then(d=>{if(d)sv("nkd_office_wa",d);}),
      _dbGet("nkd_users").then(d=>{if(d){sv("nkd_users",d);setNkdUsers(d);}else{_dbSet("nkd_users",DEFAULT_USERS);sv("nkd_users",DEFAULT_USERS);}}),
      _dbGet("payment_notifs").then(d=>{if(d){sv("nkd_pnotifs",d);setPayNotifs(d);}}),
      _dbGet("nkd_stock").then(d=>{if(d&&d.length){sv("nkd_stock",d);setStockData(d);}}),
      _dbGet("nkd_rcstatus").then(d=>{if(d&&d.length){sv("nkd_rcstatus",d);setStatusData(d);}}),
    ]).catch(()=>{}).finally(()=>setFbReady(true));
  },[]);

  useEffect(()=>{
    if(!fbReady)return;
    sv("nkd6",custs);
    _dbSet("custs",custs);
  },[custs,fbReady]);

  function notify(msg,type){setToast({msg,type});setTimeout(()=>setToast(null),3000);}
  function upd(id,patch){setCusts(p=>p.map(c=>c.id===id?{...c,...patch,updatedAt:td()}:c));}

  function addC(data){
    setCusts(p=>[{...data,branch:SM_BRANCH[data.salesman]||BRANCHES[0],id:"C"+Date.now(),attempts:0,stopped:false,billed:false,billedDate:null,photos:{},billing:null,managerApproval:null},...p]);
    notify("Customer added ✓");
  }

  function logF(cust,out,note,nxt,extra){
    const failed=(out==="no_response"||out==="switched_off");
    let attempts=(cust.attempts||0)+(failed?1:0),stopped=false,status=cust.status,followupDate=nxt;
    if(failed&&attempts>=3){stopped=true;status="Lost";}
    if(out==="not_interested"){status="Lost";}
    if(out==="booked")status="Booked";
    if(!followupDate&&!stopped)followupDate=aD(td(),FU[status]||3);
    const reason=(extra&&extra.reason)?(" [Reason: "+extra.reason+(extra.competitor?" — "+extra.competitor:"")+"]"):"";
    const rem="\n["+td()+"] "+out.toUpperCase()+": "+(note||"—")+reason;
    upd(cust.id,{attempts,stopped,status,followupDate,remarks:(cust.remarks||"")+rem,lostReason:(extra&&extra.reason)||cust.lostReason,lostCompetitor:(extra&&extra.competitor)||cust.lostCompetitor,alertDismissed:stopped?false:cust.alertDismissed,callLog:[...(cust.callLog||[]),{date:td(),out,note}]});
    if(stopped)notify("⚠️ Auto-stopped after 3 attempts","warn");else notify("Followup logged ✓");
  }

  function logCall(cust,dur){upd(cust.id,{callLog:[...(cust.callLog||[]),{date:td(),time:new Date().toLocaleTimeString("en-IN"),duration:dur,type:"call"}]});}
  function sendMRBoth(cust,billing,calc){
    try{
      const doc=makeMRDoc(cust,billing,calc);
      sharePdf(doc,"MR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",cust.phone,"Please find your Money Receipt from NKD Bajaj, Dhanbad.");
      savePdfToDrive(doc,"MR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",cust.name,"MR");
      const offNum=ld("nkd_office_wa",OFFICE_WA)||OFFICE_WA;
      const doc2=makeCombinedDoc(cust,billing,calc);
      sharePdf(doc2,"CalcMR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",offNum,"Calc Sheet + MR for "+cust.name+" ("+cust.model+") — Paid: "+fc(calc.paid)+" · Balance: "+fc(Math.max(calc.K,0)));
      savePdfToDrive(doc2,"CalcMR_"+cust.name.replace(/ /g,"_")+"_"+td()+".pdf",cust.name,"CalcSheet");
    }catch(e){}
  }

  function addPayment(custId,payment){
    const cust=custs.find(c=>c.id===custId);
    if(!cust||!cust.billing)return;
    const payments=[...(cust.billing.payments||[]),payment];
    const r=RC[cust.modelCode]||{};
    const newBilling={...cust.billing,payments};
    const newCalc=calcB(newBilling,r);
    const updC={...cust,billing:{...newBilling,calc:newCalc,paid:newCalc.paid},updatedAt:td()};
    setCusts(p=>p.map(c=>c.id===custId?updC:c));
    // Auto-send MR to customer + office
    sendMRBoth(updC,newBilling,newCalc);
    // Store notification for manager / owner
    const notif={id:Date.now(),custName:updC.name,model:updC.model||"",amt:payment.amt,mode:payment.mode,balance:Math.max(newCalc.K,0),salesman:updC.salesman||user,date:td(),time:new Date().toLocaleTimeString("en-IN",{hour:"2-digit",minute:"2-digit"}),read:false};
    const updN=[...payNotifs,notif];
    setPayNotifs(updN);
    sv("nkd_pnotifs",updN);
    _dbSet("payment_notifs",updN);
    notify("✅ Payment saved · MR sending to customer & office");
  }

  function billC(cust,data){
    const editLog=cust.billing?("\n["+td()+"] CALC SHEET EDITED by "+user):"";
    const updCust={...cust,remarks:(cust.remarks||"")+editLog,billed:true,billedDate:td(),status:"Billed",billing:data,billedBy:user,managerApproval:role==="salesman"?null:"approved",approvedBy:role==="salesman"?null:user,...data.details};
    upd(cust.id,updCust);
    // Auto-send MR to customer + office if any payment was received
    if(Number(data.paid||0)>0){
      const r=RC[cust.modelCode]||{};
      const calc=calcB(data,r);
      sendMRBoth({...cust,...data.details,billing:data},data,calc);
      notify("✅ Billed! MR sending to customer & office");
    }else{
      notify("✅ Billed! Receipt ready");
    }
  }

  function approveBill(id,ok,remark){
    const cu=custs.find(c=>c.id===id);
    const mr=remark?("\n["+td()+"] MANAGER: "+remark):"";
    if(ok){upd(id,{managerApproval:"approved",approvedBy:user,remarks:((cu&&cu.remarks)||"")+mr+"\n["+td()+"] APPROVED by "+user});notify("✅ Approved — record locked");}
    else{upd(id,{managerApproval:"rejected",billed:false,status:"Booked",billing:(cu&&cu.billing)?{...cu.billing,receiptHtml:null}:null,remarks:((cu&&cu.remarks)||"")+mr+"\n["+td()+"] BILLING REJECTED by Manager — correct and re-bill. All documents retained."});notify("❌ Rejected — sent back to executive");}
  }

  const mBr=ld("nkd_br",BRANCHES[0]);
  const myC=role==="salesman"?custs.filter(c=>c.salesman===user):role==="manager"?custs.filter(c=>(c.branch||SM_BRANCH[c.salesman])===mBr):custs; // owner/tech/admin see all
  const due=[...myC.filter(c=>!c.billed&&!c.stopped&&c.followupDate<=td())].sort((a,b)=>{const o={Hot:0,Warm:1,Cold:2,Booked:3};return(o[a.status]??9)-(o[b.status]??9);});
  const pending=custs.filter(c=>c.billing&&c.managerApproval===null);
  const myPending=role==="salesman"?pending.filter(c=>c.salesman===user):pending;
  const revivable=custs.filter(c=>{if(c.billed)return false;const base=c.reactivatedAt||c.enquiryDate;return((new Date()-new Date(base))/(864e5*30))>=6;});

  function openD(c,tab){setSel(c);if(tab)setDtab(tab);nav("detail");window.scrollTo({top:0,behavior:"instant"});}

  const [portalMode,setPortalMode]=useState(()=>ld("nkd_portal",false));
  function togglePortal(v){setPortalMode(v);sv("nkd_portal",v);}

  if(!fbReady)return(<div style={{minHeight:"100vh",background:"linear-gradient(160deg,#f0f7ff 0%,#f8fafc 100%)",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",gap:14}}><div style={{width:110,background:"#fff",borderRadius:16,padding:"8px 12px"}}><img src="/logo.png" alt="NKD Bajaj" style={{width:"100%"}}/></div><div style={{color:"#f97316",fontWeight:700,fontSize:15}}>NKD Bajaj CRM</div><div style={{color:"#94a3b8",fontSize:12}}>Connecting to database…</div></div>);
  if(!li)return <Login nkdUsers={nkdUsers} onLogin={(r,u,b)=>{setRole(r);setUser(u);if(b)sv("nkd_br",b);sv("nkd_r",r);sv("nkd_u",u);sv("nkd_li",true);setLi(true);if(isPortalRole(r))togglePortal(true);}}/>;

  // Payment notification popup for manager / owner / tech
  const notifPopup=(role==="manager"||role==="owner"||role==="tech"||role==="admin")&&(
    <PaymentNotifPopup notifs={payNotifs} onDismiss={()=>{
      const marked=payNotifs.map(n=>({...n,read:true}));
      setPayNotifs(marked);
      sv("nkd_pnotifs",marked);
      _dbSet("payment_notifs",marked);
    }}/>
  );

  // Owner / Admin → Web Portal (unless they switched to mobile)
  if(isPortalRole(role)&&portalMode){
    return <>{notifPopup}<OwnerPortal
      custs={custs} stockData={stockData} billedChassis={billedChassis} statusData={statusData}
      role={role} user={user} mBr={mBr}
      saveStockData={saveStockData} saveStatusData={saveStatusData}
      nkdUsers={nkdUsers} onSaveUsers={saveUsers}
      notify={notify} onUpd={upd}
      onLogout={()=>{sv("nkd_li",false);sv("nkd_portal",false);setPortalMode(false);setLi(false);}}
      onMobile={()=>togglePortal(false)}
    /></>;
  }

  const navItems=role==="admin"?[{id:"vault",l:"Document Vault",ic:"📁"},{id:"uploads",l:"Uploads",ic:"📤"},{id:"stock",l:"Stock",ic:"🏍️"},{id:"rcstatus",l:"RC/HSRP",ic:"🔍"}]:[{id:"dashboard",l:"Home",ic:"🏠"},{id:"followups",l:"Followup",ic:"📞",badge:due.length},{id:"customers",l:"Customers",ic:"👥"},{id:"stock",l:"Stock",ic:"🏍️"},{id:"rcstatus",l:"RC/HSRP",ic:"🔍"},{id:"approvals",l:role==="salesman"?"My Pending":"Approve",ic:"✅",badge:myPending.length},...(role!=="salesman"?[{id:"revival",l:"Revival",ic:"🔄"}]:[]),...(isOwner(role)?[{id:"reports",l:"Reports",ic:"📊"}]:[]),...(isOwner(role)?[{id:"vault",l:"Vault",ic:"📁"}]:[]),...(isOwner(role)?[{id:"uploads",l:"Uploads",ic:"📤"}]:[]),...(role!=="salesman"&&alerts.length>0?[{id:"alerts",l:"Alerts",ic:"⚠️",badge:alerts.length}]:[])];

  return(
    <div style={{minHeight:"100vh",background:"linear-gradient(160deg,#f0f7ff 0%,#e8f4ff 40%,#f8fafc 100%)",color:"#1e293b",fontFamily:"'Inter',-apple-system,sans-serif",maxWidth:480,margin:"0 auto"}}>
      {notifPopup}
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Sora:wght@400;600;700;800&display=swap');
        *{font-family:'Sora','Inter',-apple-system,sans-serif!important;-webkit-tap-highlight-color:transparent}
        body{background:#f0f7ff}
        @keyframes fadeUp{from{opacity:0;transform:translateY(12px)}to{opacity:1;transform:none}}
        @keyframes glow{0%,100%{box-shadow:0 0 14px rgba(249,115,22,.3)}50%{box-shadow:0 0 26px rgba(249,115,22,.5)}}
        @keyframes shimmer{0%{background-position:-200% 0}100%{background-position:200% 0}}
        .fu{animation:fadeUp .28s cubic-bezier(.2,.8,.3,1) both}
        .glass{background:#ffffff!important;border:1px solid #6b8fb5!important;box-shadow:0 2px 16px rgba(15,23,42,.14)!important;border-radius:16px}
        button{transition:transform .12s cubic-bezier(.2,.8,.3,1),box-shadow .2s,opacity .15s}
        button:active{transform:scale(.96);opacity:.92}
        input,select,textarea{background:#f8fafc;color:#1e293b;transition:border-color .18s,box-shadow .18s}
        input:focus,select:focus,textarea:focus{border-color:#f97316!important;box-shadow:0 0 0 3px rgba(249,115,22,.12)!important;background:#fff!important}
        ::-webkit-scrollbar{display:none}
        ::placeholder{color:#b0bec8}
      `}</style>
      <div style={{background:"rgba(255,255,255,.97)",backdropFilter:"blur(16px)",borderBottom:"1px solid #6b8fb5",padding:"12px 16px",display:"flex",alignItems:"center",justifyContent:"space-between",position:"sticky",top:0,zIndex:100,boxShadow:"0 2px 16px rgba(15,23,42,.14)"}}>
        <div style={{display:"flex",alignItems:"center",gap:10}}>
          <div style={{width:38,height:38,borderRadius:11,background:"linear-gradient(135deg,#f97316,#ef4444)",display:"flex",alignItems:"center",justifyContent:"center",fontWeight:900,fontSize:17,color:"#fff",animation:"glow 3s ease infinite"}}>B</div>
          <div><div style={{fontWeight:800,fontSize:13,color:"#1e293b"}}>NKD BAJAJ CRM</div><div style={{fontSize:10,color:"#94a3b8"}}>{user} · {role}{role==="manager"?" · "+mBr:""}</div></div>
        </div>
        <div style={{display:"flex",alignItems:"center",gap:6}}>
          {due.length>0&&<div style={{background:"rgba(249,115,22,0.2)",color:"#f97316",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:800}}>📞 {due.length}</div>}
          {myPending.length>0&&<div onClick={()=>nav("approvals")} style={{background:"rgba(139,92,246,0.2)",color:"#a78bfa",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:800,cursor:"pointer"}}>✓ {myPending.length}</div>}
          {role!=="salesman"&&alerts.length>0&&<div onClick={()=>nav("alerts")} style={{background:"rgba(239,68,68,0.2)",color:"#ef4444",borderRadius:20,padding:"2px 7px",fontSize:10,fontWeight:800,cursor:"pointer"}}>⚠️ {alerts.length}</div>}
          {isPortalRole(role)&&<button onClick={()=>togglePortal(true)} style={{background:"#dbeafe",border:"1px solid #3b82f6",color:"#1d4ed8",borderRadius:8,padding:"3px 8px",fontSize:10,cursor:"pointer",fontWeight:700}}>🖥️ Portal</button>}
          <button onClick={()=>{sv("nkd_li",false);setLi(false);}} style={{background:"transparent",border:"1px solid #6b8fb5",color:"#94a3b8",borderRadius:8,padding:"3px 8px",fontSize:10,cursor:"pointer"}}>Out</button>
        </div>
      </div>

      {toast&&<div style={{position:"fixed",top:66,left:"50%",transform:"translateX(-50%)",background:toast.type==="err"?"#7f1d1d":toast.type==="warn"?"#78350f":"#064e3b",color:"#fff",padding:"9px 18px",borderRadius:12,fontSize:13,fontWeight:600,zIndex:300,whiteSpace:"nowrap"}}>{toast.msg}</div>}

      <div style={{position:"fixed",bottom:0,left:0,right:0,maxWidth:480,margin:"0 auto",background:"rgba(255,255,255,.97)",backdropFilter:"blur(18px)",borderTop:"1px solid #6b8fb5",display:"flex",zIndex:100,paddingBottom:"env(safe-area-inset-bottom)",boxShadow:"0 -4px 24px rgba(15,23,42,.10)"}}>
        {navItems.map(t=><button key={t.id} onClick={()=>nav(t.id)} style={{flex:1,padding:"9px 2px 11px",background:"transparent",border:"none",cursor:"pointer",display:"flex",flexDirection:"column",alignItems:"center",gap:2}}>
          <span style={{fontSize:22}}>{t.ic}</span>
          <span style={{fontSize:11,fontWeight:700,color:view===t.id?"#f97316":"#94a3b8"}}>{t.l}</span>
          {t.badge>0&&<span style={{position:"absolute",marginLeft:26,marginTop:-2,background:"#ef4444",color:"#fff",fontSize:8,fontWeight:800,borderRadius:8,padding:"1px 4px"}}>{t.badge}</span>}
        </button>)}
      </div>

      <div style={{padding:16,paddingBottom:110}}>
        {role==="admin"&&view!=="vault"&&view!=="rcstatus"&&view!=="stock"&&view!=="uploads"&&setView("vault")}
        {view==="dashboard"&&<Dashboard custs={myC} role={role} onOpen={openD} onNav={nav} onNavF={st=>{setCustF(st);nav("customers");}} onSvcDone={id=>{upd(id,{serviceDone:true});notify("Service marked done ✓");}} onTeamTap={s=>{setFSM(s);nav("followups");}} onAddPayment={addPayment} onUpd={upd} notify={notify}/>}
        {view==="followups"&&<Followups items={due} onOpen={openD} onLog={logF} onCallLog={logCall} showSMFilter={role!=="salesman"} initSM={fSM}/>}
        {view==="customers"&&<CustList custs={myC} onOpen={openD} initF={custF} showSM={role!=="salesman"}/>}
        {view==="detail"&&sel&&<Detail cust={custs.find(c=>c.id===sel.id)||sel} role={role} onBack={goBack} onUpd={p=>upd(sel.id,p)} onLog={logF} onBill={()=>setBillOpen(true)} onBook={()=>setBookOpen(true)} notify={notify} initTab={dtab} clearInit={()=>setDtab(null)} onAddPayment={addPayment}/>}
        {view==="uploads"&&<div style={{padding:"0 16px 80px"}}><UploadsHub stockData={stockData} statusData={statusData} onStockUpload={saveStockData} onStatusUpload={saveStatusData} notify={notify}/></div>}
        {view==="stock"&&<div style={{padding:"0 16px 80px"}}><StockView stockData={stockData} billedChassis={billedChassis} role={role} userBranch={role==="salesman"?(SM_BRANCH[user]||BRANCHES[0]):role==="manager"?mBr:null} onUpload={saveStockData} notify={notify}/></div>}
        {view==="rcstatus"&&<div style={{padding:"0 16px 80px"}}><RCHSRPSearch statusData={statusData} role={role} onUpload={saveStatusData} notify={notify}/></div>}
        {view==="approvals"&&<Approvals custs={myPending} onApprove={approveBill} onOpen={openD} onEditCalc={c=>{setSel(c);setBillOpen(true);}} allC={myC} canApprove={role!=="salesman"}/>}
        {view==="revival"&&<Revival items={revivable} onRevive={ids=>{let si=0;const perDay={};setCusts(p=>p.map(c=>{
          if(!ids.includes(c.id))return c;
          const sm2=SM[si++%SM.length];
          perDay[sm2]=(perDay[sm2]||0)+1;
          const dayOffset=Math.floor((perDay[sm2]-1)/80);
          return{...c,reactivatedAt:td(),status:"Cold",stopped:false,attempts:0,alertDismissed:true,followupDate:aD(td(),dayOffset),salesman:sm2,remarks:(c.remarks||"")+"\n["+td()+"] REACTIVATED: cold pool — day "+(dayOffset+1)+" queue"};
        }));notify(ids.length+" reactivated — max 80 calls/day per executive, spread across days");}}/>}
        {view==="reports"&&<Reports custs={custs} onImportCust={rows=>{setCusts(p=>{const ex=new Set(p.map(c=>c.phone));return[...rows.filter(r=>!ex.has(r.phone)),...p];});}}/>}
        {view==="vault"&&<DocVault custs={custs} onImport={data=>{setCusts(data);notify("✅ Database imported: "+data.length+" customers");}}/>}
        {view==="alerts"&&(
          <div>
            <div style={{fontWeight:800,fontSize:19,color:"#1e293b",marginBottom:14}}>⚠️ Manager Alerts</div>
            {alerts.map((c,i)=>(
              <div key={i} style={{background:"rgba(239,68,68,0.07)",border:"1px solid rgba(239,68,68,0.3)",borderRadius:14,marginBottom:12,padding:"12px 14px"}}>
                <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
                  <div><div style={{fontWeight:700,fontSize:14,color:"#1e293b"}}>{c.name}</div><div style={{fontSize:12,color:"#64748b"}}>{c.phone} · {c.model}</div><div style={{fontSize:11,color:"#ef4444",marginTop:2}}>Stopped after 3 attempts · {c.salesman}</div></div>
                  <div style={{display:"flex",flexDirection:"column",gap:6}}>
                    <a href={"tel:"+c.phone} style={{background:"rgba(34,197,94,0.12)",border:"1px solid rgba(34,197,94,0.3)",borderRadius:8,padding:"6px 10px",display:"flex",alignItems:"center",gap:4,textDecoration:"none",color:"#22c55e",fontSize:11,fontWeight:700}}><PhIcon s={11}/>Call</a>
                    <a href={"https://wa.me/91"+c.phone} target="_blank" rel="noreferrer" style={{background:"rgba(37,211,102,0.08)",border:"1px solid rgba(37,211,102,0.3)",borderRadius:8,padding:"6px 10px",display:"flex",alignItems:"center",gap:4,textDecoration:"none",color:"#25D366",fontSize:11,fontWeight:700}}><WAIcon s={11}/>WA</a>
                  </div>
                </div>
                <button onClick={()=>upd(c.id,{alertDismissed:true})} style={{background:"transparent",border:"none",color:"#94a3b8",fontSize:12,cursor:"pointer",marginTop:8}}>Dismiss</button>
              </div>
            ))}
          </div>
        )}
      </div>

      {(view==="customers"||view==="dashboard")&&<button onClick={()=>setAddOpen(true)} style={{position:"fixed",bottom:78,right:20,width:58,height:56,borderRadius:28,background:"linear-gradient(135deg,#f97316,#ef4444)",border:"none",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",boxShadow:"0 6px 28px rgba(249,115,22,0.55)",zIndex:50,fontSize:28,color:"#fff",animation:"glow 3s ease infinite"}}>+</button>}
      {addOpen&&<AddModal onClose={()=>setAddOpen(false)} onSave={d=>{addC(d);setAddOpen(false);}} curUser={user} role={role} existing={custs}/>}
      {bookOpen&&sel&&<BookingModal cust={custs.find(c=>c.id===sel.id)||sel} onClose={()=>setBookOpen(false)} onSave={bk=>{const cu=custs.find(c=>c.id===sel.id)||sel;
        var brc="<!DOCTYPE html><html><head><title>Booking Receipt</title><style>*{margin:0;padding:0;box-sizing:border-box}body{font-family:Arial;font-size:13px;padding:20px;color:#111}.logo{font-size:24px;font-weight:900;letter-spacing:2px;text-align:center}.hdr{text-align:center;border-bottom:2px solid #000;padding-bottom:10px;margin-bottom:14px}.hdr p{font-size:11px;color:#444;margin-top:2px}h2{text-align:center;font-size:17px;margin:10px 0 14px}.row{display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid #eee}.v{font-weight:700}.total{display:flex;justify-content:space-between;padding:12px 0;border-top:2px solid #000;font-size:17px;font-weight:900;margin-top:8px}.sigs{display:flex;justify-content:space-between;margin-top:40px}.sigs div{text-align:center;font-size:11px}</style></head><body><div class=hdr><div class=logo>NKD BAJAJ</div><p>Authorised Main Dealer — Bajaj Auto Ltd.</p><p>Hirak Road, Near Kamal Katesaria School, Dhanbad</p><p>Ph: 7033099006 | info@nkdbajaj.com</p></div><h2>BOOKING RECEIPT</h2><div class=row><span>Date</span><span class=v>"+bk.date+"</span></div><div class=row><span>Customer</span><span class=v>"+cu.name+"</span></div><div class=row><span>Phone</span><span class=v>"+cu.phone+"</span></div><div class=row><span>Model</span><span class=v>"+(cu.model||"")+" ("+(cu.modelCode||"")+")</span></div><div class=row><span>Mode</span><span class=v>"+bk.mode+"</span></div>"+(bk.note?"<div class=row><span>Note</span><span class=v>"+bk.note+"</span></div>":"")+"<div class=total><span>BOOKING AMOUNT RECEIVED</span><span>"+fc(bk.amt)+"</span></div><p style='font-size:11px;color:#666;margin-top:8px'>Balance payable at delivery. Subject to realization of payment.</p><div class=sigs><div>____________________<br/>Customer Sign</div><div>____________________<br/>For NKD Bajaj</div></div></body></html>";
        const updCu={...cu,booking:{...bk,receiptHtml:brc},status:"Booked",photos:{...(cu.photos||{}),...(bk.proof?{booking_proof:bk.proof}:{})},remarks:(cu.remarks||"")+"\n["+td()+"] BOOKED: "+fc(bk.amt)+" ("+bk.mode+") booking date "+bk.date+(bk.note?" — "+bk.note:"")};
        upd(cu.id,updCu);
        try{const doc=makeBookingPdf(updCu);sharePdf(doc,"Booking_"+cu.name.replace(/ /g,"_")+"_"+td()+".pdf",cu.phone,"Please find your Booking Receipt from NKD Bajaj, Dhanbad.");savePdfToDrive(doc,"Booking_"+cu.name.replace(/ /g,"_")+"_"+td()+".pdf",cu.name,"Booking");}catch(e){}
        setBookOpen(false);setDtab("docs");notify("✅ Booking saved & receipt sent to customer!");}}/>}
      {billOpen&&sel&&<BillingModal cust={custs.find(c=>c.id===sel.id)||sel} onClose={()=>setBillOpen(false)} onSave={d=>{billC(custs.find(c=>c.id===sel.id)||sel,d);setBillOpen(false);}} onDraft={d=>{upd(sel.id,{billingDraft:d});setBillOpen(false);}} notify={notify} role={role} stockData={stockData} billedChassis={billedChassis}/>}
    </div>
  );
}

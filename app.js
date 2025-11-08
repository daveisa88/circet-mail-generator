// ===== Modèles de base =====
const modeles = {
  "Avancement de prod": {
    objet: "Avancement de prod {REGION} {DATE}",
    corps: "Bonjour,\n\nVoici l'avancement du jour pour {REGION} à {HEURE}.\n"
  },
  "Compte Rendu": {
    objet: "Compte rendu {REGION} {DATE}",
    corps: "Bonjour à tous,\nCompte rendu du jour :\n"
  },
  "Planning": {
    objet: "Planning {REGION} {DATE}",
    corps: "Bonsoir,\nVoici le planning prévu demain.\n"
  },
  "Synthèse du soir": {
    objet: "Synthèse {DATE} {REGION}",
    corps: "Bonjour,\nVoici la synthèse du jour.\n"
  }
};

// ===== Menus & mapping sauvegardé =====
let mappings = JSON.parse(localStorage.getItem("mappingsMenus") || "{}");

function chargerMenus() {
  const selType = document.getElementById("typeMail");
  const selReg  = document.getElementById("region");
  selType.innerHTML = "";
  [...new Set(Object.keys(modeles).concat(Object.keys(mappings).map(k=>k.split("_")[0])))].forEach(t=>{
    if(!t) return;
    const opt=document.createElement("option");
    opt.value=opt.textContent=t;
    selType.appendChild(opt);
  });
  // régions de base
  ["IDF","CVL","DOE","PRM","Corse","Lot 2 PDL","74/01","63/03","IDF LotE"].forEach(r=>{
    const o=document.createElement("option"); o.value=o.textContent=r; selReg.appendChild(o);
  });
}
chargerMenus();

// ===== Excel destinataires =====
let excelData = [], lastExcelFile = null;
function handleImport(target){ document.getElementById(target==="toEmails"?"fileTo":"fileCc").click(); }

function loadExcel(input, target){
  if(input.files && input.files[0]) lastExcelFile = input.files[0];
  if(!lastExcelFile) return alert("Sélectionne DestinataireMail.xlsx");
  const r = new FileReader();
  r.onload = (e)=>{
    const wb = XLSX.read(new Uint8Array(e.target.result), {type:"array"});
    excelData = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {header:1});
    localStorage.setItem("destinatairesExcelCache", JSON.stringify(excelData));
    remplirEmails(target);
  };
  r.readAsArrayBuffer(lastExcelFile);
}
const cache = localStorage.getItem("destinatairesExcelCache");
if(cache) excelData = JSON.parse(cache);

function remplirEmails(target){
  const key = (document.getElementById("typeMail").value + " " + document.getElementById("region").value).toLowerCase();
  for (let i=1;i<excelData.length;i++){
    if(((excelData[i][1]||"")+"").toLowerCase().trim()===key){
      document.getElementById(target).value = (target==="toEmails" ? (excelData[i][2]||"") : (excelData[i][3]||""));
      return;
    }
  }
}

// Auto-remplissage si l’utilisateur change le type/région
document.getElementById("typeMail").addEventListener("change", ()=>remplirEmails("toEmails"));
document.getElementById("region").addEventListener("change", ()=>remplirEmails("toEmails"));

// ===== Utilitaires =====
function normalizeEmails(str){
  return (str||"").replace(/\s+/g,"").replace(/[,;]/g,";").split(";").filter(Boolean);
}
function getWeekNumber(d){
  d=new Date(d); d.setHours(0,0,0,0); d.setDate(d.getDate()+3-(d.getDay()+6)%7);
  const w1=new Date(d.getFullYear(),0,4);
  return 1+Math.round(((d-w1)/86400000-3+(w1.getDay()+6)%7)/7);
}

// ===== Génération =====
function genererMail(){
  const type = document.getElementById("typeMail").value;
  const region = document.getElementById("region").value;
  const date = document.getElementById("date").value;
  const heure = document.getElementById("heure").value;
  if(!date) { alert("Sélectionne une date."); return; }

  const dateStr = new Date(date).toLocaleDateString('fr-FR');
  const semaine = getWeekNumber(date);

  let obj = (modeles[type]?.objet)||type;
  let body = (mappings[type+"_"+region] || modeles[type]?.corps || "").trim();

  const tokens = {"{REGION}":region,"{DATE}":dateStr,"{HEURE}":heure,"{SEMAINE}":semaine};
  for (const t in tokens){ obj=obj.replaceAll(t,tokens[t]); body=body.replaceAll(t,tokens[t]); }

  document.getElementById("objet").value = obj;
  document.getElementById("corps").value = body + "\n\n";
}

// ===== Ajout / Reset mappings =====
function ajouterMapping(){
  const t=document.getElementById("newType").value.trim();
  const r=document.getElementById("newRegion").value.trim();
  const b=document.getElementById("newBody").value.trim();
  if(!t || !b) return alert("Type + Corps requis");
  mappings[t+"_"+r] = b;
  localStorage.setItem("mappingsMenus", JSON.stringify(mappings));
  chargerMenus();
  alert("✅ Mapping ajouté");
}
function resetMappings(){
  if(confirm("Supprimer tous les mappings ajoutés ?")){
    localStorage.removeItem("mappingsMenus");
    mappings = {};
    chargerMenus();
  }
}

// ===== Mailto (signature Outlook conservée ; pas de PJ avec mailto) =====
function ouvrirOutlook(){
  const to = normalizeEmails(document.getElementById("toEmails").value).join("; ");
  const cc = normalizeEmails(document.getElementById("ccEmails").value).join("; ");
  const subject = document.getElementById("objet").value;
  const body = document.getElementById("corps").value + "\n\n";

  navigator.clipboard.writeText(body).finally(()=>{
    window.location.href = `mailto:${encodeURIComponent(to)}?cc=${encodeURIComponent(cc)}&subject=${encodeURIComponent(subject)}`;
    setTimeout(()=>document.execCommand("paste"), 800);
  });
}

// ===== Insertion via le complément Outlook =====
async function fileToBase64(input) {
  const f = input.files && input.files[0];
  if(!f) return null;
  return new Promise((resolve,reject)=>{
    const r=new FileReader();
    r.onload=()=>resolve((r.result||"").toString().split(",")[1]||"");
    r.onerror=reject;
    r.readAsDataURL(f);
  });
}

async function insertIntoOutlookFromAddin(){
  if(!window.Office || !Office.context || !Office.context.mailbox || !Office.context.mailbox.item){
    alert("Complément Outlook indisponible.");
    return;
  }
  const item = Office.context.mailbox.item; // mode COMPOSE

  const to = normalizeEmails(document.getElementById("toEmails").value);
  const cc = normalizeEmails(document.getElementById("ccEmails").value);
  const subject = document.getElementById("objet").value;
  const body = (document.getElementById("corps").value||"") + "\n\n";

  // Sujet / Body
  item.subject.setAsync(subject);
  item.body.setAsync(body.replace(/\n/g,"<br>"), { coercionType: Office.CoercionType.Html });

  // À / CC
  if (item.to && to.length) item.to.setAsync(to.map(a=>({emailAddress:a})));
  if (item.cc && cc.length) item.cc.setAsync(cc.map(a=>({emailAddress:a})));

  // PJ locale → base64 si API dispo
  try{
    const base64 = await fileToBase64(document.getElementById("attachmentInput"));
    if (base64 && item.addFileAttachmentFromBase64Async) {
      const filename = (document.getElementById("attachmentInput").files[0]||{}).name || "piece_jointe.bin";
      item.addFileAttachmentFromBase64Async(base64, filename, (res)=>{
        if (res.status !== Office.AsyncResultStatus.Succeeded) {
          console.warn("Attachment error:", res.error);
          alert("Pièce jointe : échec d’insertion ("+res.error.code+")");
        }
      });
    }
  }catch(e){
    console.warn(e);
  }

  alert("Données insérées dans le brouillon Outlook ✅");
}

// Expose global
window.handleImport=handleImport;
window.loadExcel=loadExcel;
window.genererMail=genererMail;
window.ouvrirOutlook=ouvrirOutlook;
window.ajouterMapping=ajouterMapping;
window.resetMappings=resetMappings;
window.insertIntoOutlookFromAddin=insertIntoOutlookFromAddin;

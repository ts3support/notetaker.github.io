/** Complete Code.gs: department forms, campaigns, Helpful Links hub data and email templates. */
function text_(v) { return v == null ? '' : String(v).trim(); }
function flag_(v, fallback) {
  if (typeof v === 'boolean') return v;
  const s = text_(v).toLowerCase();
  if (['true','yes','1'].includes(s)) return true;
  if (['false','no','0'].includes(s)) return false;
  return fallback === undefined ? false : fallback;
}
function list_(v) { return text_(v).split(/[,;\n]+/).map(text_).filter(Boolean); }
function department_(v) {
  const s = text_(v).toLowerCase().replace(/\s+/g,' ');
  if (/^(tech support|technical support)(?:$|[ -])/.test(s)) return {key:'tech',name:'Tech Support'};
  if (/^(retention)(?:$|[ -])/.test(s)) return {key:'retention',name:'Retention'};
  if (/^(member experience|mx)(?:$|[ -])/.test(s)) return {key:'mx',name:'Member Experience'};
  return null;
}
function headers_(row) { return row.map(v => text_(v).toLowerCase().replace(/\s+/g,' ')); }
function cell_(row,h,name) { const i=h.indexOf(name); return i<0 ? '' : row[i]; }
function requireHeaders_(h,names,tab) {
  const missing=names.filter(n=>!h.includes(n));
  if(missing.length) throw new Error(tab+': missing headers: '+missing.join(', '));
}
function documentId_(value) {
  const s=text_(value), match=s.match(/\/d\/([A-Za-z0-9_-]+)/)||s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  const id=match ? match[1] : s;
  if(!/^[A-Za-z0-9_-]+$/.test(id)) throw new Error('Invalid Google document ID or URL.');
  return id;
}
function doGet(e) {
  const p=e&&e.parameter||{}, page=text_(p.page).toLowerCase();
  let file='Login';
  if(['notes','main','app','index'].includes(page)) file='Notes';
  else if(page!=='login'&&(p.agent||p.extension||p.ext||p.user)) file='Notes';
  let t;
  try { t=HtmlService.createTemplateFromFile(file); }
  catch(err) { t=HtmlService.createTemplateFromFile(file.toLowerCase()); }
  t.appUrl=ScriptApp.getService().getUrl(); t.params=p;
  t.agent=p.agent||p.name||''; t.extension=p.extension||p.ext||'';
  t.department=p.department||p.dept||''; t.role=p.role||p.role_tier||'';
  return t.evaluate().setTitle('AAA Smart Home - Note Taker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function include(filename) {
  const n=text_(filename);
  for(const candidate of [n,n.charAt(0).toUpperCase()+n.slice(1),n.toLowerCase(),n.toUpperCase()]) {
    try { return HtmlService.createHtmlOutputFromFile(candidate).getContent(); } catch(err) {}
  }
  return '';
}
function getSettings() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings'), result={};
  if(sheet) sheet.getDataRange().getValues().slice(1).forEach(row=>{ if(text_(row[0])) result[text_(row[0])]=row[1]; });
  return result;
}
function getProjectExtensions() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Settings');
  // Original special-email access list stays at Settings!B3.
  return sheet ? list_(sheet.getRange('B3').getDisplayValue()) : [];
}
function isProjectExtension(ext) { return getProjectExtensions().includes(text_(ext)); }
function getLoginFields() {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  const sheet=ss.getSheetByName('login')||ss.getSheetByName('Login')||ss.getSheetByName('Index');
  if(!sheet) return [];
  return sheet.getDataRange().getValues().slice(1).filter(r=>text_(r[0])).map(r=>({
    id:text_(r[0]),internalId:text_(r[0]),label:text_(r[1])||text_(r[0]),type:text_(r[2])||'Text',
    options:list_(r[3]),visible:flag_(r[4]),required:flag_(r[5])
  }));
}
function getLoginConfig() { return getLoginFields(); }
function validateLogin(agentName,extension,department) {
  // Compatibility validation, not authentication of the supplied identity.
  const valid=!!text_(agentName)&&!!text_(extension)&&!!department_(department);
  return {success:valid,agentName:text_(agentName),extension:text_(extension),department:text_(department),
    message:valid?'':'Enter an agent name, extension, and recognized department.'};
}
function getActiveDepartmentCampaigns(department,extension) {
  const dept=department_(department); if(!dept) return [];
  const tab='Campaigns - '+dept.name;
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(tab);
  if(!sheet) return [];
  const data=sheet.getDataRange().getValues(); if(data.length<2) return [];
  const h=headers_(data[0]);
  requireHeaders_(h,['campaign id','display label','field type','options','active','target scope','target extensions'],tab);
  const result=[], seen=new Set(), ext=text_(extension);
  data.slice(1).forEach(row=>{
    if(!flag_(cell_(row,h,'active'))) return;
    const id=text_(cell_(row,h,'campaign id')); if(!id) return;
    const scope=text_(cell_(row,h,'target scope')).toLowerCase().replace(/[\s_]+/g,'-');
    const targets=list_(cell_(row,h,'target extensions'));
    const targeted=scope==='department-wide'||targets.some(v=>v.toUpperCase()==='ALL')||
      (scope==='extension-specific'&&ext!==''&&targets.includes(ext));
    if(!targeted) return;
    if(seen.has(id)) throw new Error(tab+': duplicate active campaign ID '+id);
    seen.add(id);
    const options=list_(cell_(row,h,'options'));
    result.push({id:id,campaignId:id,label:text_(cell_(row,h,'display label'))||id,
      type:text_(cell_(row,h,'field type'))||'Radio',options:options.length?options:['Yes','No'],
      required:flag_(cell_(row,h,'required')),targetScope:scope,notes:text_(cell_(row,h,'notes'))});
  });
  return result;
}
function getFormSchema(department,extension) {
  const ss=SpreadsheetApp.getActiveSpreadsheet(), ext=text_(extension), agentDept=department_(department);
  const campaigns={tech:getActiveDepartmentCampaigns('Tech Support',ext),
    retention:getActiveDepartmentCampaigns('Retention',ext),mx:getActiveDepartmentCampaigns('Member Experience',ext)};
  const schema={campaigns:campaigns,activeCampaigns:agentDept?campaigns[agentDept.key]:[],
    authorizedExtensions:getProjectExtensions(),warnings:[]};
  const specs=[['login','login','tech'],['index','login','tech'],['inbound','Inbound','tech'],
    ['outbound','Outbound','tech'],['retention','Retention','retention'],['retention_outbound','Retention Outbound','retention'],
    ['residential','Residential','tech'],['dealer','Dealer','tech'],['addendum','Addendum',null],['parts','Parts','tech'],
    ['mx','MX','mx'],['mx_outbound','MX Outbound','mx']];
  const cache={};
  specs.forEach(spec=>{
    const key=spec[0],tab=spec[1]; schema[key]={}; schema[key+'_order']=[];
    if(!Object.prototype.hasOwnProperty.call(cache,tab)) {
      const sheet=ss.getSheetByName(tab)||(tab==='login'?(ss.getSheetByName('Login')||ss.getSheetByName('Index')):null);
      cache[tab]=sheet?sheet.getDataRange().getValues():[];
    }
    const data=cache[tab]; if(data.length<2) return;
    const h=headers_(data[0]);
    requireHeaders_(h,['internal id','display label','field type','options','visible','required'],tab);
    const owner=spec[2]||(agentDept?agentDept.key:null), available=owner?campaigns[owner]:[];
    data.slice(1).forEach(row=>{
      const sourceId=text_(cell_(row,h,'internal id')); if(!sourceId) return;
      const options=text_(cell_(row,h,'options'));
      const dynamic=sourceId==='campaign_toggle'||/dynamic:campaigns/i.test(options);
      const id=dynamic?'campaign_toggle':sourceId;
      const field={id:id,internalId:id,label:text_(cell_(row,h,'display label'))||sourceId,
        type:text_(cell_(row,h,'field type'))||'Text',options:options,visible:flag_(cell_(row,h,'visible')),
        required:flag_(cell_(row,h,'required')),dependsOn:text_(cell_(row,h,'depends on')),
        copyIfBlank:flag_(cell_(row,h,'copy if blank'))};
      if(dynamic) {
        field.isCampaign=true; field.placement='top_controls'; field.sourceId=sourceId; field.campaignId='';
        if(available.length) {
          const c=available[0];field.campaignId=c.campaignId;field.label=c.label;field.type=c.type;
          field.options=c.options.join(',');field.required=c.required;
          if(available.length>1) schema.warnings.push(tab+': multiple campaigns matched; using first row: '+c.campaignId);
        } else {field.visible=false;field.required=false;field.options='';}
      }
      if(schema[key][id]) throw new Error(tab+': duplicate resolved Internal ID '+id);
      schema[key][id]=field;schema[key+'_order'].push(field);
    });
  });
  return schema;
}

// HELPFUL LINKS HUB DATA: exact primary Tech Support tab is Links-Tech Support.
function getHelpfulLinks(department) {
  const dept=department_(department);
  if(!dept) throw new Error('Helpful Links: supply a recognized department.');
  const ss=SpreadsheetApp.getActiveSpreadsheet(), candidates=['Links-'+dept.name,'Links - '+dept.name];
  if(dept.key==='mx') candidates.push('Links-MX','Links - MX');
  let sheet=null;
  for(const tab of candidates) {sheet=ss.getSheetByName(tab);if(sheet) break;}
  if(!sheet) throw new Error('Helpful Links: worksheet not found. Expected '+candidates[0]+'.');
  const data=sheet.getDataRange().getDisplayValues();
  if(!data.length||data.every(row=>row.every(v=>!text_(v)))) return [];
  const normalize=v=>text_(v).toLowerCase().replace(/[_-]+/g,' ').replace(/\s+/g,' ');
  const h=data[0].map(normalize);
  const find=names=>{for(const n of names){const i=h.indexOf(normalize(n));if(i>=0)return i;}return -1;};
  const c={id:find(['Link ID','Resource ID']),title:find(['Link Title','Title','Name']),url:find(['URL','Link URL']),
    category:find(['Category']),description:find(['Description']),instructions:find(['Agent Instructions','Instructions']),
    popup:find(['Popup Mode']),sort:find(['Sort Order']),active:find(['Active'])};
  const missing=[];
  if(c.title<0)missing.push('Link Title');if(c.url<0)missing.push('URL');if(c.active<0)missing.push('Active');
  if(missing.length)throw new Error(sheet.getName()+': missing required headers: '+missing.join(', '));
  const value=(row,col)=>col<0?'':text_(row[col]), links=[],seen=new Set();
  for(let r=1;r<data.length;r++) {
    const row=data[r];if(!flag_(value(row,c.active)))continue;
    const title=value(row,c.title),url=value(row,c.url);
    if(!title||!/^https?:\/\/[^\s/]+(?:[/?#][^\s]*)?$/i.test(url)) {
      console.warn(sheet.getName()+': skipped active row '+(r+1)+'; invalid title or HTTP(S) URL.');continue;
    }
    const id=value(row,c.id)||dept.key+'_row_'+(r+1);
    if(seen.has(id))throw new Error(sheet.getName()+': duplicate Link ID '+id);seen.add(id);
    const rawMode=value(row,c.popup).toLowerCase(),mode=['popup','modal','chatbot'].includes(rawMode)?'popup':'new_tab';
    const rawSort=value(row,c.sort),numeric=rawSort===''?NaN:Number(rawSort),instructions=value(row,c.instructions);
    links.push({id:id,name:title,title:title,url:url,category:value(row,c.category)||'General',
      description:value(row,c.description),instructions:instructions,agentInstructions:instructions,department:dept.name,
      popupMode:mode,popup:mode==='popup'?'chatbot':null,sortOrder:Number.isFinite(numeric)?numeric:r+1,sourceRow:r+1,active:true});
  }
  links.sort((a,b)=>a.sortOrder-b.sortOrder||a.sourceRow-b.sourceRow);
  return links;
}
function getHelpfulLinksFromDoc(department) {
  if(text_(department))return getHelpfulLinks(department);
  const id=text_(getSettings()['Doc ID - Helpful Links']);if(!id)return [];
  const links=[];
  DocumentApp.openById(documentId_(id)).getBody().getParagraphs().forEach(p=>{
    const name=text_(p.getText());if(!name)return;
    const t=p.editAsText();
    for(let i=0;i<t.getText().length;i++) {
      const url=t.getLinkUrl(i);
      if(url&&/^https?:\/\//i.test(url)){links.push({name:name,title:name,url:url,category:'General',
        popup:/chatbot/i.test(name+' '+url)?'chatbot':null});break;}
    }
  });
  return links;
}
function getEmailTemplatesHtml(department) {
  const settings=getSettings(),dept=department_(department);
  if(text_(department)&&!dept)throw new Error('Unrecognized email-template department.');
  const specific=dept?text_(settings['Doc ID - Standard Emails - '+dept.name]):'';
  const configured=specific||text_(settings['Doc ID - Standard Emails']);
  if(!configured)throw new Error('No standard email document configured for '+(dept?dept.name:'the shared email library')+'.');
  const url='https://docs.google.com/document/d/'+encodeURIComponent(documentId_(configured))+'/export?format=html';
  const response=UrlFetchApp.fetch(url,{method:'get',headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()},muteHttpExceptions:true});
  const status=response.getResponseCode();
  if(status<200||status>=300)throw new Error('Email export failed (HTTP '+status+'). Check the document ID and permissions.');
  return response.getContentText();
}
function getSpecialEmailTemplates(extension) {
  if(!isProjectExtension(extension))return [];
  return readSpecialEmails_();
}
function getSpecialEmails(extension) {return getSpecialEmailTemplates(extension);}
function readSpecialEmails_() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Special Emails');if(!sheet)return [];
  return sheet.getDataRange().getValues().slice(1).filter(r=>text_(r[0])).map(r=>({
    id:text_(r[0]),brand:text_(r[1])||'AAA Smart Home',subject:text_(r[2]),body:text_(r[3])}));
}
function checkZipCode(zipInput) {
  const zip=text_(zipInput);if(!zip)return {found:false,message:'Please enter a valid zip code.'};
  const settings=getSettings(),id=settings['Sheet ID - Zip Codes']||'1O7L6NVy7N4vS47IMqlGZmExdOD13l8e8kmAuXw2AAKw';
  try {
    const ss=SpreadsheetApp.openById(id),pro=ss.getSheetByName('Official Pro Zips');
    if(pro&&pro.getDataRange().getDisplayValues().some(r=>r.some(v=>v.trim()===zip)))
      return {found:true,status:'PRO',message:'✅ In Footprint (Pro Territory)'};
    const ds=ss.getSheetByName('Official DS (Upgrade & Service)')||ss.getSheetByName('Official DS')||ss.getSheetByName('Dealer Service');
    if(ds) {
      const data=ds.getDataRange().getDisplayValues();
      if(data.length>2)for(let col=0;col<data[0].length;col++) {
        const noTech=/NO TECH/i.test(data[0][col]+' '+data[1][col]);
        for(let row=2;row<data.length;row++) {
          if(data[row][col].trim()!==zip)continue;
          if(noTech)return {found:false,status:'NO_TECH',message:'❌ No Tech Available (Dealer No Tech Territory)'};
          const territory=data[0][col]||data[1][col]||'Dealer Service';
          return {found:true,status:'DEALER',message:'✅ Dealer Service ('+territory.trim()+')'};
        }
      }
    }
    return {found:false,status:'NOT_FOUND',message:'❌ No Tech Available'};
  }catch(err){console.error(err);return {found:false,status:'ERROR',message:'Error searching zip footprint: '+err.message};}
}
function submitFeedback(payload) {
  try {
    const p=payload||{},sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Feedback');
    if(!sheet)return {success:false,message:'Feedback sheet could not be found.'};
    const comments=text_(p.comments||p.details||p.message);if(!comments)return {success:false,message:'Please enter comments or details.'};
    const safe=v=>{const s=text_(v);return /^[=+@-]/.test(s)?"'"+s:s;};
    sheet.appendRow([new Date(),safe(p.name||'Anonymous'),safe(p.extension||p.phone||'-'),safe(p.dept||p.department||'-'),
      safe(p.category||'General'),safe(p.section||p.form||p.applicableForm||'General'),safe(comments)]);
    return {success:true,message:'Feedback submitted successfully. Thank you!'};
  }catch(err){return {success:false,message:'Error submitting feedback: '+err.message};}
}
function getChangeLog() {
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Change Log');if(!sheet)return [];
  return sheet.getDataRange().getValues().slice(1).filter(r=>flag_(r[5],true)).map(r=>({
    date:r[0] instanceof Date?Utilities.formatDate(r[0],Session.getScriptTimeZone(),'MM/dd/yyyy'):text_(r[0]),
    version:text_(r[1])||'v1.x',form:text_(r[2])||'General',description:text_(r[3]),updatedBy:text_(r[4])||'Admin'})).reverse();
}
function validateAdminPasscode(enteredPasscode) {
  const correct=text_(getSettings()['Admin Password']);return correct!==''&&text_(enteredPasscode)===correct;
}

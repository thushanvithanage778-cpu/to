require('dotenv').config();
const path = require('path');
const fs = require('fs');
const express = require('express');
const bcrypt = require('bcryptjs');
const jwt = require('jsonwebtoken');
const Database = require('better-sqlite3');
const { ConfidentialClientApplication } = require('@azure/msal-node');

const app = express();
const PORT = Number(process.env.PORT || 3000);
const JWT_SECRET = process.env.JWT_SECRET || 'CHANGE_ME_IN_PRODUCTION';

if (JWT_SECRET === 'CHANGE_ME_IN_PRODUCTION') {
  console.warn('WARNING: Set JWT_SECRET in .env before production use.');
}

app.use(express.json({ limit: '2mb' }));
app.use(express.urlencoded({ extended: true }));

const DATA_DIR = path.join(__dirname, 'data');
fs.mkdirSync(DATA_DIR, { recursive: true });
const db = new Database(path.join(DATA_DIR, 'insee.sqlite'));
db.pragma('journal_mode = WAL');

db.exec(`
CREATE TABLE IF NOT EXISTS users (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  username TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role TEXT NOT NULL,
  department TEXT DEFAULT '',
  area TEXT DEFAULT '',
  active INTEGER NOT NULL DEFAULT 1,
  created_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS mappings (
  area TEXT PRIMARY KEY,
  name TEXT NOT NULL,
  email TEXT NOT NULL,
  esc_manager TEXT DEFAULT '',
  esc_email TEXT DEFAULT '',
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS observations (
  id TEXT PRIMARY KEY,
  date TEXT NOT NULL,
  observer TEXT,
  task TEXT,
  wi TEXT,
  area TEXT,
  manager TEXT,
  manager_email TEXT,
  action_required TEXT,
  risk TEXT,
  status TEXT DEFAULT 'Submitted',
  data_json TEXT NOT NULL,
  created_at TEXT NOT NULL,
  updated_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS actions (
  id TEXT PRIMARY KEY,
  obs_id TEXT NOT NULL UNIQUE,
  area TEXT,
  manager TEXT,
  email TEXT,
  finding TEXT,
  action TEXT,
  risk TEXT,
  target TEXT,
  status TEXT DEFAULT 'Open',
  completion TEXT DEFAULT '',
  updated_at TEXT NOT NULL,
  FOREIGN KEY(obs_id) REFERENCES observations(id) ON DELETE CASCADE
);
`);

const defaultManagers = {
  Production:{name:'Production Area Manager',email:'production.manager@insee.example',escManager:'',escEmail:''},
  'Packing Plant':{name:'Packing Plant Area Manager',email:'packing.manager@insee.example',escManager:'',escEmail:''},
  'Mechanical Workshop':{name:'Mechanical Area Manager',email:'mechanical.manager@insee.example',escManager:'',escEmail:''},
  Electrical:{name:'Electrical Area Manager',email:'electrical.manager@insee.example',escManager:'',escEmail:''},
  Logistics:{name:'Logistics Area Manager',email:'logistics.manager@insee.example',escManager:'',escEmail:''},
  Quarry:{name:'Quarry Area Manager',email:'quarry.manager@insee.example',escManager:'',escEmail:''},
  Administration:{name:'Administration Area Manager',email:'admin.manager@insee.example',escManager:'',escEmail:''}
};

function seed() {
  const userCount = db.prepare('SELECT COUNT(*) AS n FROM users').get().n;
  if (!userCount) {
    const insert = db.prepare(`INSERT INTO users (name,username,password_hash,role,department,area,active,created_at) VALUES (?,?,?,?,?,?,1,?)`);
    const users = [
      ['System Administrator','admin','admin123','Admin','Health & Safety',''],
      ['Safety Manager','safety','safety123','Safety Manager','Health & Safety',''],
      ['Area Manager','manager','manager123','Area Manager','Production','Production'],
      ['Safety Observer','observer','observer123','Observer','Health & Safety','']
    ];
    const tx = db.transaction(() => users.forEach(u => insert.run(u[0],u[1],bcrypt.hashSync(u[2],10),u[3],u[4],u[5],new Date().toISOString())));
    tx();
    console.log('Seeded demo users. Change passwords before production use.');
  }
  const mappingCount = db.prepare('SELECT COUNT(*) AS n FROM mappings').get().n;
  if (!mappingCount) {
    const insert = db.prepare(`INSERT INTO mappings(area,name,email,esc_manager,esc_email,updated_at) VALUES (?,?,?,?,?,?)`);
    const tx = db.transaction(() => Object.entries(defaultManagers).forEach(([area,m]) => insert.run(area,m.name,m.email,m.escManager,m.escEmail,new Date().toISOString())));
    tx();
  }
}
seed();

function safeUser(row) {
  return {id:row.id,name:row.name,username:row.username,role:row.role,department:row.department,area:row.area||'',active:!!row.active};
}
function tokenFor(user) {
  return jwt.sign({id:user.id,username:user.username,role:user.role}, JWT_SECRET, {expiresIn:'8h'});
}
function auth(req,res,next) {
  const header = req.headers.authorization || '';
  const token = header.startsWith('Bearer ') ? header.slice(7) : '';
  if (!token) return res.status(401).json({error:'Authentication required.'});
  try {
    const payload = jwt.verify(token, JWT_SECRET);
    const row = db.prepare('SELECT * FROM users WHERE id=? AND active=1').get(payload.id);
    if (!row) return res.status(401).json({error:'User account is inactive or unavailable.'});
    req.user = safeUser(row);
    next();
  } catch (_) { return res.status(401).json({error:'Invalid or expired session.'}); }
}
function requireAdmin(req,res,next){ if(req.user.role!=='Admin') return res.status(403).json({error:'Admin access required.'}); next(); }
function requireActionEditor(req,res,next){ if(!['Admin','Safety Manager','Area Manager'].includes(req.user.role)) return res.status(403).json({error:'You do not have permission to edit corrective actions.'}); next(); }
function graphConfigured(){
  return !!(process.env.M365_TENANT_ID && process.env.M365_CLIENT_ID && process.env.M365_CLIENT_SECRET && process.env.M365_SENDER_EMAIL);
}
let msalClient=null;
function getMsalClient(){
  if(!graphConfigured()) return null;
  if(!msalClient){
    msalClient=new ConfidentialClientApplication({
      auth:{
        clientId:process.env.M365_CLIENT_ID,
        authority:`https://login.microsoftonline.com/${process.env.M365_TENANT_ID}`,
        clientSecret:process.env.M365_CLIENT_SECRET
      }
    });
  }
  return msalClient;
}
async function getGraphToken(){
  const client=getMsalClient();
  if(!client) return null;
  const result=await client.acquireTokenByClientCredential({scopes:['https://graph.microsoft.com/.default']});
  if(!result || !result.accessToken) throw new Error('Microsoft Graph access token was not returned.');
  return result.accessToken;
}
async function sendMicrosoft365Mail({to,cc,subject,text}){
  if(!graphConfigured()) return {sent:false,reason:'Microsoft 365 Graph email is not configured. Observation was saved successfully.'};
  const token=await getGraphToken();
  const recipients=(value)=>String(value||'').split(',').map(x=>x.trim()).filter(Boolean).map(address=>({emailAddress:{address}}));
  const payload={
    message:{
      subject,
      body:{contentType:'Text',content:text},
      toRecipients:recipients(to),
      ...(cc?{ccRecipients:recipients(cc)}:{}),
      from:{emailAddress:{address:process.env.M365_SENDER_EMAIL}}
    },
    saveToSentItems:true
  };
  const response=await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(process.env.M365_SENDER_EMAIL)}/sendMail`,{
    method:'POST',
    headers:{Authorization:`Bearer ${token}`,'Content-Type':'application/json'},
    body:JSON.stringify(payload)
  });
  if(!response.ok){
    const detail=await response.text();
    throw new Error(`Microsoft Graph sendMail failed (${response.status}): ${detail.slice(0,500)}`);
  }
  return {sent:true,provider:'Microsoft Graph'};
}
async function sendObservationEmail(rec, mapping, action) {
  const subject = `INSEE Task Observation – ${action ? 'Corrective Action Required – ' : ''}${rec.area} – ${rec.id}`;
  const to = mapping.email;
  const cc = (action && mapping.escEmail) ? mapping.escEmail : undefined;
  const body = [
    `Dear ${mapping.name},`, '',
    'A new task observation has been recorded in the INSEE Digital Task Observation System.', '',
    `Observation No.: ${rec.id}`,
    `Date: ${new Date(rec.date).toLocaleString()}`,
    `Observer: ${rec.observer || '-'}`,
    `Work Area: ${rec.area}`,
    `Task / WI: ${rec.task || '-'} / ${rec.wi || '-'}`,
    `Risk Level: ${rec.risk || '-'}`,
    `Corrective Action Required: ${rec.actionRequired || 'No'}`,
    '',
    `Observation: ${rec.unsafeDesc || 'No unsafe act/condition description provided.'}`,
    '',
    action ? `Required Corrective Action: ${rec.actionText || 'Review and correct the identified issue.'}` : '',
    action ? `Responsible Person: ${rec.responsible || '-'}` : '',
    action ? `Target Completion Date: ${rec.targetDate || '-'}` : '',
    '',
    'Please log in to the INSEE Digital Task Observation System to review and update the record.', '',
    'Regards,',
    process.env.M365_SENDER_NAME || 'INSEE Digital Task Observation System'
  ].filter(Boolean).join('\n');
  try {
    return await sendMicrosoft365Mail({to,cc,subject,text:body});
  } catch(e){
    console.error('Microsoft 365 email error:',e.message);
    return {sent:false,reason:e.message};
  }
}

app.post('/api/login', (req,res) => {
  const {username,password} = req.body || {};
  const row = db.prepare('SELECT * FROM users WHERE username=? AND active=1').get(String(username||'').trim());
  if (!row || !bcrypt.compareSync(String(password||''), row.password_hash)) return res.status(401).json({error:'Invalid username or password.'});
  const user = safeUser(row);
  res.json({token:tokenFor(user),user});
});

app.get('/api/bootstrap', auth, (req,res) => {
  const mappings = db.prepare('SELECT area,name,email,esc_manager AS escManager,esc_email AS escEmail FROM mappings ORDER BY area').all();
  const observations = db.prepare('SELECT * FROM observations ORDER BY datetime(date) DESC').all().map(r=>({...JSON.parse(r.data_json),id:r.id,date:r.date,status:r.status}));
  const actions = db.prepare('SELECT id,obs_id AS obsId,area,manager,email,finding,action,risk,target,status,completion FROM actions ORDER BY datetime(updated_at) DESC').all();
  const users = req.user.role==='Admin' ? db.prepare('SELECT * FROM users ORDER BY name').all().map(safeUser) : [];
  const next = nextObservationNumber();
  res.json({user:req.user,mappings,observations,actions,users,nextObservationNo:next});
});
function nextObservationNumber(){
  const year = new Date().getFullYear();
  const row = db.prepare(`SELECT COUNT(*) AS n FROM observations WHERE id LIKE ?`).get(`TO-${year}-%`);
  return `TO-${year}-${String(Number(row.n)+1).padStart(5,'0')}`;
}
app.get('/api/next-observation-number', auth, (req,res)=>res.json({id:nextObservationNumber()}));

app.get('/api/mappings', auth, (req,res)=>res.json(db.prepare('SELECT area,name,email,esc_manager AS escManager,esc_email AS escEmail FROM mappings ORDER BY area').all()));
app.post('/api/mappings', auth, requireAdmin, (req,res)=>{
  const {area,name,email,escManager='',escEmail=''}=req.body||{};
  if(!area||!name||!email) return res.status(400).json({error:'Work Area, Area Manager Name and Manager Email are required.'});
  try {
    db.prepare(`INSERT INTO mappings(area,name,email,esc_manager,esc_email,updated_at) VALUES (?,?,?,?,?,?)`).run(String(area).trim(),String(name).trim(),String(email).trim(),String(escManager||'').trim(),String(escEmail||'').trim(),new Date().toISOString());
    res.status(201).json({ok:true});
  } catch(e){res.status(409).json({error:'A mapping for this Work Area already exists.'});}
});
app.put('/api/mappings/:area', auth, requireAdmin, (req,res)=>{
  const oldArea=req.params.area, {area,name,email,escManager='',escEmail=''}=req.body||{};
  if(!area||!name||!email) return res.status(400).json({error:'Work Area, Area Manager Name and Manager Email are required.'});
  try {
    const tx=db.transaction(()=>{db.prepare('DELETE FROM mappings WHERE area=?').run(oldArea);db.prepare(`INSERT INTO mappings(area,name,email,esc_manager,esc_email,updated_at) VALUES (?,?,?,?,?,?)`).run(String(area).trim(),String(name).trim(),String(email).trim(),String(escManager||'').trim(),String(escEmail||'').trim(),new Date().toISOString());});
    tx(); res.json({ok:true});
  } catch(e){res.status(409).json({error:'A mapping for this Work Area already exists.'});}
});
app.delete('/api/mappings/:area', auth, requireAdmin, (req,res)=>{db.prepare('DELETE FROM mappings WHERE area=?').run(req.params.area);res.json({ok:true});});

function canEditObservationServer(user,o){
  if(['Admin','Safety Manager'].includes(user.role)) return true;
  if(user.role==='Observer') return o.observer===user.name;
  if(user.role==='Area Manager') return o.area===user.area;
  return false;
}

app.post('/api/observations', auth, async (req,res)=>{
  const rec=req.body||{};
  if(!rec.id||!rec.task||!rec.wi||!rec.area) return res.status(400).json({error:'Observation No., Task/WI, WI Number and Work Area are required.'});
  const existingRow=db.prepare('SELECT * FROM observations WHERE id=?').get(rec.id);
  if(existingRow){
    const existing=JSON.parse(existingRow.data_json);
    if(!canEditObservationServer(req.user,existing)) return res.status(403).json({error:'You do not have edit permission for this observation.'});
  }
  const mapping=db.prepare('SELECT area,name,email,esc_manager AS escManager,esc_email AS escEmail FROM mappings WHERE area=?').get(rec.area);
  if(!mapping) return res.status(400).json({error:'No Area Manager Mapping exists for the selected work area.'});
  rec.manager=mapping.name; rec.managerEmail=mapping.email; rec.date=rec.date||new Date().toISOString();
  const now=new Date().toISOString();
  const exists=existingRow;
  const tx=db.transaction(()=>{
    if(exists) db.prepare(`UPDATE observations SET date=?,observer=?,task=?,wi=?,area=?,manager=?,manager_email=?,action_required=?,risk=?,status=?,data_json=?,updated_at=? WHERE id=?`).run(rec.date,rec.observer||'',rec.task,rec.wi,rec.area,rec.manager,rec.managerEmail,rec.actionRequired||'No',rec.risk||'Low','Submitted',JSON.stringify(rec),now,rec.id);
    else db.prepare(`INSERT INTO observations(id,date,observer,task,wi,area,manager,manager_email,action_required,risk,status,data_json,created_at,updated_at) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?)`).run(rec.id,rec.date,rec.observer||'',rec.task,rec.wi,rec.area,rec.manager,rec.managerEmail,rec.actionRequired||'No',rec.risk||'Low','Submitted',JSON.stringify(rec),now,now);
    db.prepare('DELETE FROM actions WHERE obs_id=?').run(rec.id);
    if(rec.actionRequired && rec.actionRequired!=='No') {
      db.prepare(`INSERT INTO actions(id,obs_id,area,manager,email,finding,action,risk,target,status,completion,updated_at) VALUES (?,?,?,?,?,?,?,?,?,'Open','',?)`).run('CA-'+String(Date.now()).slice(-9),rec.id,rec.area,rec.manager,rec.managerEmail,rec.unsafeDesc||'Corrective action identified during task observation.',rec.actionText||'Review and correct the identified issue.',rec.risk||'Low',rec.targetDate||'',now);
    }
  });
  try {
    tx();
    let emailResult={sent:false,reason:'No email sent.'};
    if(rec.actionRequired && rec.actionRequired!=='No') emailResult=await sendObservationEmail(rec,mapping,true);
    res.status(exists?200:201).json({ok:true,observation:rec,email:emailResult});
  } catch(e){console.error(e);res.status(500).json({error:'Observation could not be saved.'});}
});

app.patch('/api/actions/:id', auth, requireActionEditor, (req,res)=>{
  const a=db.prepare('SELECT * FROM actions WHERE id=?').get(req.params.id);
  if(!a) return res.status(404).json({error:'Corrective action not found.'});
  const status=String(req.body.status||a.status), completion=req.body.completion===undefined?a.completion:String(req.body.completion);
  db.prepare('UPDATE actions SET status=?,completion=?,updated_at=? WHERE id=?').run(status,completion,new Date().toISOString(),a.id);
  res.json({ok:true});
});

app.get('/api/observations/:id', auth, (req,res)=>{
  const r=db.prepare('SELECT * FROM observations WHERE id=?').get(req.params.id);
  if(!r) return res.status(404).json({error:'Observation not found.'});
  const o={...JSON.parse(r.data_json),id:r.id,date:r.date,status:r.status};
  const can=['Admin','Safety Manager'].includes(req.user.role) || (req.user.role==='Observer'&&o.observer===req.user.name) || (req.user.role==='Area Manager'&&o.area===req.user.area);
  if(!can) return res.status(403).json({error:'You do not have edit permission for this observation.'});
  res.json(o);
});

app.post('/api/users', auth, requireAdmin, async (req,res)=>{
  const {name,username,password,role,department='',area=''}=req.body||{};
  if(!name||!username||!password||!role) return res.status(400).json({error:'Name, username, password and role are required.'});
  try {const info=db.prepare(`INSERT INTO users(name,username,password_hash,role,department,area,active,created_at) VALUES (?,?,?,?,?,?,1,?)`).run(name.trim(),username.trim(),await bcrypt.hash(password,10),role,department,area,new Date().toISOString());res.status(201).json({id:info.lastInsertRowid});}
  catch(e){res.status(409).json({error:'Username already exists.'});}
});
app.put('/api/users/:id', auth, requireAdmin, async (req,res)=>{
  const old=db.prepare('SELECT * FROM users WHERE id=?').get(req.params.id); if(!old)return res.status(404).json({error:'User not found.'});
  const {name,username,password,role,department='',area=''}=req.body||{};
  if(!name||!username||!role)return res.status(400).json({error:'Name, username and role are required.'});
  try {const hash=password?await bcrypt.hash(password,10):old.password_hash;db.prepare(`UPDATE users SET name=?,username=?,password_hash=?,role=?,department=?,area=? WHERE id=?`).run(name.trim(),username.trim(),hash,role,department,area,old.id);res.json({ok:true});}
  catch(e){res.status(409).json({error:'Username already exists.'});}
});
app.patch('/api/users/:id/status', auth, requireAdmin, (req,res)=>{const u=db.prepare('SELECT * FROM users WHERE id=?').get(req.params.id);if(!u)return res.status(404).json({error:'User not found.'});if(u.username==='admin')return res.status(400).json({error:'The primary admin account cannot be deactivated.'});db.prepare('UPDATE users SET active=? WHERE id=?').run(u.active?0:1,u.id);res.json({ok:true});});

app.get('/api/health', (req,res)=>res.json({ok:true,service:'INSEE Digital Task Observation API'}));

app.use(express.static(__dirname));
app.use((req,res,next) => {
  if (req.method === 'GET' && req.path.startsWith('/api/')) return res.status(404).json({error:'API endpoint not found.'});
  if (req.method === 'GET' && !req.path.includes('.')) return res.sendFile(path.join(__dirname,'index.html'));
  next();
});

app.listen(PORT,()=>console.log(`INSEE Digital Task Observation System running at http://localhost:${PORT}`));

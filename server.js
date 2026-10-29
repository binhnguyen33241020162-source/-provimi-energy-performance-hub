const express=require('express');
const cors=require('cors');
const {Pool}=require('pg');
const app=express();
const PORT=process.env.PORT||10000;
const ALLOWED_ORIGIN=process.env.APP_ORIGIN||'https://provimi-energy-performance-hub.onrender.com';
const DATABASE_URL=process.env.DATABASE_URL||'';
app.use(cors({origin:(origin,cb)=>{if(!origin||origin===ALLOWED_ORIGIN)return cb(null,true);cb(new Error('Origin not allowed'));},methods:['GET','PUT','OPTIONS'],allowedHeaders:['Content-Type']}));
app.use(express.json({limit:'8mb'}));
let pool=DATABASE_URL?new Pool({connectionString:DATABASE_URL,ssl:{rejectUnauthorized:false}}):null;
async function ensureSchema(){if(!pool)return;await pool.query('CREATE TABLE IF NOT EXISTS app_state (id INTEGER PRIMARY KEY, payload JSONB NOT NULL, updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW())');}
app.get('/health',async(req,res)=>{if(!pool)return res.status(503).json({ok:false,cloud:false,reason:'DATABASE_URL not configured'});try{await ensureSchema();await pool.query('SELECT 1');res.json({ok:true,cloud:true});}catch(e){res.status(503).json({ok:false,cloud:false,reason:e.message});}});
app.get('/api/state',async(req,res)=>{if(!pool)return res.status(503).json({ok:false,reason:'DATABASE_URL not configured'});try{await ensureSchema();const r=await pool.query('SELECT payload, updated_at FROM app_state WHERE id=1');if(!r.rows.length)return res.json({ok:true,state:null,updated_at:null});res.json({ok:true,state:r.rows[0].payload,updated_at:r.rows[0].updated_at});}catch(e){res.status(500).json({ok:false,reason:e.message});}});
app.put('/api/state',async(req,res)=>{if(!pool)return res.status(503).json({ok:false,reason:'DATABASE_URL not configured'});const state=req.body&&req.body.state;if(!state||typeof state!=='object')return res.status(400).json({ok:false,reason:'Missing state object'});try{await ensureSchema();const r=await pool.query('INSERT INTO app_state(id,payload,updated_at) VALUES(1,$1::jsonb,NOW()) ON CONFLICT(id) DO UPDATE SET payload=EXCLUDED.payload, updated_at=NOW() RETURNING updated_at',[JSON.stringify(state)]);res.json({ok:true,updated_at:r.rows[0].updated_at});}catch(e){res.status(500).json({ok:false,reason:e.message});}});
app.listen(PORT,()=>console.log('Provimi cloud API listening on '+PORT));

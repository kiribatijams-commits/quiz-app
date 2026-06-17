import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

const DEFAULT_QUIZ = [
  { id: 1, question: "微信的英文名称是什么？", options: ["WeChat", "WePay", "WeLink", "WeTalk"], answer: 0, explanation: "微信的英文名称是 WeChat，由腾讯公司开发。" },
  { id: 2, question: "微信扫一扫功能可以识别哪些码？", options: ["只能识别二维码", "只能识别条形码", "二维码和条形码都可以", "只能识别微信专属码"], answer: 2, explanation: "微信扫一扫可以识别二维码、条形码等多种码型。" },
  { id: 3, question: "微信朋友圈最多可以发几张照片？", options: ["6张", "9张", "12张", "不限张数"], answer: 1, explanation: "微信朋友圈单次最多可以发布 9 张照片。" },
  { id: 4, question: "微信红包单次最高可发多少元？", options: ["100元", "200元", "500元", "1000元"], answer: 2, explanation: "微信红包单次最高可发 500 元。" },
  { id: 5, question: "以下哪个是微信小程序的正确描述？", options: ["需要下载安装", "用完即走无需安装", "只能在安卓上使用", "需要关注公众号才能用"], answer: 1, explanation: "微信小程序无需下载安装，即用即走，体验轻便。" },
];

const JSON_TEMPLATE = JSON.stringify([
  { question: "题目内容写在这里？", options: ["选项A", "选项B", "选项C", "选项D"], answer: 0, explanation: "解析：正确答案是A，原因……" },
  { question: "第二道题示例？", options: ["选项A", "选项B", "选项C", "选项D"], answer: 2, explanation: "正确答案是C。" }
], null, 2);

const CSV_TEMPLATE = `题目,选项A,选项B,选项C,选项D,正确答案(0-3),解析
微信的英文名称是什么？,WeChat,WePay,WeLink,WeTalk,0,微信英文名是WeChat
微信朋友圈最多发几张图？,6张,9张,12张,不限,1,最多9张`;

// ── helpers ──────────────────────────────────────────────────────────────────
function parseJSON(text) {
  const arr = JSON.parse(text);
  if (!Array.isArray(arr)) throw new Error("JSON 必须是数组");
  return arr.map((item, i) => {
    if (!item.question) throw new Error(`第 ${i + 1} 条缺少 question 字段`);
    if (!Array.isArray(item.options) || item.options.length < 2) throw new Error(`第 ${i + 1} 条 options 至少需要 2 个`);
    if (typeof item.answer !== "number" || item.answer < 0 || item.answer >= item.options.length)
      throw new Error(`第 ${i + 1} 条 answer 超出范围`);
    return { id: i + 1, question: item.question, options: item.options, answer: item.answer, explanation: item.explanation || "" };
  });
}

function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) throw new Error("CSV 至少需要标题行 + 1 条数据");
  return lines.slice(1).map((line, i) => {
    const cols = line.split(",");
    if (cols.length < 6) throw new Error(`第 ${i + 2} 行列数不足，需要至少6列`);
    const [question, a, b, c, d, ansRaw, explanation = ""] = cols;
    const answer = parseInt(ansRaw, 10);
    const options = [a, b, c, d].filter(Boolean);
    if (isNaN(answer) || answer < 0 || answer >= options.length) throw new Error(`第 ${i + 2} 行正确答案格式错误`);
    return { id: i + 1, question: question.trim(), options: options.map(o => o.trim()), answer, explanation: explanation.trim() };
  });
}

function parseXLSX(buffer) {
  const wb = XLSX.read(buffer, { type: "array" });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });
  if (rows.length < 2) throw new Error("表格至少需要标题行 + 1 条数据");
  return rows.slice(1).filter(r => r[0]).map((row, i) => {
    const [question, a, b, c, d, ansRaw, explanation = ""] = row;
    const options = [a, b, c, d].filter(v => v !== "");
    const answer = parseInt(ansRaw, 10);
    if (isNaN(answer) || answer < 0 || answer >= options.length) throw new Error(`第 ${i + 2} 行正确答案格式错误`);
    return { id: i + 1, question: String(question).trim(), options: options.map(o => String(o).trim()), answer, explanation: String(explanation).trim() };
  });
}

// ── sub-components ───────────────────────────────────────────────────────────
function QRCodeSVG({ size = 160 }) {
  const cells = [];
  const fp = (ox, oy) => {
    cells.push(<rect key={`f${ox}${oy}`} x={ox} y={oy} width={7} height={7} fill="#111" />);
    cells.push(<rect key={`fi${ox}${oy}`} x={ox+1} y={oy+1} width={5} height={5} fill="white" />);
    cells.push(<rect key={`fc${ox}${oy}`} x={ox+2} y={oy+2} width={3} height={3} fill="#111" />);
  };
  fp(0,0); fp(14,0); fp(0,14);
  for (let i=8;i<13;i++) if(i%2===0) {
    cells.push(<rect key={`th${i}`} x={i} y={6} width={1} height={1} fill="#111" />);
    cells.push(<rect key={`tv${i}`} x={6} y={i} width={1} height={1} fill="#111" />);
  }
  [[1,0,1,0,0,1,0,0,1,1],[0,1,1,0,1,0,1,0,0,1],[1,1,0,1,0,0,1,1,0,0],[0,0,1,1,1,0,0,1,1,0],[1,0,0,0,1,1,0,0,1,1],[0,1,0,1,0,1,1,0,0,1],[1,0,1,0,1,0,0,1,0,0],[0,1,0,1,0,0,1,0,1,1],[1,1,0,0,1,0,1,0,0,1],[0,0,1,0,0,1,0,1,1,0]].forEach((row,ri)=>row.forEach((cell,ci)=>{if(cell){const x=9+ci,y=9+ri;if(!(x>=14&&y<=7)&&!(x<=7&&y>=14))cells.push(<rect key={`d${ri}${ci}`} x={x} y={y} width={1} height={1} fill="#111"/>);}}));
  return <svg width={size} height={size} viewBox="0 0 21 21" style={{imageRendering:"pixelated"}}><rect width={21} height={21} fill="white"/>{cells}</svg>;
}

function ScanAnimation() {
  return (
    <div style={{position:"relative",display:"inline-block"}}>
      <div style={{width:180,height:180,border:"2px solid #e5e7eb",borderRadius:12,display:"flex",alignItems:"center",justifyContent:"center",background:"#fafafa",position:"relative",overflow:"hidden"}}>
        {[{top:8,left:8,borderTop:"3px solid #07c160",borderLeft:"3px solid #07c160"},{top:8,right:8,borderTop:"3px solid #07c160",borderRight:"3px solid #07c160"},{bottom:8,left:8,borderBottom:"3px solid #07c160",borderLeft:"3px solid #07c160"},{bottom:8,right:8,borderBottom:"3px solid #07c160",borderRight:"3px solid #07c160"}].map((s,i)=><div key={i} style={{position:"absolute",width:18,height:18,...s}}/>)}
        <QRCodeSVG size={120}/>
        <div style={{position:"absolute",left:14,right:14,height:2,background:"linear-gradient(90deg,transparent,#07c160,transparent)",animation:"scanLine 2s ease-in-out infinite"}}/>
      </div>
    </div>
  );
}

function WeChatIcon({ size = 28 }) {
  return (
    <svg width={size} height={size} viewBox="0 0 28 28" fill="none">
      <circle cx="14" cy="14" r="14" fill="#07c160"/>
      <ellipse cx="10.5" cy="13" rx="5.5" ry="4" fill="white"/>
      <circle cx="8.8" cy="12.5" r="0.9" fill="#07c160"/>
      <circle cx="11.8" cy="12.5" r="0.9" fill="#07c160"/>
      <ellipse cx="18.5" cy="16" rx="4.5" ry="3.3" fill="white"/>
      <circle cx="17.1" cy="15.7" r="0.75" fill="#07c160"/>
      <circle cx="19.7" cy="15.7" r="0.75" fill="#07c160"/>
    </svg>
  );
}

const GLOBAL_STYLE = `
  @keyframes scanLine{0%{top:14px;opacity:0}10%{opacity:1}90%{opacity:1}100%{top:166px;opacity:0}}
  @keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}
  @keyframes popIn{0%{transform:scale(0.96)}60%{transform:scale(1.03)}100%{transform:scale(1)}}
  @keyframes slideUp{from{opacity:0;transform:translateY(20px)}to{opacity:1;transform:translateY(0)}}
  @keyframes shake{0%,100%{transform:translateX(0)}25%{transform:translateX(-6px)}75%{transform:translateX(6px)}}
  * { box-sizing: border-box; }
`;

// ── ImportModal ───────────────────────────────────────────────────────────────
function ImportModal({ onImport, onClose }) {
  const [tab, setTab] = useState("file"); // file | json | csv | template
  const [jsonText, setJsonText] = useState("");
  const [csvText, setCsvText] = useState("");
  const [error, setError] = useState("");
  const [preview, setPreview] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef();

  const handleParsed = (data) => {
    setError("");
    setPreview(data);
  };

  const tryParse = (fn, arg) => {
    try { handleParsed(fn(arg)); }
    catch (e) { setError(e.message); setPreview(null); }
  };

  const handleFile = (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (ext === "json") {
      const r = new FileReader();
      r.onload = e => tryParse(parseJSON, e.target.result);
      r.readAsText(file, "utf-8");
    } else if (ext === "csv") {
      const r = new FileReader();
      r.onload = e => tryParse(parseCSV, e.target.result);
      r.readAsText(file, "utf-8");
    } else if (ext === "xlsx" || ext === "xls") {
      const r = new FileReader();
      r.onload = e => tryParse(parseXLSX, new Uint8Array(e.target.result));
      r.readAsArrayBuffer(file);
    } else {
      setError("不支持的文件格式，请上传 .json / .csv / .xlsx 文件");
    }
  };

  const downloadTemplate = (type) => {
    const content = type === "json" ? JSON_TEMPLATE : CSV_TEMPLATE;
    const mime = type === "json" ? "application/json" : "text/csv;charset=utf-8;";
    const blob = new Blob([content], { type: mime });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = `题库模板.${type}`;
    a.click();
  };

  const downloadXLSXTemplate = () => {
    const ws = XLSX.utils.aoa_to_sheet([
      ["题目", "选项A", "选项B", "选项C", "选项D", "正确答案(0-3)", "解析"],
      ["微信的英文名称是什么？", "WeChat", "WePay", "WeLink", "WeTalk", 0, "微信英文名是WeChat"],
      ["微信朋友圈最多发几张图？", "6张", "9张", "12张", "不限", 1, "最多9张"],
    ]);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "题库");
    XLSX.writeFile(wb, "题库模板.xlsx");
  };

  const btn = (active, label, onClick) => (
    <button onClick={onClick} style={{
      padding:"7px 14px", borderRadius:8, border:"none", cursor:"pointer", fontSize:13, fontWeight:600,
      background: active ? "#07c160" : "#f3f4f6",
      color: active ? "white" : "#374151",
      transition:"all 0.15s",
    }}>{label}</button>
  );

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,0.45)",zIndex:100,display:"flex",alignItems:"flex-end",justifyContent:"center",animation:"fadeIn 0.2s ease"}}>
      <div style={{background:"white",borderRadius:"20px 20px 0 0",width:"100%",maxWidth:480,maxHeight:"90vh",overflow:"auto",padding:"24px 20px 32px",animation:"slideUp 0.3s ease"}}>
        
        {/* Header */}
        <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:20}}>
          <div style={{display:"flex",alignItems:"center",gap:8}}>
            <span style={{fontSize:22}}>📦</span>
            <span style={{fontSize:17,fontWeight:700,color:"#111827"}}>导入题库</span>
          </div>
          <button onClick={onClose} style={{width:30,height:30,borderRadius:"50%",border:"none",background:"#f3f4f6",cursor:"pointer",fontSize:16,color:"#6b7280"}}>✕</button>
        </div>

        {/* Tabs */}
        <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap"}}>
          {btn(tab==="file","📁 上传文件",()=>setTab("file"))}
          {btn(tab==="json","{ } JSON",()=>setTab("json"))}
          {btn(tab==="csv","📊 CSV文本",()=>setTab("csv"))}
          {btn(tab==="template","📄 下载模板",()=>setTab("template"))}
        </div>

        {/* File upload */}
        {tab==="file" && (
          <div>
            <div
              onDragOver={e=>{e.preventDefault();setDragOver(true)}}
              onDragLeave={()=>setDragOver(false)}
              onDrop={e=>{e.preventDefault();setDragOver(false);handleFile(e.dataTransfer.files[0])}}
              onClick={()=>fileRef.current.click()}
              style={{
                border:`2px dashed ${dragOver?"#07c160":"#d1d5db"}`,
                borderRadius:12, padding:"32px 20px", textAlign:"center",
                cursor:"pointer", background: dragOver?"#f0fdf4":"#fafafa",
                transition:"all 0.15s", marginBottom:12,
              }}
            >
              <div style={{fontSize:36,marginBottom:8}}>📂</div>
              <div style={{fontSize:14,fontWeight:600,color:"#374151",marginBottom:4}}>点击或拖拽文件到此处</div>
              <div style={{fontSize:12,color:"#9ca3af"}}>支持 .json / .csv / .xlsx / .xls</div>
            </div>
            <input ref={fileRef} type="file" accept=".json,.csv,.xlsx,.xls" style={{display:"none"}} onChange={e=>handleFile(e.target.files[0])}/>
          </div>
        )}

        {/* JSON text */}
        {tab==="json" && (
          <div>
            <div style={{fontSize:12,color:"#6b7280",marginBottom:8}}>粘贴 JSON 格式题库（数组，每项含 question/options/answer）</div>
            <textarea
              value={jsonText}
              onChange={e=>setJsonText(e.target.value)}
              placeholder={JSON_TEMPLATE}
              style={{width:"100%",height:180,border:"1.5px solid #e5e7eb",borderRadius:10,padding:12,fontSize:12,fontFamily:"monospace",resize:"vertical",outline:"none",color:"#374151"}}
            />
            <button onClick={()=>tryParse(parseJSON,jsonText)} style={{marginTop:8,padding:"9px 20px",background:"#07c160",color:"white",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>
              解析 JSON
            </button>
          </div>
        )}

        {/* CSV text */}
        {tab==="csv" && (
          <div>
            <div style={{fontSize:12,color:"#6b7280",marginBottom:8}}>粘贴 CSV 格式（第一行为标题，字段：题目,选项A,选项B,选项C,选项D,正确答案(0-3),解析）</div>
            <textarea
              value={csvText}
              onChange={e=>setCsvText(e.target.value)}
              placeholder={CSV_TEMPLATE}
              style={{width:"100%",height:160,border:"1.5px solid #e5e7eb",borderRadius:10,padding:12,fontSize:12,fontFamily:"monospace",resize:"vertical",outline:"none",color:"#374151"}}
            />
            <button onClick={()=>tryParse(parseCSV,csvText)} style={{marginTop:8,padding:"9px 20px",background:"#07c160",color:"white",border:"none",borderRadius:8,cursor:"pointer",fontSize:13,fontWeight:600}}>
              解析 CSV
            </button>
          </div>
        )}

        {/* Template download */}
        {tab==="template" && (
          <div style={{display:"flex",flexDirection:"column",gap:10}}>
            {[
              {icon:"📋",label:"下载 JSON 模板",sub:"适合开发者，结构清晰",fn:()=>downloadTemplate("json")},
              {icon:"📊",label:"下载 CSV 模板",sub:"适合 Excel 编辑后另存",fn:()=>downloadTemplate("csv")},
              {icon:"📗",label:"下载 Excel 模板",sub:"直接用 Excel/WPS 填写",fn:downloadXLSXTemplate},
            ].map((item,i)=>(
              <button key={i} onClick={item.fn} style={{display:"flex",alignItems:"center",gap:12,padding:"14px 16px",border:"1.5px solid #e5e7eb",borderRadius:12,background:"white",cursor:"pointer",textAlign:"left",width:"100%"}}>
                <span style={{fontSize:24}}>{item.icon}</span>
                <div>
                  <div style={{fontSize:14,fontWeight:600,color:"#111827"}}>{item.label}</div>
                  <div style={{fontSize:12,color:"#9ca3af"}}>{item.sub}</div>
                </div>
                <span style={{marginLeft:"auto",color:"#07c160",fontSize:18}}>⬇</span>
              </button>
            ))}
            <div style={{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:10,padding:"12px 14px",fontSize:12,color:"#166534",lineHeight:1.7}}>
              💡 填写完模板后，切换到「上传文件」或「CSV文本」标签页导入
            </div>
          </div>
        )}

        {/* Error */}
        {error && (
          <div style={{marginTop:12,padding:"10px 14px",background:"#fff1f2",border:"1px solid #fecdd3",borderRadius:10,fontSize:13,color:"#b91c1c",animation:"shake 0.3s ease"}}>
            ❌ {error}
          </div>
        )}

        {/* Preview */}
        {preview && (
          <div style={{marginTop:16,animation:"fadeIn 0.3s ease"}}>
            <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:10}}>
              <span style={{fontSize:14,fontWeight:700,color:"#166534"}}>✅ 解析成功</span>
              <span style={{background:"#07c160",color:"white",borderRadius:99,padding:"2px 10px",fontSize:12,fontWeight:600}}>{preview.length} 道题</span>
            </div>
            <div style={{maxHeight:160,overflow:"auto",border:"1px solid #bbf7d0",borderRadius:10,background:"#f0fdf4"}}>
              {preview.map((q,i)=>(
                <div key={i} style={{padding:"8px 12px",borderBottom:i<preview.length-1?"1px solid #d1fae5":"none"}}>
                  <div style={{fontSize:12,fontWeight:600,color:"#166534"}}>Q{i+1}. {q.question}</div>
                  <div style={{fontSize:11,color:"#4b7c5e",marginTop:2}}>
                    正确答案：{q.options[q.answer]}
                  </div>
                </div>
              ))}
            </div>
            <button
              onClick={()=>onImport(preview)}
              style={{marginTop:12,width:"100%",padding:"13px 0",background:"linear-gradient(135deg,#07c160,#05a350)",color:"white",border:"none",borderRadius:11,fontSize:15,fontWeight:700,cursor:"pointer",boxShadow:"0 3px 12px rgba(7,193,96,0.3)"}}>
              导入 {preview.length} 道题并开始 →
            </button>
          </div>
        )}
      </div>
    </div>
  );
}

// ── Main App ──────────────────────────────────────────────────────────────────
export default function App() {
  const [quizData, setQuizData] = useState(DEFAULT_QUIZ);
  const [phase, setPhase] = useState("landing");
  const [currentQ, setCurrentQ] = useState(0);
  const [selected, setSelected] = useState(null);
  const [answers, setAnswers] = useState([]);
  const [showExplain, setShowExplain] = useState(false);
  const [scanProgress, setScanProgress] = useState(0);
  const [showImport, setShowImport] = useState(false);
  const timerRef = useRef(null);

  const startScan = () => {
    setPhase("scanning"); setScanProgress(0);
    let p = 0;
    timerRef.current = setInterval(() => {
      p += 2; setScanProgress(p);
      if (p >= 100) { clearInterval(timerRef.current); setTimeout(() => { setPhase("quiz"); setCurrentQ(0); setAnswers([]); setSelected(null); setShowExplain(false); }, 400); }
    }, 40);
  };

  useEffect(() => () => clearInterval(timerRef.current), []);

  const handleImport = (data) => {
    setQuizData(data);
    setShowImport(false);
    setTimeout(() => startScan(), 300);
  };

  const handleSelect = (idx) => { if (selected !== null) return; setSelected(idx); setShowExplain(true); };

  const handleNext = () => {
    const newAnswers = [...answers, selected];
    if (currentQ + 1 < quizData.length) { setAnswers(newAnswers); setCurrentQ(currentQ + 1); setSelected(null); setShowExplain(false); }
    else { setAnswers(newAnswers); setPhase("result"); }
  };

  const score = answers.filter((a, i) => a === quizData[i]?.answer).length;

  const getGrade = (s, total) => {
    const pct = s / total;
    if (pct === 1) return { label: "满分达人 🏆", color: "#f59e0b" };
    if (pct >= 0.8) return { label: "优秀！", color: "#07c160" };
    if (pct >= 0.6) return { label: "良好", color: "#3b82f6" };
    return { label: "继续加油", color: "#ef4444" };
  };

  const containerStyle = { minHeight:"100vh", background:"linear-gradient(160deg,#f0fdf4 0%,#dcfce7 40%,#f0f9ff 100%)", fontFamily:"'PingFang SC','Hiragino Sans GB','Microsoft YaHei',sans-serif", display:"flex", flexDirection:"column", alignItems:"center", justifyContent:"center", padding:"24px 16px" };
  const cardStyle = { background:"white", borderRadius:20, boxShadow:"0 8px 40px rgba(7,193,96,0.12),0 2px 12px rgba(0,0,0,0.06)", padding:"28px 24px", width:"100%", maxWidth:390, animation:"fadeIn 0.4s ease" };

  // ── LANDING ──
  if (phase === "landing") return (
    <div style={containerStyle}>
      <style>{GLOBAL_STYLE}</style>
      {showImport && <ImportModal onImport={handleImport} onClose={() => setShowImport(false)} />}
      <div style={cardStyle}>
        <div style={{textAlign:"center",marginBottom:22}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"center",gap:10,marginBottom:6}}>
            <WeChatIcon/><span style={{fontSize:20,fontWeight:700,color:"#111827"}}>微信答题挑战</span>
          </div>
          <p style={{color:"#6b7280",fontSize:13,margin:0}}>扫描二维码，开启知识之旅</p>
        </div>

        <div style={{display:"flex",justifyContent:"center",marginBottom:22}}><ScanAnimation/></div>

        {/* Quiz bank info */}
        <div style={{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:12,padding:"12px 14px",marginBottom:16,display:"flex",alignItems:"center",gap:10}}>
          <span style={{fontSize:20}}>📋</span>
          <div style={{flex:1}}>
            <div style={{fontSize:13,fontWeight:600,color:"#166534"}}>当前题库</div>
            <div style={{fontSize:12,color:"#4b7c5e"}}>共 {quizData.length} 道题 · {quizData === DEFAULT_QUIZ ? "默认微信知识题" : "自定义题库"}</div>
          </div>
          {quizData !== DEFAULT_QUIZ && (
            <button onClick={()=>setQuizData(DEFAULT_QUIZ)} style={{padding:"4px 10px",border:"1px solid #bbf7d0",borderRadius:6,background:"white",color:"#166534",fontSize:11,cursor:"pointer",fontWeight:600}}>
              重置
            </button>
          )}
        </div>

        {/* Import button */}
        <button
          onClick={()=>setShowImport(true)}
          style={{width:"100%",padding:"12px 0",border:"2px dashed #07c160",borderRadius:12,background:"#f0fdf4",color:"#07c160",fontSize:14,fontWeight:600,cursor:"pointer",marginBottom:12,display:"flex",alignItems:"center",justifyContent:"center",gap:8,transition:"all 0.15s"}}
          onMouseEnter={e=>{e.currentTarget.style.background="#dcfce7"}}
          onMouseLeave={e=>{e.currentTarget.style.background="#f0fdf4"}}
        >
          <span style={{fontSize:18}}>📦</span> 导入自定义题库
        </button>

        <button
          onClick={startScan}
          style={{width:"100%",padding:"14px 0",background:"linear-gradient(135deg,#07c160,#05a350)",color:"white",border:"none",borderRadius:12,fontSize:16,fontWeight:600,cursor:"pointer",boxShadow:"0 4px 16px rgba(7,193,96,0.35)"}}
        >
          📷 扫码开始答题
        </button>
      </div>
    </div>
  );

  // ── SCANNING ──
  if (phase === "scanning") return (
    <div style={containerStyle}>
      <style>{GLOBAL_STYLE}</style>
      <div style={cardStyle}>
        <div style={{textAlign:"center",marginBottom:24}}>
          <div style={{fontSize:18,fontWeight:700,color:"#111827",marginBottom:4}}>正在识别二维码</div>
          <p style={{color:"#6b7280",fontSize:13,margin:0}}>请将二维码对准扫描框</p>
        </div>
        <div style={{display:"flex",justifyContent:"center",marginBottom:20}}><ScanAnimation/></div>
        <div style={{marginBottom:12}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:6}}>
            <span style={{fontSize:12,color:"#6b7280"}}>识别进度</span>
            <span style={{fontSize:12,color:"#07c160",fontWeight:600}}>{scanProgress}%</span>
          </div>
          <div style={{height:6,background:"#e5e7eb",borderRadius:99}}>
            <div style={{height:"100%",width:`${scanProgress}%`,background:"linear-gradient(90deg,#07c160,#05a350)",borderRadius:99,transition:"width 0.04s linear"}}/>
          </div>
        </div>
        <div style={{textAlign:"center",color:"#9ca3af",fontSize:13}}>
          {scanProgress < 60 ? "🔍 扫描中..." : scanProgress < 90 ? "✅ 二维码识别成功！" : "🚀 即将进入答题..."}
        </div>
      </div>
    </div>
  );

  // ── QUIZ ──
  if (phase === "quiz") {
    const q = quizData[currentQ];
    const isCorrect = selected === q.answer;
    return (
      <div style={containerStyle}>
        <style>{GLOBAL_STYLE}</style>
        <div style={{...cardStyle, padding:"22px 20px 26px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <WeChatIcon size={22}/><span style={{fontSize:13,fontWeight:600,color:"#374151"}}>微信答题</span>
            </div>
            <div style={{background:"#f0fdf4",border:"1px solid #bbf7d0",borderRadius:99,padding:"3px 12px",fontSize:12,color:"#07c160",fontWeight:600}}>
              {currentQ+1} / {quizData.length}
            </div>
          </div>
          <div style={{height:4,background:"#f3f4f6",borderRadius:99,marginBottom:20}}>
            <div style={{height:"100%",width:`${((currentQ+1)/quizData.length)*100}%`,background:"linear-gradient(90deg,#07c160,#05a350)",borderRadius:99,transition:"width 0.3s ease"}}/>
          </div>
          <div style={{fontSize:15,fontWeight:700,color:"#111827",lineHeight:1.55,marginBottom:18}}>{currentQ+1}. {q.question}</div>
          <div style={{display:"flex",flexDirection:"column",gap:9,marginBottom:18}}>
            {q.options.map((opt, i) => {
              let bg="white", border="1.5px solid #e5e7eb", color="#374151", icon=null;
              if (selected !== null) {
                if (i===q.answer) { bg="#f0fdf4"; border="1.5px solid #07c160"; color="#166534"; icon=<span style={{marginLeft:"auto",color:"#07c160",fontSize:15}}>✓</span>; }
                else if (i===selected && selected!==q.answer) { bg="#fff1f2"; border="1.5px solid #ef4444"; color="#b91c1c"; icon=<span style={{marginLeft:"auto",color:"#ef4444",fontSize:15}}>✗</span>; }
                else { bg="#f9fafb"; color="#9ca3af"; }
              }
              return (
                <button key={i} onClick={()=>handleSelect(i)} style={{width:"100%",padding:"11px 14px",background:bg,border,borderRadius:10,fontSize:13,color,cursor:selected!==null?"default":"pointer",display:"flex",alignItems:"center",gap:10,textAlign:"left",transition:"all 0.15s",animation:selected===i?"popIn 0.2s ease":"none"}}>
                  <span style={{width:22,height:22,borderRadius:"50%",background:i===q.answer&&selected!==null?"#07c160":i===selected&&selected!==q.answer?"#ef4444":"#f3f4f6",color:(i===q.answer||i===selected)&&selected!==null?"white":"#6b7280",display:"flex",alignItems:"center",justifyContent:"center",fontSize:11,fontWeight:700,flexShrink:0}}>
                    {String.fromCharCode(65+i)}
                  </span>
                  <span style={{flex:1}}>{opt}</span>{icon}
                </button>
              );
            })}
          </div>
          {showExplain && (
            <div style={{background:isCorrect?"#f0fdf4":"#fff7ed",border:`1px solid ${isCorrect?"#bbf7d0":"#fed7aa"}`,borderRadius:10,padding:"11px 14px",marginBottom:14,animation:"fadeIn 0.3s ease"}}>
              <div style={{fontSize:13,fontWeight:700,color:isCorrect?"#166534":"#92400e",marginBottom:3}}>{isCorrect?"🎉 回答正确！":"💡 解析"}</div>
              {q.explanation && <div style={{fontSize:12,color:isCorrect?"#166534":"#78350f",lineHeight:1.6}}>{q.explanation}</div>}
            </div>
          )}
          {selected !== null && (
            <button onClick={handleNext} style={{width:"100%",padding:"13px 0",background:"linear-gradient(135deg,#07c160,#05a350)",color:"white",border:"none",borderRadius:11,fontSize:15,fontWeight:600,cursor:"pointer",animation:"fadeIn 0.3s ease",boxShadow:"0 3px 12px rgba(7,193,96,0.3)"}}>
              {currentQ+1 < quizData.length ? "下一题 →" : "查看结果 🏆"}
            </button>
          )}
        </div>
      </div>
    );
  }

  // ── RESULT ──
  if (phase === "result") {
    const grade = getGrade(score, quizData.length);
    return (
      <div style={containerStyle}>
        <style>{GLOBAL_STYLE}</style>
        <div style={cardStyle}>
          <div style={{textAlign:"center",marginBottom:20}}>
            <div style={{fontSize:48,marginBottom:6}}>{score===quizData.length?"🏆":score/quizData.length>=0.8?"🎉":score/quizData.length>=0.6?"👍":"📚"}</div>
            <div style={{fontSize:20,fontWeight:800,color:grade.color,marginBottom:2}}>{grade.label}</div>
            <div style={{color:"#6b7280",fontSize:13}}>答题完成！共 {quizData.length} 道题</div>
          </div>
          <div style={{display:"flex",justifyContent:"center",marginBottom:20}}>
            <div style={{width:100,height:100,borderRadius:"50%",background:`conic-gradient(${grade.color} ${(score/quizData.length)*360}deg,#e5e7eb 0deg)`,display:"flex",alignItems:"center",justifyContent:"center"}}>
              <div style={{width:74,height:74,borderRadius:"50%",background:"white",display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center"}}>
                <span style={{fontSize:24,fontWeight:800,color:grade.color,lineHeight:1}}>{score}</span>
                <span style={{fontSize:11,color:"#9ca3af"}}>/ {quizData.length}</span>
              </div>
            </div>
          </div>
          <div style={{marginBottom:20,maxHeight:260,overflow:"auto"}}>
            <div style={{fontSize:13,fontWeight:700,color:"#374151",marginBottom:8}}>答题回顾</div>
            {quizData.map((q, i) => {
              const userAns = answers[i]; const correct = userAns === q.answer;
              return (
                <div key={i} style={{display:"flex",alignItems:"flex-start",gap:10,padding:"9px 12px",background:correct?"#f0fdf4":"#fff1f2",borderRadius:9,marginBottom:6,border:`1px solid ${correct?"#bbf7d0":"#fecdd3"}`}}>
                  <span style={{fontSize:14,flexShrink:0,marginTop:1}}>{correct?"✅":"❌"}</span>
                  <div>
                    <div style={{fontSize:11,fontWeight:600,color:correct?"#166534":"#b91c1c",marginBottom:2}}>第{i+1}题</div>
                    <div style={{fontSize:11,color:"#374151",lineHeight:1.4}}>
                      {!correct && <span style={{color:"#ef4444"}}>你选了「{q.options[userAns]}」，</span>}
                      正确答案：「{q.options[q.answer]}」
                    </div>
                  </div>
                </div>
              );
            })}
          </div>
          <div style={{display:"flex",gap:10}}>
            <button onClick={()=>setShowImport(true)&&setPhase("landing")||setPhase("landing")} style={{flex:1,padding:"12px 0",border:"1.5px solid #07c160",borderRadius:11,background:"white",color:"#07c160",fontSize:14,fontWeight:600,cursor:"pointer"}}>
              换题库
            </button>
            <button onClick={()=>{setPhase("landing");setAnswers([]);setSelected(null);setShowExplain(false);}} style={{flex:2,padding:"12px 0",background:"linear-gradient(135deg,#07c160,#05a350)",color:"white",border:"none",borderRadius:11,fontSize:14,fontWeight:600,cursor:"pointer",boxShadow:"0 3px 12px rgba(7,193,96,0.3)"}}>
              🔄 再来一次
            </button>
          </div>
        </div>
        {showImport && <ImportModal onImport={handleImport} onClose={()=>setShowImport(false)}/>}
      </div>
    );
  }
}

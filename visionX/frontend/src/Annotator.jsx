import { useState, useEffect, useRef } from "react";
import { Stage, Layer, Rect, Text } from "react-konva";

const API = "http://localhost:8000";

export default function Annotator({ videoId, frames, videoUrl }) {
  const [idx, setIdx] = useState(0);
  const [boxes, setBoxes] = useState([]);
  const [drawing, setDrawing] = useState(null);
  const [loading, setLoading] = useState(false);
  const [jobId, setJobId] = useState(null);
  const [progress, setProgress] = useState(0);
  const [grouped, setGrouped] = useState({}); // cache all frames: {frameName: [boxes]}

  const frameName = frames[idx];

  // --- Load boxes for current frame (from cache or API) ---
  useEffect(() => {
    if (!frameName) return;
    if (grouped[frameName]) {
      setBoxes(grouped[frameName]);
      return;
    }
    fetch(`${API}/annotations/${videoId}/${frameName}`)
     .then((r) => r.json())
     .then((data) => {
        const b = Array.isArray(data)? data : [];
        setBoxes(b);
        setGrouped(g => ({...g, [frameName]: b}));
      });
  }, [frameName, videoId]);

  // --- Save ---
  const save = async () => {
    await fetch(`${API}/annotations/${videoId}/${frameName}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(boxes),
    });
    setGrouped(g => ({...g, [frameName]: boxes}));
  };

  // --- Auto This Frame (uses C++ engine) ---
  const autoLabel = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API}/auto-label/${videoId}/${frameName}`, { method: "POST" });
      const data = await res.json();
      const mapped = data.map(d => ({ x: d.x, y: d.y, width: d.w, height: d.h, label: d.class_name, conf: d.conf, auto: true }));
      setBoxes(mapped);
    } catch { alert("Auto-label failed"); }
    setLoading(false);
  };

  // --- UPDATED: Auto ALL uses NEW C++ AsyncVideoEngine ---
  const autoLabelAll = async () => {
    if(!confirm(`Auto-label all ${frames.length} frames with C++ async engine?`)) return;
    setLoading(true);
    try {
      // This now hits the async C++ engine we built
      const res = await fetch(`${API}/upload-video-async/${videoId}`, { method: "POST" });
      const { job_id } = await res.json();
      setJobId(job_id);

      // Poll progress
      const poll = async () => {
        const r = await fetch(`${API}/video-job/${job_id}`);
        const s = await r.json();
        setProgress(s.progress);
        if (!s.done) {
          setTimeout(poll, 1000);
        } else {
          // Load all results into cache
          const rr = await fetch(`${API}/video-job/${job_id}/results`);
          const dd = await rr.json();
          const newGrouped = {};
          frames.forEach((f, i) => {
            const dets = dd.grouped[i] || [];
            newGrouped[f] = dets.map(d => ({ x: d.x, y: d.y, width: d.w, height: d.h, label: d.class_name, conf: d.conf, auto: true }));
          });
          setGrouped(newGrouped);
          setBoxes(newGrouped[frameName] || []);
          setLoading(false);
          alert(`Done: ${s.processed} frames processed by C++`);
        }
      };
      poll();
    } catch (e) {
      console.error(e);
      alert("Bulk auto-label failed");
      setLoading(false);
    }
  };

  const trackToNext = async () => {
    if (idx >= frames.length - 1 || boxes.length === 0) return;
    setLoading(true);
    const nextFrame = frames[idx + 1];
    try {
      const res = await fetch(`${API}/track/${videoId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ from_frame: frameName, to_frame: nextFrame, boxes }),
      });
      const tracked = await res.json();
      const mapped = tracked.map(b => ({...b, tracked: true}));
      setIdx(idx + 1);
      setBoxes(mapped);
      setGrouped(g => ({...g, [nextFrame]: mapped}));
    } catch { alert("Tracking failed"); }
    setLoading(false);
  };

  const exportCOCO = async () => {
    const res = await fetch(`${API}/export/${videoId}`);
    const data = await res.json();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url; a.download = `${videoId}_coco.json`; a.click();
    URL.revokeObjectURL(url);
  };

  // --- Konva drawing ---
  const handleMouseDown = (e) => {
    const pos = e.target.getStage().getPointerPosition();
    setDrawing({ x: pos.x, y: pos.y, width: 0, height: 0, label: "object" });
  };
  const handleMouseMove = (e) => {
    if (!drawing) return;
    const pos = e.target.getStage().getPointerPosition();
    setDrawing({...drawing, width: pos.x - drawing.x, height: pos.y - drawing.y });
  };
  const handleMouseUp = () => {
    if (drawing && Math.abs(drawing.width) > 10 && Math.abs(drawing.height) > 10) {
      const norm = {
        x: drawing.width < 0? drawing.x + drawing.width : drawing.x,
        y: drawing.height < 0? drawing.y + drawing.height : drawing.y,
        width: Math.abs(drawing.width),
        height: Math.abs(drawing.height),
        label: "object",
      };
      setBoxes([...boxes, norm]);
    }
    setDrawing(null);
  };

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
        <button onClick={() => setIdx(Math.max(0, idx - 1))}>⬅️ Prev</button>
        <b>{idx + 1} / {frames.length} - {frameName}</b>
        <button onClick={() => setIdx(Math.min(frames.length - 1, idx + 1))}>Next ➡️</button>
        <button onClick={autoLabel} disabled={loading} style={{ background: "#6d28d9", color: "#fff", padding: "8px 12px" }}>
          {loading? "..." : "🤖 Auto This"}
        </button>
        <button onClick={autoLabelAll} disabled={loading} style={{ background: "#4f46e5", color: "#fff", padding: "8px 12px" }}>
          🚀 Auto ALL (C++ Async)
        </button>
        <button onClick={trackToNext} disabled={loading} style={{ background: "#059669", color: "#fff", padding: "8px 12px" }}>➡️ Track</button>
        <button onClick={save} style={{ background: "#000", color: "#fff", padding: "8px 12px" }}>💾 Save</button>
        <button onClick={exportCOCO} style={{ background: "#f59e0b", color: "#000", padding: "8px 12px" }}>📥 COCO</button>
        <button onClick={() => setBoxes([])} style={{ color: "red" }}>Clear</button>
      </div>

      {jobId && (
        <div style={{ marginBottom: 12 }}>
          <div style={{ height: 8, background: "#222", borderRadius: 4 }}>
            <div style={{ height: "100%", width: `${progress * 100}%`, background: "#00ff00", transition: "width 0.3s" }} />
          </div>
          <div style={{ fontSize: 12, marginTop: 4 }}>C++ Job {jobId}: {Math.round(progress*100)}% - {loading? "Processing in C++ thread pool" : "Done"}</div>
        </div>
      )}

      <div style={{ display: "flex", gap: 20 }}>
        <Stage
          width={960} height={540}
          onMouseDown={handleMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp}
          style={{ border: "2px solid #333", backgroundImage: `url(${API}/frames/${videoId}/${frameName})`, backgroundSize: "cover", backgroundColor: "#000" }}
        >
          <Layer>
            {boxes.map((b, i) => (
              <Text key={`l-${i}`} x={b.x} y={b.y - 16} text={`${b.label} ${b.conf? Math.round(b.conf*100)+'%' : ''}`} fill="white" fontSize={13} />
            ))}
            {boxes.map((b, i) => (
              <Rect key={i} x={b.x} y={b.y} width={b.width} height={b.height}
                stroke={b.auto? "#f59e0b" : b.tracked? "#06b6d4" : "lime"} strokeWidth={2} dash={b.auto? [6, 4] : []} />
            ))}
            {drawing && <Rect x={drawing.x} y={drawing.y} width={drawing.width} height={drawing.height} stroke="red" strokeWidth={2} dash={[5,5]} />}
          </Layer>
        </Stage>

        <div style={{ width: 220 }}>
          <h4>Boxes ({boxes.length})</h4>
          {boxes.map((b, i) => (
            <div key={i} style={{ display:"flex", gap:4, marginBottom:6 }}>
              <input value={b.label} onChange={(e)=>{ const c=[...boxes]; c[i].label=e.target.value; setBoxes(c); }} style={{ width: 100 }} />
              <button onClick={()=> setBoxes(boxes.filter((_,j)=>j!==i))}>x</button>
            </div>
          ))}
          <div style={{marginTop:12, fontSize:12}}>
            <div>🟩 Manual</div><div>🟧 Auto (C++)</div><div>🟦 Tracked</div>
          </div>
        </div>
      </div>
    </div>
  );
}
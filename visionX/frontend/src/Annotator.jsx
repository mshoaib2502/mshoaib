import { useState, useEffect } from "react";
import { Stage, Layer, Rect, Text } from "react-konva";

const API = "http://localhost:8000";

export default function Annotator({ videoId, frames }) {
  const [idx, setIdx] = useState(0);
  const [boxes, setBoxes] = useState([]);
  const [drawing, setDrawing] = useState(null);
  const [loading, setLoading] = useState(false);

  const frameName = frames[idx];

  useEffect(() => {
    if (!frameName) return;
    fetch(`${API}/annotations/${videoId}/${frameName}`)
     .then((r) => r.json())
     .then((data) => setBoxes(Array.isArray(data)? data : []));
  }, [frameName, videoId]);

  const save = async () => {
    await fetch(`${API}/annotations/${videoId}/${frameName}`, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(boxes),
    });
    alert(`Saved ${boxes.length} boxes for ${frameName}`);
  };

  const autoLabel = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API}/auto-label/${videoId}/${frameName}`, { method: "POST" });
      const data = await res.json();
      setBoxes(data);
    } catch { alert("Auto-label failed"); }
    setLoading(false);
  };

  const autoLabelAll = async () => {
    if(!confirm(`Auto-label all ${frames.length} frames? This may take 1-2 min`)) return;
    setLoading(true);
    try {
      const res = await fetch(`${API}/auto-label-all/${videoId}`, { method: "POST" });
      const data = await res.json();
      alert(`Done: ${data.frames_processed} frames, ${data.total_boxes} boxes`);
      // reload current
      const r = await fetch(`${API}/annotations/${videoId}/${frameName}`);
      setBoxes(await r.json());
    } catch { alert("Bulk auto-label failed"); }
    setLoading(false);
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
      setIdx(idx + 1);
      setBoxes(tracked);
    } catch { alert("Tracking failed"); }
    setLoading(false);
  };

  const exportCOCO = async () => {
    const res = await fetch(`${API}/export/${videoId}`);
    const data = await res.json();
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${videoId}_coco.json`;
    a.click();
    URL.revokeObjectURL(url);
  };

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
      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
        <button onClick={() => setIdx(Math.max(0, idx - 1))}>⬅️ Prev</button>
        <b>{idx + 1} / {frames.length}</b>
        <button onClick={() => setIdx(Math.min(frames.length - 1, idx + 1))}>Next ➡️</button>
        <button onClick={autoLabel} disabled={loading} style={{ background: "#6d28d9", color: "#fff" }}>
          {loading? "..." : "🤖 Auto This"}
        </button>
        <button onClick={autoLabelAll} disabled={loading} style={{ background: "#4f46e5", color: "#fff" }}>
          🚀 Auto ALL Frames
        </button>
        <button onClick={trackToNext} disabled={loading} style={{ background: "#059669", color: "#fff" }}>
          ➡️ Track
        </button>
        <button onClick={save} style={{ background: "#000", color: "#fff" }}>💾 Save</button>
        <button onClick={exportCOCO} style={{ background: "#f59e0b", color: "#000" }}>📥 Export COCO</button>
        <button onClick={() => setBoxes([])} style={{ color: "red" }}>Clear</button>
      </div>

      <div style={{ display: "flex", gap: 20 }}>
        <Stage
          width={1280} height={720}
          onMouseDown={handleMouseDown} onMouseMove={handleMouseMove} onMouseUp={handleMouseUp}
          style={{ border: "2px solid #333", backgroundImage: `url(${API}/frames/${frameName})`, backgroundSize: "cover" }}
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
            <div>🟩 Manual</div><div>🟧 Auto</div><div>🟦 Tracked</div>
          </div>
        </div>
      </div>
    </div>
  );
}
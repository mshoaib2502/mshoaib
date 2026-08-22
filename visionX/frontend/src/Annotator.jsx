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
    alert(`Saved ${boxes.length} boxes`);
  };

  const autoLabel = async () => {
    setLoading(true);
    try {
      const res = await fetch(`${API}/auto-label/${videoId}/${frameName}`, {
        method: "POST",
      });
      const data = await res.json();
      setBoxes(data);
    } catch (e) {
      alert("Auto-label failed");
    }
    setLoading(false);
  };

  const trackToNext = async () => {
    if (idx >= frames.length - 1) return;
    if (boxes.length === 0) {
      alert("Draw boxes first");
      return;
    }
    setLoading(true);
    const nextFrame = frames[idx + 1];
    try {
      const res = await fetch(`${API}/track/${videoId}`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          from_frame: frameName,
          to_frame: nextFrame,
          boxes,
        }),
      });
      const tracked = await res.json();
      setIdx(idx + 1);
      setBoxes(tracked);
    } catch (e) {
      alert("Tracking failed");
    }
    setLoading(false);
  };

  const handleMouseDown = (e) => {
    const stage = e.target.getStage();
    const pos = stage.getPointerPosition();
    setDrawing({ x: pos.x, y: pos.y, width: 0, height: 0, label: "object" });
  };

  const handleMouseMove = (e) => {
    if (!drawing) return;
    const pos = e.target.getStage().getPointerPosition();
    setDrawing({
     ...drawing,
      width: pos.x - drawing.x,
      height: pos.y - drawing.y,
    });
  };

  const handleMouseUp = () => {
    if (drawing && Math.abs(drawing.width) > 10 && Math.abs(drawing.height) > 10) {
      const normalized = {
        x: drawing.width < 0? drawing.x + drawing.width : drawing.x,
        y: drawing.height < 0? drawing.y + drawing.height : drawing.y,
        width: Math.abs(drawing.width),
        height: Math.abs(drawing.height),
        label: "object",
      };
      setBoxes([...boxes, normalized]);
    }
    setDrawing(null);
  };

  const updateLabel = (i, newLabel) => {
    const copy = [...boxes];
    copy[i].label = newLabel;
    setBoxes(copy);
  };

  if (!frameName) return <div>No frames</div>;

  return (
    <div>
      <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap", alignItems:"center" }}>
        <button onClick={() => setIdx(Math.max(0, idx - 1))}>⬅️ Prev</button>
        <span style={{fontWeight:"bold"}}>{idx + 1} / {frames.length}</span>
        <button onClick={() => setIdx(Math.min(frames.length - 1, idx + 1))}>Next ➡️</button>
        <button onClick={autoLabel} disabled={loading} style={{ background: "#6d28d9", color: "#fff", padding:"6px 12px" }}>
          {loading? "Running..." : "🤖 Auto-Label"}
        </button>
        <button onClick={trackToNext} disabled={loading} style={{ background: "#059669", color: "#fff", padding:"6px 12px" }}>
          ➡️ Track to Next
        </button>
        <button onClick={save} style={{ background: "#000", color: "#fff", padding:"6px 12px" }}>💾 Save</button>
        <button onClick={() => setBoxes([])} style={{ color: "red" }}>Clear</button>
      </div>

      <div style={{display:"flex", gap:20}}>
        <Stage
          width={1280}
          height={720}
          onMouseDown={handleMouseDown}
          onMouseMove={handleMouseMove}
          onMouseUp={handleMouseUp}
          style={{
            border: "2px solid #333",
            backgroundImage: `url(${API}/frames/${frameName})`,
            backgroundSize: "cover",
          }}
        >
          <Layer>
            {boxes.map((b, i) => (
              <React.Fragment key={i}>
                <Rect
                  x={b.x}
                  y={b.y}
                  width={b.width}
                  height={b.height}
                  stroke={b.auto? "#f59e0b" : b.tracked? "#06b6d4" : "lime"}
                  strokeWidth={2}
                  dash={b.auto? [6, 4] : []}
                />
                <Text x={b.x} y={b.y - 16} text={`${b.label} ${b.conf? Math.round(b.conf*100)+'%' : ''}`} fill="white" fontSize={14} fontStyle="bold" />
              </React.Fragment>
            ))}
            {drawing && (
              <Rect x={drawing.x} y={drawing.y} width={drawing.width} height={drawing.height} stroke="red" strokeWidth={2} dash={[5,5]} />
            )}
          </Layer>
        </Stage>

        <div style={{width:200}}>
          <h4>Boxes ({boxes.length})</h4>
          {boxes.map((b,i)=>(
            <div key={i} style={{marginBottom:8, fontSize:13, borderBottom:"1px solid #eee", paddingBottom:4}}>
              <input value={b.label} onChange={(e)=>updateLabel(i, e.target.value)} style={{width:90}} />
              <button onClick={()=> setBoxes(boxes.filter((_,j)=>j!==i))} style={{marginLeft:4}}>x</button>
            </div>
          ))}
          <div style={{marginTop:16, fontSize:12, color:"#666"}}>
            <div><span style={{color:"lime"}}>■</span> Manual</div>
            <div><span style={{color:"#f59e0b"}}>■</span> Auto (YOLO)</div>
            <div><span style={{color:"#06b6d4"}}>■</span> Tracked</div>
          </div>
        </div>
      </div>
    </div>
  );
}
import { useState, useEffect } from "react";
import { Stage, Layer, Rect, Text } from "react-konva";

const API = "http://localhost:8000";

export default function Annotator({ videoId, frames }) {
  const [idx, setIdx] = useState(0);
  const [boxes, setBoxes] = useState([]);
  const [drawing, setDrawing] = useState(null);

  const frameName = frames[idx];

  useEffect(() => {
    fetch(`${API}/annotations/${videoId}/${frameName}`)
     .then(r=>r.json()).then(setBoxes);
  }, [frameName]);

  const save = () => {
    fetch(`${API}/annotations/${videoId}/${frameName}`, {
      method: "POST",
      headers: {"Content-Type":"application/json"},
      body: JSON.stringify(boxes)
    }).then(()=> alert("Saved"));
  };

  const handleMouseDown = (e) => {
    const {x,y} = e.target.getStage().getPointerPosition();
    setDrawing({x,y,width:0,height:0,label:"object"});
  };
  const handleMouseMove = (e) => {
    if(!drawing) return;
    const {x,y} = e.target.getStage().getPointerPosition();
    setDrawing({...drawing, width: x-drawing.x, height: y-drawing.y});
  };
  const handleMouseUp = () => {
    if(drawing && Math.abs(drawing.width) > 10){
      setBoxes([...boxes, drawing]);
    }
    setDrawing(null);
  };

  return (
    <div>
      <div style={{display:"flex", gap:10, marginBottom:10}}>
        <button onClick={()=>setIdx(Math.max(0,idx-1))}>Prev</button>
        <span>{idx+1} / {frames.length} - {frameName}</span>
        <button onClick={()=>setIdx(Math.min(frames.length-1,idx+1))}>Next</button>
        <button onClick={save} style={{marginLeft:20, background:"#000", color:"#fff"}}>Save Frame</button>
      </div>

      <Stage width={1280} height={720}
        onMouseDown={handleMouseDown}
        onMouseMove={handleMouseMove}
        onMouseUp={handleMouseUp}
        style={{background:`url(${API}/frames/${frameName})`, backgroundSize:"cover"}}>
        <Layer>
          {boxes.map((b,i)=>(
            <>
              <Rect key={i} {...b} stroke="lime" strokeWidth={2} />
              <Text x={b.x} y={b.y-15} text={b.label} fill="lime" />
            </>
          ))}
          {drawing && <Rect {...drawing} stroke="red" dash={[5,5]} />}
        </Layer>
      </Stage>
    </div>
  );
}
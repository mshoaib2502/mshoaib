import { useState } from "react";
import Annotator from "./Annotator";

export default function App(){
  const [data, setData] = useState(null);
  const upload = async (e) => {
    const fd = new FormData();
    fd.append("file", e.target.files[0]);
    const res = await fetch("http://localhost:8000/upload", {method:"POST", body:fd});
    setData(await res.json());
  };
  return (
    <div style={{padding:16, fontFamily:"sans-serif"}}>
      <h2>Vision X - Robotics Annotator</h2>
      <input type="file" accept="video/*" onChange={upload} />
      {data && (
        <div style={{marginTop:12}}>
          <p>Video: {data.video_id} | Frames: {data.count}</p>
          <Annotator videoId={data.video_id} frames={data.frames} />
        </div>
      )}
    </div>
  );
}
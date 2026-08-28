from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import os, shutil, json, uuid

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])
os.makedirs("data/videos", exist_ok=True)
os.makedirs("../frontend", exist_ok=True)
app.mount("/videos", StaticFiles(directory="data/videos"), name="videos")
app.mount("/frontend", StaticFiles(directory="../frontend", html=True), name="frontend")

# Load C++ Engine
engine = None
try:
    import visionx_engine
    engine = visionx_engine.AsyncVideoEngine()
    engine.load_model("yolov8n.onnx")
    print("C++ VIDEO ENGINE LOADED - ASYNC")
except Exception as e:
    print(f"C++ Engine not found: {e} - build it with cmake")

@app.post("/upload-video")
async def upload_video(file: UploadFile = File(...)):
    path = f"data/videos/{file.filename}"
    with open(path, "wb") as out:
        shutil.copyfileobj(file.file, out)
    if engine is None:
        return {"error": "Build C++ engine first"}
    job_id = engine.submit_video(path, sample_fps=5)
    return {"job_id": job_id, "video_url": f"/videos/{file.filename}", "video_path": path}

@app.get("/video-job/{job_id}")
def job_status(job_id: str):
    return engine.get_video_status(job_id)

@app.get("/video-job/{job_id}/results")
def job_results(job_id: str):
    results = engine.get_video_results(job_id)
    # Flatten for frontend
    flat = []
    for frame_dets in results:
        for d in frame_dets:
            flat.append({"x": d.x, "y": d.y, "w": d.w, "h": d.h, "class_id": d.class_id, "class_name": d.class_name, "frame_idx": d.frame_idx, "conf": d.conf})
    return {"frames": len(results), "detections": flat, "grouped": results}
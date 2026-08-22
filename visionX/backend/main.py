from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import cv2, os, uuid, json
from pathlib import Path

app = FastAPI(title="Robotics Annotation API")

app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

BASE = Path("data")
FRAMES = BASE / "frames"
UPLOADS = BASE / "uploads"
FRAMES.mkdir(parents=True, exist_ok=True)
UPLOADS.mkdir(parents=True, exist_ok=True)

# In-memory store for MVP, move to Postgres later
DB = BASE / "annotations.json"
if not DB.exists():
    DB.write_text("{}")

def get_db():
    return json.loads(DB.read_text())
def save_db(d):
    DB.write_text(json.dumps(d, indent=2))

@app.post("/upload")
async def upload_video(file: UploadFile = File(...)):
    video_id = str(uuid.uuid4())[:8]
    video_path = UPLOADS / f"{video_id}_{file.filename}"
    with open(video_path, "wb") as f:
        f.write(await file.read())

    # OpenCV extract frames at 2 FPS
    cap = cv2.VideoCapture(str(video_path))
    fps = cap.get(cv2.CAP_PROP_FPS) or 30
    step = max(int(fps // 2), 1) # 2 fps

    frame_ids = []
    idx = 0
    saved = 0
    while True:
        ret, frame = cap.read()
        if not ret: break
        if idx % step == 0:
            # Resize for web
            frame = cv2.resize(frame, (1280, 720))
            frame_name = f"{video_id}_{saved}.jpg"
            cv2.imwrite(str(FRAMES / frame_name), frame)
            frame_ids.append(frame_name)
            saved += 1
        idx += 1
    cap.release()

    db = get_db()
    db[video_id] = {"video": file.filename, "frames": frame_ids, "annotations": {}}
    save_db(db)
    return {"video_id": video_id, "frames": frame_ids, "count": saved}

@app.get("/videos")
def list_videos():
    return get_db()

@app.get("/annotations/{video_id}/{frame_name}")
def get_annotations(video_id: str, frame_name: str):
    db = get_db()
    return db.get(video_id, {}).get("annotations", {}).get(frame_name, [])

@app.post("/annotations/{video_id}/{frame_name}")
def save_annotations(video_id: str, frame_name: str, boxes: list):
    # boxes = [{x,y,width,height,label}]
    db = get_db()
    if video_id not in db: return {"error": "video not found"}
    db[video_id]["annotations"][frame_name] = boxes
    save_db(db)
    return {"ok": True}

@app.get("/export/{video_id}")
def export_coco(video_id: str):
    db = get_db()
    project = db.get(video_id)
    if not project: return {"error": "not found"}
    coco = []
    for frame in project["frames"]:
        for b in project["annotations"].get(frame, []):
            coco.append({"image": frame, **b})
    return {"video_id": video_id, "annotations": coco}

app.mount("/frames", StaticFiles(directory=FRAMES), name="frames")
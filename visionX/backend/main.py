from fastapi import FastAPI, UploadFile, File
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
import cv2
import os
import uuid
import json
from pathlib import Path
from ultralytics import YOLO

app = FastAPI(title="Robotics Annotation API - Encord Lite")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE = Path("data")
FRAMES = BASE / "frames"
UPLOADS = BASE / "uploads"
FRAMES.mkdir(parents=True, exist_ok=True)
UPLOADS.mkdir(parents=True, exist_ok=True)

DB = BASE / "annotations.json"
if not DB.exists():
    DB.write_text("{}")

def get_db():
    try:
        return json.loads(DB.read_text())
    except:
        return {}

def save_db(d):
    DB.write_text(json.dumps(d, indent=2))

# Load YOLO once - first run will download yolov8n.pt (~6MB)
print("Loading YOLO model...")
model = YOLO("yolov8n.pt")
print("YOLO loaded")

@app.post("/upload")
async def upload_video(file: UploadFile = File(...)):
    video_id = str(uuid.uuid4())[:8]
    video_path = UPLOADS / f"{video_id}_{file.filename}"
    with open(video_path, "wb") as f:
        f.write(await file.read())

    cap = cv2.VideoCapture(str(video_path))
    fps = cap.get(cv2.CAP_PROP_FPS) or 30
    step = max(int(fps // 2), 1) # 2 FPS for annotation

    frame_ids = []
    idx = 0
    saved = 0
    while True:
        ret, frame = cap.read()
        if not ret:
            break
        if idx % step == 0:
            frame = cv2.resize(frame, (1280, 720))
            frame_name = f"{video_id}_{saved:05d}.jpg"
            cv2.imwrite(str(FRAMES / frame_name), frame)
            frame_ids.append(frame_name)
            saved += 1
        idx += 1
    cap.release()

    db = get_db()
    db[video_id] = {
        "video": file.filename,
        "frames": frame_ids,
        "annotations": {}
    }
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
    db = get_db()
    if video_id not in db:
        return {"error": "video not found"}
    db[video_id]["annotations"][frame_name] = boxes
    save_db(db)
    return {"ok": True, "count": len(boxes)}

@app.post("/auto-label/{video_id}/{frame_name}")
def auto_label(video_id: str, frame_name: str, conf: float = 0.4):
    frame_path = FRAMES / frame_name
    if not frame_path.exists():
        return {"error": "frame not found"}

    img = cv2.imread(str(frame_path))
    results = model(img, conf=conf, verbose=False)

    boxes = []
    for r in results:
        for box in r.boxes:
            x1, y1, x2, y2 = box.xyxy[0].tolist()
            cls = int(box.cls[0])
            label = model.names[cls]
            boxes.append({
                "x": float(x1),
                "y": float(y1),
                "width": float(x2 - x1),
                "height": float(y2 - y1),
                "label": label,
                "conf": float(box.conf[0]),
                "auto": True
            })

    db = get_db()
    if video_id in db:
        db[video_id]["annotations"][frame_name] = boxes
        save_db(db)
    return boxes

@app.post("/track/{video_id}")
def track_boxes(video_id: str, payload: dict):
    from_frame = payload["from_frame"]
    to_frame = payload["to_frame"]
    boxes = payload["boxes"]

    path1 = FRAMES / from_frame
    path2 = FRAMES / to_frame
    if not path1.exists() or not path2.exists():
        return {"error": "frame not found"}

    img1 = cv2.imread(str(path1))
    img2 = cv2.imread(str(path2))

    tracked = []
    for b in boxes:
        try:
            tracker = cv2.TrackerCSRT_create()
            bbox = (int(b["x"]), int(b["y"]), int(b["width"]), int(b["height"]))
            tracker.init(img1, bbox)
            ok, new_box = tracker.update(img2)
            if ok:
                x, y, w, h = new_box
                tracked.append({**b, "x": float(x), "y": float(y), "width": float(w), "height": float(h), "tracked": True})
            else:
                tracked.append(b)
        except Exception as e:
            tracked.append(b)

    db = get_db()
    if video_id in db:
        db[video_id]["annotations"][to_frame] = tracked
        save_db(db)
    return tracked

@app.get("/export/{video_id}")
def export_coco(video_id: str):
    db = get_db()
    project = db.get(video_id)
    if not project:
        return {"error": "not found"}
    coco = []
    for frame in project["frames"]:
        for b in project["annotations"].get(frame, []):
            coco.append({"image": frame, **b})
    return {"video_id": video_id, "video_name": project["video"], "total_annotations": len(coco), "annotations": coco}

# Serve frames as static
app.mount("/frames", StaticFiles(directory=FRAMES), name="frames")
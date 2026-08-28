bash# 1. Get ONNX model
pip install ultralytics
yolo export model=yolov8n.pt format=onnx

# 2. Build C++ engine
mkdir backend/build && cd backend/build
cmake.. && make -j
cp visionx_engine*.so../
cd../..

# 3. Run
cd backend
pip install -r requirements.txt
uvicorn main:app --reload --port 8000

# 4. Open http://localhost:8000/frontend/
# Upload mp4 -> C++ processes in background -> watch progress bar -> video with boxes
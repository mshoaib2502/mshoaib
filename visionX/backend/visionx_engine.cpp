#include <pybind11/pybind11.h>
#include <pybind11/stl.h>
#include <opencv2/opencv.hpp>
#include <opencv2/dnn.hpp>
#include <queue>
#include <thread>
#include <mutex>
#include <atomic>
#include <future>
#include <unordered_map>
namespace py = pybind11;
using namespace pybind11::literals;

struct Detection
{
    int x, y, w, h;
    int class_id;
    float conf;
    std::string class_name;
    int frame_idx;
};
struct VideoJob
{
    std::string job_id;
    std::string video_path;
    int total_frames;
    int processed;
    std::vector<std::vector<Detection>> all_detections;
    bool done = false;
};

class AsyncVideoEngine
{
private:
    cv::dnn::Net net;
    std::unordered_map<std::string, VideoJob> jobs;
    std::mutex job_mtx;
    std::vector<std::string> names = {"person", "bicycle", "car", "motorcycle", "airplane", "bus", "train", "truck", "boat", "traffic light", "fire hydrant", "stop sign", "parking meter", "bench", "bird", "cat", "dog", "horse", "sheep", "cow", "elephant", "bear", "zebra", "giraffe", "backpack", "umbrella", "handbag", "tie", "suitcase", "frisbee", "skis", "snowboard", "sports ball", "kite", "baseball bat", "baseball glove", "skateboard", "surfboard", "tennis racket", "bottle", "wine glass", "cup", "fork", "knife", "spoon", "bowl", "banana", "apple", "sandwich", "orange", "broccoli", "carrot", "hot dog", "pizza", "donut", "cake", "chair", "couch", "potted plant", "bed", "dining table", "toilet", "tv", "laptop", "mouse", "remote", "keyboard", "cell phone", "microwave", "oven", "toaster", "sink", "refrigerator", "book", "clock", "vase", "scissors", "teddy bear", "hair drier", "toothbrush"};

    std::vector<Detection> run_yolo(cv::Mat &img, int frame_idx)
    {
        int W = img.cols, H = img.rows;
        cv::Mat blob;
        cv::dnn::blobFromImage(img, blob, 1.0 / 255.0, cv::Size(640, 640), cv::Scalar(), true, false);
        net.setInput(blob);
        std::vector<cv::Mat> outs;
        net.forward(outs, net.getUnconnectedOutLayersNames());
        cv::Mat ot;
        cv::transpose(outs[0].reshape(1, 84), ot);
        std::vector<cv::Rect> boxes;
        std::vector<float> confs;
        std::vector<int> cids;
        for (int i = 0; i < ot.rows; i++)
        {
            float *d = (float *)ot.row(i).data;
            cv::Mat s(1, 80, CV_32F, d + 4);
            cv::Point id;
            double sc;
            cv::minMaxLoc(s, 0, &sc, 0, &id);
            if (sc > 0.5)
            {
                boxes.push_back(cv::Rect(int((d[0] - d[2] / 2) * W / 640), int((d[1] - d[3] / 2) * H / 640), int(d[2] * W / 640), int(d[3] * H / 640)));
                confs.push_back(sc);
                cids.push_back(id.x);
            }
        }
        std::vector<int> idx;
        cv::dnn::NMSBoxes(boxes, confs, 0.5, 0.45, idx);
        std::vector<Detection> res;
        for (int i : idx)
            res.push_back({boxes[i].x, boxes[i].y, boxes[i].width, boxes[i].height, cids[i], confs[i], names[cids[i]], frame_idx});
        return res;
    }

public:
    bool load_model(const std::string &path)
    {
        net = cv::dnn::readNetFromONNX(path);
        return true;
    }

    std::string submit_video(const std::string &video_path, int sample_fps = 5)
    {
        std::string job_id = "job_" + std::to_string(rand() % 100000);
        {
            std::lock_guard<std::mutex> lock(job_mtx);
            jobs[job_id] = {job_id, video_path, 0, 0, {}, false};
        }
        std::thread([this, job_id, video_path, sample_fps]
                    {
            cv::VideoCapture cap(video_path); double fps=cap.get(cv::CAP_PROP_FPS); int total=cap.get(cv::CAP_PROP_FRAME_COUNT);
            int step=std::max(1,(int)(fps/sample_fps)); int out_total=total/step;
            { std::lock_guard<std::mutex> lock(job_mtx); jobs[job_id].total_frames=out_total; jobs[job_id].all_detections.resize(out_total); }
            cv::Mat frame; int f_idx=0,out_idx=0;
            while(cap.read(frame)){
                if(f_idx%step==0){
                    auto dets=run_yolo(frame,f_idx);
                    std::lock_guard<std::mutex> lock(job_mtx);
                    if(jobs.count(job_id)){ jobs[job_id].all_detections[out_idx]=dets; jobs[job_id].processed++; }
                    out_idx++;
                }
                f_idx++;
            }
            std::lock_guard<std::mutex> lock(job_mtx); if(jobs.count(job_id)) jobs[job_id].done=true; })
            .detach();
        return job_id;
    }

    py::dict get_video_status(const std::string &job_id)
    {
        std::lock_guard<std::mutex> lock(job_mtx);
        if (!jobs.count(job_id))
        {
            py::dict d;
            d["error"] = "not found";
            return d;
        }
        auto &j = jobs[job_id];
        py::dict d;
        d["job_id"] = j.job_id;
        d["total"] = j.total_frames;
        d["processed"] = j.processed;
        d["progress"] = j.total_frames > 0 ? (float)j.processed / j.total_frames : 0;
        d["done"] = j.done;
        return d;
    }

    std::vector<std::vector<Detection>> get_video_results(const std::string &job_id)
    {
        std::lock_guard<std::mutex> lock(job_mtx);
        if (!jobs.count(job_id))
            return {};
        return jobs[job_id].all_detections;
    }
};

PYBIND11_MODULE(visionx_engine, m)
{
    py::class_<Detection>(m, "Detection").def_readonly("x", &Detection::x).def_readonly("y", &Detection::y).def_readonly("w", &Detection::w).def_readonly("h", &Detection::h).def_readonly("class_id", &Detection::class_id).def_readonly("conf", &Detection::conf).def_readonly("class_name", &Detection::class_name).def_readonly("frame_idx", &Detection::frame_idx);
    py::class_<AsyncVideoEngine>(m, "AsyncVideoEngine").def(py::init<>()).def("load_model", &AsyncVideoEngine::load_model).def("submit_video", &AsyncVideoEngine::submit_video).def("get_video_status", &AsyncVideoEngine::get_video_status).def("get_video_results", &AsyncVideoEngine::get_video_results);
}
import { useState, useRef, useEffect } from 'react'

export default function App() {
  const [jobId, setJobId] = useState(null)
  const [progress, setProgress] = useState(0)
  const [status, setStatus] = useState('Idle')
  const [grouped, setGrouped] = useState([]) // [[dets frame0], [frame1]...]
  const [videoUrl, setVideoUrl] = useState(null)
  const [isPlaying, setIsPlaying] = useState(false)

  const videoRef = useRef(null)
  const canvasRef = useRef(null)
  const fileRef = useRef(null)

  // 1. Upload Video -> C++ async job
  const uploadVideo = async () => {
    const file = fileRef.current.files[0]
    if (!file) return
    const fd = new FormData()
    fd.append('file', file)
    setStatus('Uploading...')
    const res = await fetch('/upload-video', { method: 'POST', body: fd })
    const data = await res.json()
    setJobId(data.job_id)
    setVideoUrl(data.video_url)
    setStatus('C++ processing in background...')
  }

  // 2. Poll C++ job progress
  useEffect(() => {
    if (!jobId) return
    const poll = async () => {
      const r = await fetch(`/video-job/${jobId}`)
      const s = await r.json()
      setProgress(s.progress || 0)
      setStatus(s.done? `DONE - ${s.processed}/${s.total} frames` : `Processing: ${s.processed}/${s.total} frames in C++ thread pool`)
      if (!s.done) {
        setTimeout(poll, 1000)
      } else {
        // Load results
        const rr = await fetch(`/video-job/${jobId}/results`)
        const dd = await rr.json()
        setGrouped(dd.grouped)
      }
    }
    poll()
  }, [jobId])

  // 3. Draw boxes synced with video time
  useEffect(() => {
    const video = videoRef.current
    const canvas = canvasRef.current
    if (!video ||!canvas) return
    const ctx = canvas.getContext('2d')

    const onTimeUpdate = () => {
      // We sampled at 5fps in C++, so map time -> sampled frame index
      const SAMPLE_FPS = 5
      const frameIdx = Math.floor(video.currentTime * SAMPLE_FPS)
      const dets = grouped[frameIdx] || []

      ctx.clearRect(0, 0, canvas.width, canvas.height)
      ctx.drawImage(video, 0, 0, canvas.width, canvas.height)

      dets.forEach(d => {
        const scaleX = canvas.width / video.videoWidth
        const scaleY = canvas.height / video.videoHeight
        ctx.strokeStyle = '#00ff00'
        ctx.lineWidth = 2
        ctx.strokeRect(d.x * scaleX, d.y * scaleY, d.w * scaleX, d.h * scaleY)
        ctx.fillStyle = '#00ff00'
        ctx.font = '14px monospace'
        ctx.fillText(`${d.class_name} ${d.conf.toFixed(2)}`, d.x * scaleX, d.y * scaleY - 4)
      })
    }

    video.addEventListener('timeupdate', onTimeUpdate)
    return () => video.removeEventListener('timeupdate', onTimeUpdate)
  }, [grouped])

  return (
    <div style={{ display: 'flex', height: '100vh' }}>
      {/* Sidebar */}
      <div style={{ width: 320, background: '#161616', padding: 16, borderRight: '1px solid #222' }}>
        <h2 style={{ margin: '0 0 12px', color: '#00ff00' }}>visionX VIDEO</h2>

        <input ref={fileRef} type="file" accept="video/*" style={{ width: '100%', marginBottom: 12 }} />
        <button onClick={uploadVideo} style={{ width: '100%', padding: 12, background: '#00ff00', color: '#000', border: 0, fontWeight: 800 }}>UPLOAD & AUTO-LABEL (C++)</button>

        {jobId && (
          <div style={{ marginTop: 16 }}>
            <div style={{ fontSize: 12, opacity: 0.6 }}>JOB ID: {jobId}</div>
            <div style={{ height: 10, background: '#333', marginTop: 8, borderRadius: 4 }}>
              <div style={{ height: '100%', width: `${progress * 100}%`, background: '#00ff00', transition: 'width 0.3s' }} />
            </div>
            <div style={{ marginTop: 8, fontSize: 13 }}>{status}</div>
          </div>
        )}

        <div style={{ marginTop: 20 }}>
          <button onClick={() => {
            if (videoRef.current.paused) { videoRef.current.play(); setIsPlaying(true) }
            else { videoRef.current.pause(); setIsPlaying(false) }
          }} style={{ width: '100%', padding: 10, background: '#222', color: '#fff', border: '1px solid #333' }}>
            {isPlaying? 'Pause' : 'Play'}
          </button>
        </div>

        <div style={{ marginTop: 20, fontSize: 12, opacity: 0.5, lineHeight: 1.5 }}>
          <b>How it works:</b><br/>
          1. Python receives mp4<br/>
          2. Calls C++ AsyncVideoEngine.submit_video()<br/>
          3. C++ thread pool samples 5 FPS<br/>
          4. YOLO in C++ (OpenCV DNN)<br/>
          5. Frontend polls /video-job/id<br/>
          6. Canvas draws boxes synced to video time
        </div>
      </div>

      {/* Main Video + Canvas */}
      <div style={{ flex: 1, background: '#000', display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', position: 'relative' }}>
        {/* Hidden video element for decoding, visible canvas for drawing */}
        <video ref={videoRef} src={videoUrl} crossOrigin="anonymous" style={{ display: 'none' }} />
        <canvas ref={canvasRef} width={960} height={540} style={{ border: '1px solid #222', background: '#000', maxWidth: '90%', aspectRatio: '16/9' }} />
        {!videoUrl && <div style={{ opacity: 0.3, marginTop: 20 }}>Upload a video to start</div>}
        {videoUrl && (
          <input type="range" min={0} max={100} defaultValue={0} onChange={e => {
            const v = videoRef.current
            v.currentTime = (e.target.value / 100) * v.duration
          }} style={{ width: 960, maxWidth: '90%', marginTop: 12 }} />
        )}
      </div>
    </div>
  )
}
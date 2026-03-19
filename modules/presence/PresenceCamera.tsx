
import React, { useRef, useState, useEffect } from 'react';
import { Camera, RefreshCw, ShieldCheck, ArrowRight, ArrowLeft, Loader2, X } from 'lucide-react';
import { FaceLandmarker, FilesetResolver } from "@mediapipe/tasks-vision";

interface PresenceCameraProps {
  onCapture: (blob: Blob) => void;
  onClose: () => void;
  isProcessing?: boolean;
}

const PresenceCamera: React.FC<PresenceCameraProps> = ({ onCapture, onClose, isProcessing }) => {
  const videoRef = useRef<HTMLVideoElement>(null);
  const landmarkerRef = useRef<any>(null);
  const requestRef = useRef<number>(null);
  
  const [stream, setStream] = useState<MediaStream | null>(null);
  const [step, setStep] = useState<'RIGHT' | 'LEFT' | 'READY'>('RIGHT');
  const [isAiLoaded, setIsAiLoaded] = useState(false);
  const isComponentMounted = useRef(true);
  const lastVideoTimeRef = useRef(-1);

  useEffect(() => {
    isComponentMounted.current = true;
    initializeAi();

    return () => {
      isComponentMounted.current = false;
      if (requestRef.current) cancelAnimationFrame(requestRef.current);
      stopCamera();
    };
  }, []);

  useEffect(() => {
    if (isAiLoaded) {
      startCamera();
    }
  }, [isAiLoaded]);

  const initializeAi = async () => {
    try {
      const filesetResolver = await FilesetResolver.forVisionTasks(
        "https://cdn.jsdelivr.net/npm/@mediapipe/tasks-vision@0.10.3/wasm"
      );
      const landmarker = await FaceLandmarker.createFromOptions(filesetResolver, {
        baseOptions: {
          modelAssetPath: `https://storage.googleapis.com/mediapipe-models/face_landmarker/face_landmarker/float16/1/face_landmarker.task`,
          delegate: "GPU"
        },
        outputFaceBlendshapes: true,
        runningMode: "VIDEO",
        numFaces: 1
      });
      
      if (isComponentMounted.current) {
        landmarkerRef.current = landmarker;
        setIsAiLoaded(true);
      }
    } catch (err) {
      console.error("AI Init Error:", err);
    }
  };

  const startCamera = async () => {
    try {
      // Menggunakan resolusi 720p ideal untuk menghindari cropping digital (zoom) di perangkat mobile
      const s = await navigator.mediaDevices.getUserMedia({ 
        video: { 
          facingMode: 'user', 
          width: { ideal: 720 }, 
          height: { ideal: 1280 } 
        } 
      });
      
      if (isComponentMounted.current) {
        setStream(s);
        if (videoRef.current) {
          videoRef.current.srcObject = s;
          videoRef.current.onloadedmetadata = () => {
            if (videoRef.current) videoRef.current.play();
            predictLoop();
          };
        }
      }
    } catch (err) {
      console.error("Camera Error:", err);
    }
  };

  const stopCamera = () => {
    if (stream) {
      stream.getTracks().forEach(track => track.stop());
    }
  };

  const predictLoop = () => {
    // PAUSE loop jika sedang memproses capture untuk membebaskan resource hardware
    if (!isComponentMounted.current || !videoRef.current || !landmarkerRef.current || isProcessing) {
      if (isComponentMounted.current && !isProcessing) {
        requestRef.current = requestAnimationFrame(predictLoop);
      }
      return;
    }

    const video = videoRef.current;
    if (video.readyState >= 2 && video.currentTime !== lastVideoTimeRef.current) {
      lastVideoTimeRef.current = video.currentTime;
      
      try {
        const results = landmarkerRef.current.detectForVideo(video, performance.now());

        if (results.faceLandmarks && results.faceLandmarks.length > 0) {
          const landmarks = results.faceLandmarks[0];
          const nose = landmarks[4];
          const rightEdge = landmarks[234];
          const leftEdge = landmarks[454];

          const faceWidth = Math.abs(leftEdge.x - rightEdge.x);
          const noseRelativeX = (nose.x - Math.min(rightEdge.x, leftEdge.x)) / faceWidth;

          setStep(prev => {
            if (prev === 'RIGHT' && noseRelativeX < 0.35) return 'LEFT';
            if (prev === 'LEFT' && noseRelativeX > 0.65) return 'READY';
            return prev;
          });
        }
      } catch (e) {
        console.error("In-loop detection error:", e);
      }
    }

    if (isComponentMounted.current) {
      requestRef.current = requestAnimationFrame(predictLoop);
    }
  };

  const handleCapture = () => {
    const video = videoRef.current;
    if (!video || video.readyState < 2 || isProcessing) return;

    const canvas = document.createElement('canvas');
    canvas.width = video.videoWidth;
    canvas.height = video.videoHeight;
    const ctx = canvas.getContext('2d');
    
    if (ctx) {
      ctx.translate(canvas.width, 0);
      ctx.scale(-1, 1);
      ctx.drawImage(video, 0, 0, canvas.width, canvas.height);
      
      canvas.toBlob((blob) => {
        if (blob) onCapture(blob);
      }, 'image/jpeg', 0.85);
    }
  };

  return (
    <div className="w-full max-w-md mx-auto animate-in fade-in zoom-in duration-500">
      {!isAiLoaded ? (
        <div className="aspect-[9/16] w-full bg-slate-950 rounded-3xl flex flex-col items-center justify-center text-white/50 gap-4 border border-white/5">
          <Loader2 className="animate-spin text-[#00FFE4]" size={40} />
          <p className="text-[10px] font-bold uppercase tracking-widest text-center px-6">Inisialisasi Keamanan AI...</p>
        </div>
      ) : (
        <div className="relative w-full aspect-[9/16] bg-black rounded-3xl overflow-hidden shadow-2xl ring-1 ring-white/10">
          <video 
            ref={videoRef} 
            autoPlay 
            playsInline 
            muted 
            className="absolute inset-0 w-full h-full object-cover scale-x-[-1]"
          />
          
          <button 
            onClick={onClose}
            disabled={isProcessing}
            className="absolute top-5 right-5 z-[60] p-2.5 bg-black/40 backdrop-blur-md rounded-full text-white/80 hover:text-white disabled:opacity-50 transition-all border border-white/10"
          >
            <X size={20} />
          </button>

          {/* Area Deteksi & UI */}
          <div className="relative h-full w-full flex flex-col items-center justify-between p-6 pointer-events-none z-40">
            <div className="mt-8 bg-black/60 backdrop-blur-xl px-6 py-4 rounded-2xl border border-white/10 text-center animate-in fade-in zoom-in duration-500 w-fit">
              {step === 'RIGHT' && (
                <div className="flex flex-col items-center gap-2">
                  <div className="w-10 h-10 bg-[#00FFE4]/20 rounded-full flex items-center justify-center">
                    <ArrowRight className="text-[#00FFE4] animate-bounce" size={24} />
                  </div>
                  <p className="text-white text-[10px] font-bold uppercase tracking-widest">Tengok ke Kanan</p>
                </div>
              )}
              {step === 'LEFT' && (
                <div className="flex flex-col items-center gap-2">
                  <div className="w-10 h-10 bg-[#00FFE4]/20 rounded-full flex items-center justify-center">
                    <ArrowLeft className="text-[#00FFE4] animate-bounce" size={24} />
                  </div>
                  <p className="text-white text-[10px] font-bold uppercase tracking-widest">Tengok ke Kiri</p>
                </div>
              )}
              {step === 'READY' && (
                <div className="flex flex-col items-center gap-2">
                  <div className="w-10 h-10 bg-emerald-500/20 rounded-full flex items-center justify-center">
                    <ShieldCheck className="text-emerald-400" size={26} />
                  </div>
                  <p className="text-emerald-400 text-[10px] font-bold uppercase tracking-widest">Identitas Valid</p>
                </div>
              )}
            </div>

            <div className={`w-56 h-72 border-2 border-dashed rounded-[100px] transition-all duration-700 ${
              step === 'READY' 
              ? 'border-emerald-500 bg-emerald-500/5 scale-105 shadow-[0_0_50px_rgba(16,185,129,0.2)]' 
              : 'border-white/20 bg-white/5'
            }`}></div>

            <div className="w-full space-y-5 mb-6">
              <div className="px-4">
                <div className="h-1 w-full bg-white/10 rounded-full overflow-hidden">
                  <div 
                    className={`h-full transition-all duration-700 ease-out shadow-[0_0_15px_rgba(0,255,228,0.5)] bg-[#00FFE4]`} 
                    style={{ width: step === 'RIGHT' ? '33%' : step === 'LEFT' ? '66%' : '100%' }}
                  />
                </div>
              </div>

              <div className="flex justify-center gap-3 pointer-events-auto">
                <button 
                  onClick={() => { 
                    setStep('RIGHT');
                    lastVideoTimeRef.current = -1;
                  }}
                  disabled={isProcessing}
                  className="p-4 text-white/50 hover:text-white bg-white/5 backdrop-blur-lg rounded-full border border-white/10 transition-all hover:bg-white/10 disabled:opacity-30"
                >
                  <RefreshCw size={24} />
                </button>
                <button 
                  onClick={handleCapture}
                  disabled={step !== 'READY' || isProcessing}
                  className={`flex items-center gap-2 px-10 py-4 rounded-full font-extrabold uppercase text-[10px] tracking-[0.2em] shadow-2xl transition-all ${
                    step === 'READY'
                    ? 'bg-[#00FFE4] text-[#006E62] hover:scale-105 active:scale-95 shadow-[#00FFE4]/20' 
                    : 'bg-white/5 text-white/20 cursor-not-allowed border border-white/5'
                  }`}
                >
                  <Camera size={20} />
                  {isProcessing ? 'PROSES...' : 'AMBIL FOTO'}
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default PresenceCamera;

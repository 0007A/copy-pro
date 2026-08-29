import os
import sys
import time
import json
import numpy as np
import speech_recognition as sr
import wave

PROFILE_PATH = r"g:\My Drive\All project\copy Pro\speaker_profile.json"
SAMPLE_WAV_PATH = r"g:\My Drive\All project\copy Pro\my_voice_sample.wav"

def extract_voice_features(raw_audio_bytes, sample_rate=16000):
    if not raw_audio_bytes or len(raw_audio_bytes) < 512:
        return None
    samples = np.frombuffer(raw_audio_bytes, dtype=np.int16).astype(np.float32)
    if len(samples) < 512:
        return None
    samples = samples - np.mean(samples)
    rms = float(np.sqrt(np.mean(samples**2)))
    zcr = float(np.mean(np.abs(np.diff(np.sign(samples)))) / 2.0)
    fft_vals = np.abs(np.fft.rfft(samples * np.hanning(len(samples))))
    freqs = np.fft.rfftfreq(len(samples), 1.0 / sample_rate)
    vocal_mask = (freqs >= 80) & (freqs <= 3500)
    vocal_fft = fft_vals[vocal_mask]
    vocal_freqs = freqs[vocal_mask]
    if len(vocal_fft) == 0 or np.sum(vocal_fft) == 0:
        return None
    spectral_centroid = float(np.sum(vocal_freqs * vocal_fft) / np.sum(vocal_fft))
    dominant_pitch = float(vocal_freqs[np.argmax(vocal_fft)])
    band_low = float(np.sum(fft_vals[(freqs >= 80) & (freqs < 500)]))
    band_mid = float(np.sum(fft_vals[(freqs >= 500) & (freqs < 1500)]))
    band_high = float(np.sum(fft_vals[(freqs >= 1500) & (freqs <= 3500)]))
    total_band = band_low + band_mid + band_high + 1e-6
    return {
        "pitch": dominant_pitch,
        "centroid": spectral_centroid,
        "zcr": zcr,
        "rms_volume": rms,
        "band_low_ratio": float(band_low / total_band),
        "band_mid_ratio": float(band_mid / total_band),
        "band_high_ratio": float(band_high / total_band)
    }

def record_sample():
    print("=" * 65)
    print("       🎙️ COPY PRO - VOICE CALIBRATION & SAMPLE RECORDER       ")
    print("=" * 65)
    print("\n[1/3] Initializing Microphone...")
    
    r = sr.Recognizer()
    r.dynamic_energy_threshold = True
    
    with sr.Microphone() as source:
        print("[2/3] Measuring background ambient noise (Please stay quiet for 1 sec)...")
        r.adjust_for_ambient_noise(source, duration=1.0)
        ambient_energy = r.energy_threshold
        print(f"      Ambient Noise Floor: {ambient_energy:.1f}")
        
        print("\n" + "#" * 65)
        print(">>> 🎙️ RECORDING STARTED! PLEASE SPEAK NOW (4 SECONDS) <<<")
        print(">>> Say: 'Option A, Option B, Next Question, Important Details' <<<")
        print("#" * 65 + "\n")
        
        # Countdown 4 seconds
        audio = r.record(source, duration=4.0)
        print("\n[3/3] Audio recording complete! Analyzing voiceprint...")

    raw_bytes = audio.get_raw_data()
    sample_rate = audio.sample_rate
    
    # Save WAV file
    with wave.open(SAMPLE_WAV_PATH, "wb") as wf:
        wf.setnchannels(1)
        wf.setsampwidth(audio.sample_width)
        wf.setframerate(sample_rate)
        wf.writeframes(raw_bytes)
    print(f"      Audio sample saved to: {SAMPLE_WAV_PATH}")
    
    feats = extract_voice_features(raw_bytes, sample_rate)
    if feats:
        feats["calibrated_energy_threshold"] = max(150, float(ambient_energy * 1.5))
        with open(PROFILE_PATH, "w", encoding="utf-8") as f:
            json.dump(feats, f, indent=2)
        print("\n" + "=" * 65)
        print("  ✅ SUCCESS! YOUR VOICE HAS BEEN CALIBRATED SUCCESSFULLY!")
        print("=" * 65)
        print(f"  • Dominant Pitch: {feats['pitch']:.1f} Hz")
        print(f"  • Spectral Centroid: {feats['centroid']:.1f} Hz")
        print(f"  • Speaking RMS Volume: {feats['rms_volume']:.1f}")
        print(f"  • Recommended Energy Threshold: {feats['calibrated_energy_threshold']:.1f}")
        print(f"  • Profile saved to: {PROFILE_PATH}")
        print("=" * 65)
    else:
        print("\n⚠️ Warning: Audio was very quiet or no speech was detected.")
        print("Please check your microphone volume in Windows settings and try again.")

if __name__ == "__main__":
    record_sample()

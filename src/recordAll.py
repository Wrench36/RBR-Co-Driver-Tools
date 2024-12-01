"""
    Holy oh shit this actually works.
    Don't record this next to your furnace.
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sounddevice as sd
import numpy as np
from scipy.signal import convolve
from scipy.io import wavfile
import soundfile as sf

def record_audio(duration, fs):
    print("Recording...")
    audio = sd.rec(int(duration * fs), samplerate=fs, channels=1, dtype='float32')
    sd.wait()
    print("Recording finished.")
    audio = audio.flatten()
    audio = audio / np.max(np.abs(audio))  # Normalize
    return audio

def apply_reverb(audio, ir_file):
    # Load impulse response
    _, ir = wavfile.read(ir_file)  # Extract the audio data
    ir = np.array(ir, dtype=np.float32)  # Convert to NumPy array
    ir = ir / np.max(np.abs(ir))

    # Apply convolution reverb
    reverb_audio = convolve(audio, ir, mode='full')
    reverb_audio = reverb_audio / np.max(np.abs(reverb_audio))
    return reverb_audio

def detect_signal_chunk(audio_chunk, noise_rms, noise_std, threshold_factor):
    """
    Determines if a chunk contains signal or noise based on mean and standard deviation.
    - `audio_chunk`: A numpy array representing a chunk of audio data.
    - `noise_mean`: The mean amplitude of the background noise.
    - `noise_std`: The standard deviation of the background noise.
    - `threshold_factor`: The factor to multiply the noise std to decide if a chunk is signal.
    Returns True if the chunk is signal, False if it is noise.
    """
    chunk_mean = np.mean(np.abs(audio_chunk))
    chunk_max = np.max(np.abs(audio_chunk))
    chunk_std = np.std(np.abs(audio_chunk))


    # Consider the chunk as signal if the mean is significantly higher than the noise level
    if chunk_max > noise_rms - (noise_std*threshold_factor):
        return True
    return False

def auto_trim_by_chunks(audio_np, fs, chunk_size_ms, threshold_factor):
    """
    Automatically trims the audio based on chunk mean and standard deviation to detect signal.
    - `audio_np`: The audio signal as a 1D numpy array.
    - `fs`: The sampling frequency of the audio.
    - `chunk_size_ms`: The size of each chunk in milliseconds.
    - `threshold_factor`: The factor to multiply the noise std to decide if a chunk is signal.
    """
    chunk_size_samples = int(fs * chunk_size_ms / 1000)
    num_chunks = len(audio_np) // chunk_size_samples

    # Calculate the mean and std of the entire audio as the noise reference
    noise_mean = np.mean(np.abs(audio_np))
    noise_rms  = np.sqrt(np.mean(audio_np**2))
    noise_max = np.max(np.abs(audio_np))
    noise_std = np.std(np.abs(audio_np))
    """print ("mean: " + str(noise_mean))
    print ("rms: " + str(noise_rms))
    print ("max: " + str(noise_max))
    print ("std: " + str(noise_std))
    print ("Compare Thresh: " + str(noise_rms - (noise_std*threshold_factor)))"""


        # Plot the audio chunk to visualize it
    """plt.figure(figsize=(10, 4))
    plt.plot(audio_np)
    plt.title("Audio Chunk")
    plt.xlabel("Sample Index")
    plt.ylabel("Amplitude")
    plt.ylim(0, 1)
    plt.grid(True)
    plt.show()  # Show the plot"""


    signal_chunks = []

    for i in range(num_chunks):
        start_idx = i * chunk_size_samples
        end_idx = (i + 1) * chunk_size_samples
        chunk = audio_np[start_idx:end_idx]

        if detect_signal_chunk(chunk, noise_rms, noise_std, threshold_factor):
            signal_chunks.append(chunk)

    # Combine the detected signal chunks
    signal_audio = np.concatenate(signal_chunks)

    return signal_audio

def save_audio(audio, fs, filename):
    #wavfile.write(filename, fs, (audio * 32767).astype(np.int16))
    sf.write(filename, audio, fs)
    print(f"Audio saved as {filename}")


def main():
    duration = 2  # seconds
    fs = 44100  # Hz
    ir_file = 'D:/Richard Burns Rally/Apps/new pacenote mod/Pacenote/Doc/record/walkietalkie2.wav'  # Replace with your IR file
    parent_dir = os.path.dirname(os.path.abspath(__file__)) 
    docs_dir = os.path.abspath(os.path.join(parent_dir, "../"))
    rec_path = os.path.abspath(os.path.join(parent_dir, "../rec"))
    if not os.path.exists(rec_path):
        os.makedirs(rec_path)
    file = filedialog.askopenfilename(initialdir=docs_dir,title=f"Select Script File",filetypes=[(".txt", "*")])
    with open(file, 'r') as f:
        lines = f.readlines()
        for line in lines:
            print(line)
            line = line.replace("\n", "")
            output_file = f"{rec_path}/{line}.ogg"
            audio = record_audio(duration, fs)
            reverb_audio = apply_reverb(audio, ir_file)
            trimmed_audio = auto_trim_by_chunks(reverb_audio, fs, 50, 0.85)
            save_audio(trimmed_audio, fs, output_file)
if __name__ == "__main__":
    main()

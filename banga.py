import tkinter as tk
from tkinter import ttk
from pathlib import Path
import pyaudio
import wave
import pandas as pd
from datetime import datetime
import threading
import time
import os
import numpy as np

BASE_DIR = Path(__file__).parent

class BangaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Banga Audio Recorder")
        self.phrase_index = 0
        self.synonym_index = 1
        self.synonym_data = {}
        self.phrases = []
        self.recording = False
        self.current_audio_path = None
        self.playing = False
        self.current_frame = 0
        self.total_frames = 0
        self.play_thread = None
        self.playback_timer_thread = None
        self.wf = None
        self.stream = None
        self.pyaudio_instance = None
        self.recording_thread = None
        self.timer_thread = None
        self.synonym_menu = None
        self.resources_cleaned = False

        self.load_phrases()
        for phrase in self.phrases:
            if phrase not in self.synonym_data:
                self.synonym_data[phrase] = []
        self.load_existing_metadata()
        for phrase in self.phrases:
            if not self.synonym_data[phrase]:
                self.synonym_data[phrase].append({
                    "synonym": 1,
                    "audio_path": "",
                    "transcription": ""
                })

        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.selection_frame = ttk.LabelFrame(self.main_frame, text="Select Phrase and Synonym", padding="5")
        self.selection_frame.grid(row=0, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.phrase_label = ttk.Label(self.selection_frame, text="Phrase:")
        self.phrase_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.phrase_var = tk.StringVar(value=self.phrases[0] if self.phrases else "No Phrases Loaded")
        self.phrase_menu = ttk.OptionMenu(self.selection_frame, self.phrase_var, self.phrases[0] if self.phrases else "No Phrases Loaded", *self.phrases, command=self.update_phrase)
        self.phrase_menu.grid(row=0, column=1, padx=5, pady=5)

        self.synonym_label = ttk.Label(self.selection_frame, text="Synonym:")
        self.synonym_label.grid(row=1, column=0, padx=5, pady=5, sticky=tk.W)
        self.synonym_var = tk.StringVar(value=f"Synonym {self.synonym_index}")
        self.update_synonym_menu()

        self.add_synonym_button = ttk.Button(self.selection_frame, text="Add Synonym", command=self.add_synonym)
        self.add_synonym_button.grid(row=1, column=2, padx=5, pady=5)

        ttk.Separator(self.main_frame, orient=tk.HORIZONTAL).grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.status_frame = ttk.LabelFrame(self.main_frame, text="Status", padding="5")
        self.status_frame.grid(row=2, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.status_label = ttk.Label(self.status_frame, text="Ready to record")
        self.status_label.grid(row=0, column=0, columnspan=2, pady=5)

        self.timer_label = ttk.Label(self.status_frame, text="00:00")
        self.timer_label.grid(row=1, column=0, columnspan=2, pady=5)

        ttk.Separator(self.main_frame, orient=tk.HORIZONTAL).grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.recording_frame = ttk.LabelFrame(self.main_frame, text="Recording Controls", padding="5")
        self.recording_frame.grid(row=4, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.record_button = ttk.Button(self.recording_frame, text="Start Recording", command=self.start_recording)
        self.record_button.grid(row=0, column=0, padx=5, pady=5)

        self.stop_button = ttk.Button(self.recording_frame, text="Stop Recording", command=self.stop_recording, state=tk.DISABLED)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5)

        self.playback_frame = ttk.LabelFrame(self.main_frame, text="Playback Controls", padding="5")
        self.playback_frame.grid(row=5, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.play_button = ttk.Button(self.playback_frame, text="Play", command=self.toggle_play_pause, state=tk.DISABLED)
        self.play_button.grid(row=0, column=0, padx=5, pady=5)

        self.new_recording_button = ttk.Button(self.playback_frame, text="Re-Record", command=self.reset_to_recording_mode)

        self.progress_label = ttk.Label(self.playback_frame, text="Playback Progress:")
        self.progress_label.grid(row=1, column=0, pady=5, sticky=tk.W)
        self.progress_bar = ttk.Progressbar(self.playback_frame, length=200, mode='determinate')
        self.progress_bar.grid(row=1, column=1, pady=5)

        ttk.Separator(self.main_frame, orient=tk.HORIZONTAL).grid(row=6, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=5)

        self.transcription_frame = ttk.LabelFrame(self.main_frame, text="Transcription", padding="5")
        self.transcription_frame.grid(row=7, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.transcription_label = ttk.Label(self.transcription_frame, text="Transcription:")
        self.transcription_label.grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.transcription_entry = ttk.Entry(self.transcription_frame, width=30)
        self.transcription_entry.grid(row=0, column=1, padx=5, pady=5)
        self.save_trans_button = ttk.Button(self.transcription_frame, text="Save Transcription", command=self.save_transcription)
        self.save_trans_button.grid(row=0, column=2, padx=5, pady=5)

        self.recordings_frame = ttk.LabelFrame(self.main_frame, text="Recordings (Click to Play)", padding="5")
        self.recordings_frame.grid(row=8, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))

        self.recordings_listbox = tk.Listbox(self.recordings_frame, width=50, height=5)
        self.recordings_listbox.grid(row=0, column=0, columnspan=3, padx=5, pady=5)
        self.recordings_listbox.bind('<<ListboxSelect>>', self.on_recording_select)

        self.save_excel_button = ttk.Button(self.main_frame, text="Save to Excel", command=self.save_to_excel)
        self.save_excel_button.grid(row=9, column=0, columnspan=3, pady=10)

        self.phrase_var.set(self.phrases[0] if self.phrases else "No Phrases Loaded")
        self.synonym_var.set(f"Synonym {self.synonym_index}")
        self.update_recordings_list()
        self.update_transcription_field()
        self.update_ui_for_recording_status()

    def load_phrases(self):
        phrases_path = BASE_DIR / "phrases.xlsx"
        if phrases_path.exists():
            try:
                df = pd.read_excel(phrases_path)
                if "Phrase" not in df.columns:
                    raise ValueError("Excel file must contain a 'Phrase' column")
                self.phrases = [str(phrase).strip() for phrase in df["Phrase"] if str(phrase).strip()]
            except Exception as e:
                print(f"Error loading phrases from Excel: {e}")
                self.phrases = ["Default Phrase"]
        else:
            self.phrases = ["Default Phrase"]
        if not self.phrases:
            self.phrases = ["Default Phrase"]

    def load_existing_metadata(self):
        metadata_dir = BASE_DIR / "metadata"
        excel_path = metadata_dir / "all_phrases_recordings.xlsx"
        if excel_path.exists():
            try:
                df = pd.read_excel(excel_path)
                df = df.fillna("")
                for _, row in df.iterrows():
                    phrase_key = str(row["Phrase"]).strip()
                    if phrase_key not in self.synonym_data:
                        self.synonym_data[phrase_key] = []
                    synonym_cols = [col for col in df.columns if col.startswith("Synonym_") and col.endswith("_Audio")]
                    for audio_col in synonym_cols:
                        syn_num = int(audio_col.split("_")[1])
                        trans_col = f"Synonym_{syn_num}_Transcription"
                        audio_path = str(row.get(audio_col, ""))
                        transcription = str(row.get(trans_col, ""))
                        if audio_path and not Path(audio_path).exists():
                            audio_path = ""
                        if audio_path or transcription:
                            entry_exists = False
                            for entry in self.synonym_data[phrase_key]:
                                if entry["synonym"] == syn_num:
                                    entry["audio_path"] = audio_path
                                    entry["transcription"] = transcription
                                    entry_exists = True
                                    break
                            if not entry_exists:
                                self.synonym_data[phrase_key].append({
                                    "synonym": syn_num,
                                    "audio_path": audio_path,
                                    "transcription": transcription
                                })
            except Exception as e:
                print(f"Error loading existing metadata: {e}")

    def update_synonym_menu(self):
        if not self.phrases:
            return
        current_phrase = self.phrases[self.phrase_index]
        synonyms = [f"Synonym {entry['synonym']}" for entry in self.synonym_data[current_phrase]]
        if not synonyms:
            synonyms = ["Synonym 1"]
        if self.synonym_menu is not None:
            self.synonym_menu.destroy()
        self.synonym_var = tk.StringVar(value=synonyms[0])
        self.synonym_menu = ttk.OptionMenu(self.selection_frame, self.synonym_var, synonyms[0], *synonyms, command=self.update_synonym)
        self.synonym_menu.grid(row=1, column=1, padx=5, pady=5)

    def add_synonym(self):
        if not self.phrases:
            self.status_label.config(text="Error: No phrases loaded")
            return
        current_phrase = self.phrases[self.phrase_index]
        max_synonym = max(entry["synonym"] for entry in self.synonym_data[current_phrase]) if self.synonym_data[current_phrase] else 0
        new_synonym_num = max_synonym + 1
        self.synonym_data[current_phrase].append({
            "synonym": new_synonym_num,
            "audio_path": "",
            "transcription": ""
        })
        self.update_synonym_menu()
        self.synonym_var.set(f"Synonym {new_synonym_num}")
        self.synonym_index = new_synonym_num
        self.update_transcription_field()
        self.update_recordings_list()
        self.update_ui_for_recording_status()

    def update_phrase(self, *args):
        if not self.phrases:
            self.status_label.config(text="Error: No phrases loaded")
            return
        self.phrase_index = self.phrases.index(self.phrase_var.get())
        current_phrase = self.phrases[self.phrase_index]
        if current_phrase not in self.synonym_data:
            self.synonym_data[current_phrase] = [{"synonym": 1, "audio_path": "", "transcription": ""}]
        self.update_synonym_menu()
        self.synonym_index = 1
        self.synonym_var.set(f"Synonym {self.synonym_index}")
        self.update_recordings_list()
        self.update_transcription_field()
        self.update_ui_for_recording_status()

    def update_synonym(self, *args):
        self.synonym_index = int(self.synonym_var.get().split()[1])
        self.update_transcription_field()
        self.update_ui_for_recording_status()

    def update_ui_for_recording_status(self):
        if not self.phrases:
            return
        current_phrase = self.phrases[self.phrase_index]
        recording_exists = False
        for entry in self.synonym_data[current_phrase]:
            if entry["synonym"] == self.synonym_index and entry["audio_path"]:
                self.current_audio_path = entry["audio_path"]
                if not self.recording:
                    self.play_button.config(state=tk.NORMAL, text="Play")
                    self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
                self.record_button.grid_remove()
                self.stop_button.grid_remove()
                self.recording_frame.grid_remove()  # Hide the recording frame if a recording exists
                self.status_label.config(text="Existing recording found. Play the audio or re-record.")
                recording_exists = True
                break
        if not recording_exists:
            self.current_audio_path = None
            self.play_button.config(state=tk.DISABLED, text="Play")
            self.new_recording_button.grid_remove()
            self.recording_frame.grid(row=4, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))  # Show the recording frame
            self.record_button.grid(row=0, column=0, padx=5, pady=5)
            self.stop_button.grid(row=0, column=1, padx=5, pady=5)
            self.status_label.config(text="No recording found. Start a new recording.")

    def cleanup_audio_resources(self):
        if self.resources_cleaned:
            return
        try:
            if self.stream:
                self.stream.stop_stream()
                self.stream.close()
                self.stream = None
            if self.wf:
                self.wf.close()
                self.wf = None
            if self.pyaudio_instance:
                self.pyaudio_instance.terminate()
                self.pyaudio_instance = None
            self.resources_cleaned = True
        except Exception as e:
            print(f"Error during resource cleanup: {e}")

    def on_recording_select(self, event):
        selection = self.recordings_listbox.curselection()
        if not selection:
            return
        index = selection[0]
        current_phrase = self.phrases[self.phrase_index]
        selected_entry = self.synonym_data[current_phrase][index]
        if not selected_entry["audio_path"]:
            self.status_label.config(text="Error: No audio file for this entry")
            return
        if self.playing or self.current_frame > 0:
            self.playing = False
            if self.play_thread and self.play_thread.is_alive():
                self.play_thread.join(timeout=0.5)
            if self.playback_timer_thread and self.playback_timer_thread.is_alive():
                self.playback_timer_thread.join(timeout=0.5)
            self.current_frame = 0
            self.progress_bar['value'] = 0
            self.timer_label.config(text="00:00")
            self.cleanup_audio_resources()
        self.current_audio_path = selected_entry["audio_path"]
        if not self.recording:
            self.play_button.config(state=tk.NORMAL, text="Play")
            self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
            self.status_label.config(text=f"Selected recording: {Path(self.current_audio_path).name}. Click Play to listen.")
        else:
            self.status_label.config(text=f"Cannot play: Recording in progress.")

    def update_transcription_field(self):
        self.transcription_entry.delete(0, tk.END)
        current_phrase = self.phrases[self.phrase_index]
        for entry in self.synonym_data[current_phrase]:
            if entry["synonym"] == self.synonym_index and entry["transcription"]:
                self.transcription_entry.insert(0, entry["transcription"])
                break

    def start_recording(self):
        self.recording = True
        current_phrase = self.phrases[self.phrase_index]
        for entry in self.synonym_data[current_phrase]:
            if entry["synonym"] == self.synonym_index and entry["audio_path"]:
                try:
                    if os.path.exists(entry["audio_path"]):
                        os.remove(entry["audio_path"])
                        print(f"Deleted old recording: {entry['audio_path']}")
                    entry["audio_path"] = ""
                except Exception as e:
                    print(f"Error deleting old recording: {e}")
                break

        self.current_audio_path = None
        self.record_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.play_button.config(state=tk.DISABLED)
        self.new_recording_button.grid_remove()
        self.status_label.config(text=f"Recording synonym {self.synonym_index} for phrase {self.phrases[self.phrase_index]}...")
        self.timer_label.config(text="00:00")

        self.recording_thread = threading.Thread(target=self.record_audio)
        self.recording_thread.start()
        self.timer_thread = threading.Thread(target=self.update_timer)
        self.timer_thread.start()

    def stop_recording(self):
        self.recording = False
        self.status_label.config(text="Stopping...")
        self.root.update_idletasks()

        if self.recording_thread and self.recording_thread.is_alive():
            self.recording_thread.join(timeout=0.5)

        if self.timer_thread and self.timer_thread.is_alive():
            self.timer_thread.join(timeout=0.1)

        self.record_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        self.record_button.grid_remove()
        self.stop_button.grid_remove()
        self.recording_frame.grid_remove()  # Hide the Recording Controls frame after recording stops

        if self.current_audio_path and os.path.exists(self.current_audio_path):
            self.play_button.config(state=tk.NORMAL, text="Play")
            self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
            self.status_label.config(text="Recording stopped. Play the audio or re-record.")
            self.progress_bar['value'] = 0
        else:
            self.play_button.config(state=tk.DISABLED, text="Play")
            self.new_recording_button.grid_remove()
            self.status_label.config(text="Error: Recorded audio file not found. Start a new recording.")

    def record_audio(self):
        current_phrase = self.phrases[self.phrase_index]
        fs = 48000
        min_duration = 10
        max_duration = 60
        chunk = 1024

        p = None
        stream = None
        try:
            p = pyaudio.PyAudio()
            stream = p.open(format=pyaudio.paInt16,
                            channels=1,
                            rate=fs,
                            input=True,
                            frames_per_buffer=chunk)

            frames = []
            start_time = time.time()
            elapsed_time = 0

            while elapsed_time < max_duration:
                if not self.recording:
                    break
                data = stream.read(chunk, exception_on_overflow=False)
                frames.append(data)
                elapsed_time = time.time() - start_time
                if elapsed_time >= min_duration:
                    self.stop_button.config(state=tk.NORMAL)

            if elapsed_time < min_duration and self.recording:
                remaining_time = min_duration - elapsed_time
                while elapsed_time < min_duration and self.recording:
                    data = stream.read(chunk, exception_on_overflow=False)
                    frames.append(data)
                    elapsed_time = time.time() - start_time

            stream.stop_stream()
            stream.close()
            p.terminate()
            stream = None
            p = None

            if not frames:
                raise ValueError("No audio data recorded")

            output_dir = BASE_DIR / "recordings"
            output_dir.mkdir(exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_phrase = current_phrase.replace(" ", "_")
            filename = f"{safe_phrase}_syn_{self.synonym_index}_{timestamp}.wav"
            filepath = output_dir / filename

            wf = wave.open(str(filepath), 'wb')
            wf.setnchannels(1)
            wf.setsampwidth(2)
            wf.setframerate(fs)
            wf.writeframes(b''.join(frames))
            wf.close()

            if not os.path.exists(filepath):
                raise FileNotFoundError(f"Audio file {filepath} was not created")

            wf_test = wave.open(str(filepath), 'rb')
            frames_to_read = min(48000, wf_test.getnframes())
            data = wf_test.readframes(frames_to_read)
            audio_data = np.frombuffer(data, dtype=np.int16)
            avg_amplitude = np.mean(np.abs(audio_data))
            wf_test.close()

            self.current_audio_path = filepath

            entry_exists = False
            for entry in self.synonym_data[current_phrase]:
                if entry["synonym"] == self.synonym_index:
                    entry["audio_path"] = str(filepath)
                    entry_exists = True
                    break

            if not entry_exists:
                self.synonym_data[current_phrase].append({
                    "synonym": self.synonym_index,
                    "audio_path": str(filepath),
                    "transcription": ""
                })

            self.update_recordings_list()

        except Exception as e:
            self.status_label.config(text=f"Error recording: {e}")
            self.current_audio_path = None
            if stream:
                stream.stop_stream()
                stream.close()
            if p:
                p.terminate()

    def update_timer(self):
        start_time = time.time()
        while self.recording:
            elapsed_time = time.time() - start_time
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            self.timer_label.config(text=f"{minutes:02d}:{seconds:02d}")
            time.sleep(0.1)
            if elapsed_time >= 60:
                self.recording = False
                break

    def update_playback_timer(self, frame_rate):
        while self.playing and self.current_frame < self.total_frames:
            elapsed_time = self.current_frame / frame_rate
            minutes = int(elapsed_time // 60)
            seconds = int(elapsed_time % 60)
            self.timer_label.config(text=f"{minutes:02d}:{seconds:02d}")
            self.progress_bar['value'] = (self.current_frame / self.total_frames) * 100
            self.root.update_idletasks()
            time.sleep(0.1)
        self.timer_label.config(text="00:00")
        self.progress_bar['value'] = 0

    def toggle_play_pause(self):
        if self.recording:
            self.status_label.config(text="Error: Cannot play audio while recording")
            return
        if not self.current_audio_path or not os.path.exists(self.current_audio_path):
            self.status_label.config(text="Error: No recording to play")
            return

        if self.playing:
            self.playing = False
            self.status_label.config(text="Pausing...")
            self.root.update_idletasks()
            if self.play_thread and self.play_thread.is_alive():
                self.play_thread.join(timeout=0.5)
            if self.playback_timer_thread and self.playback_timer_thread.is_alive():
                self.playback_timer_thread.join(timeout=0.5)
            self.play_button.config(text="Play")
            self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
            self.status_label.config(text=f"Paused recording: {Path(self.current_audio_path).name}")
        else:
            self.playing = True
            self.resources_cleaned = False
            self.play_button.config(state=tk.NORMAL, text="Pause")
            self.new_recording_button.grid_remove()
            self.status_label.config(text=f"Resuming playback: {Path(self.current_audio_path).name}")
            self.root.update_idletasks()
            if self.play_thread and self.play_thread.is_alive():
                self.play_thread.join(timeout=0.5)
            self.cleanup_audio_resources()
            self.play_thread = threading.Thread(target=self._play_audio_thread)
            self.play_thread.start()

    def _play_audio_callback(self, in_data, frame_count, time_info, status):
        if not self.playing or self.current_frame >= self.total_frames:
            return (None, pyaudio.paComplete)
        data = self.wf.readframes(frame_count)
        if not data:
            self.playing = False
            return (None, pyaudio.paComplete)
        self.current_frame += frame_count
        return (data, pyaudio.paContinue)

    def _play_audio_thread(self):
        try:
            self.pyaudio_instance = pyaudio.PyAudio()
            self.wf = wave.open(str(self.current_audio_path), 'rb')
            sample_width = self.wf.getsampwidth()
            channels = self.wf.getnchannels()
            frame_rate = self.wf.getframerate()
            if sample_width not in [1, 2, 4] or channels < 1 or frame_rate < 1:
                raise ValueError(f"Invalid wave file parameters")
            format_map = {1: pyaudio.paInt8, 2: pyaudio.paInt16, 4: pyaudio.paInt32}
            audio_format = format_map.get(sample_width, pyaudio.paInt16)
            self.stream = self.pyaudio_instance.open(
                format=audio_format,
                channels=channels,
                rate=frame_rate,
                output=True,
                frames_per_buffer=4096,
                stream_callback=self._play_audio_callback
            )
            self.total_frames = self.wf.getnframes()
            if self.current_frame > 0:
                self.wf.setpos(self.current_frame)
            self.stream.start_stream()
            self.playback_timer_thread = threading.Thread(target=self.update_playback_timer, args=(frame_rate,))
            self.playback_timer_thread.start()
            while self.playing and self.current_frame < self.total_frames:
                if not self.stream.is_active():
                    break
                time.sleep(0.1)
            self.playing = False
            self.play_button.config(state=tk.NORMAL, text="Play")
            self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
            if self.current_frame >= self.total_frames:
                self.current_frame = 0
                self.status_label.config(text=f"Playback finished: {Path(self.current_audio_path).name}")
            else:
                self.status_label.config(text=f"Paused recording: {Path(self.current_audio_path).name}")
            if self.playback_timer_thread and self.playback_timer_thread.is_alive():
                self.playback_timer_thread.join(timeout=0.5)
            self.cleanup_audio_resources()
        except Exception as e:
            self.status_label.config(text=f"Error playing audio: {e}")
            self.playing = False
            self.current_frame = 0
            self.progress_bar['value'] = 0
            self.play_button.config(state=tk.NORMAL, text="Play")
            self.new_recording_button.grid(row=0, column=1, padx=5, pady=5)
            self.cleanup_audio_resources()

    def reset_to_recording_mode(self):
        self.playing = False
        self.current_frame = 0
        self.progress_bar['value'] = 0
        self.timer_label.config(text="00:00")
        self.play_button.config(state=tk.DISABLED, text="Play")
        self.new_recording_button.grid_remove()
        self.recording_frame.grid(row=4, column=0, columnspan=3, pady=5, sticky=(tk.W, tk.E))  # Show the recording frame
        self.record_button.grid(row=0, column=0, padx=5, pady=5)
        self.stop_button.grid(row=0, column=1, padx=5, pady=5)
        self.status_label.config(text="Ready to record")
        if self.playback_timer_thread and self.playback_timer_thread.is_alive():
            self.playback_timer_thread.join(timeout=0.5)
        self.cleanup_audio_resources()

    def save_transcription(self):
        transcription = self.transcription_entry.get().strip()
        if not transcription:
            self.status_label.config(text="Error: Transcription cannot be empty")
            return
        current_phrase = self.phrases[self.phrase_index]
        entry_exists = False
        for entry in self.synonym_data[current_phrase]:
            if entry["synonym"] == self.synonym_index:
                entry["transcription"] = transcription
                entry_exists = True
                break
        if not entry_exists:
            self.synonym_data[current_phrase].append({
                "synonym": self.synonym_index,
                "audio_path": "",
                "transcription": transcription
            })
        self.status_label.config(text=f"Transcription saved for Synonym {self.synonym_index}")
        self.transcription_entry.delete(0, tk.END)
        self.update_recordings_list()

    def update_recordings_list(self):
        self.recordings_listbox.delete(0, tk.END)
        current_phrase = self.phrases[self.phrase_index]
        for entry in self.synonym_data[current_phrase]:
            audio_display = Path(entry["audio_path"]).name if entry["audio_path"] else "No recording"
            self.recordings_listbox.insert(tk.END, f"Synonym {entry['synonym']}: {audio_display} | Transcription: {entry['transcription']}")

    def save_to_excel(self):
        try:
            output_dir = BASE_DIR / "recordings"
            output_dir.mkdir(exist_ok=True)
            metadata_dir = BASE_DIR / "metadata"
            metadata_dir.mkdir(exist_ok=True)
            excel_path = metadata_dir / "all_phrases_recordings.xlsx"
            all_data = []
            max_synonyms = max(len(self.synonym_data[phrase]) for phrase in self.synonym_data) if self.synonym_data else 1
            for phrase_key in sorted(self.synonym_data.keys()):
                phrase_data = {"Phrase": phrase_key}
                recordings = self.synonym_data[phrase_key]
                for syn in range(1, max_synonyms + 1):
                    entry_for_synonym = None
                    for entry in recordings:
                        if entry["synonym"] == syn:
                            entry_for_synonym = entry
                            break
                    if entry_for_synonym:
                        phrase_data[f"Synonym_{syn}_Audio"] = entry_for_synonym["audio_path"]
                        phrase_data[f"Synonym_{syn}_Transcription"] = entry_for_synonym["transcription"]
                    else:
                        phrase_data[f"Synonym_{syn}_Audio"] = ""
                        phrase_data[f"Synonym_{syn}_Transcription"] = ""
                all_data.append(phrase_data)
            df = pd.DataFrame(all_data)
            columns = ["Phrase"] + [col for col in df.columns if col != "Phrase"]
            df = df[columns]
            df.to_excel(excel_path, index=False)
            self.status_label.config(text=f"Data saved to {excel_path}")
        except Exception as e:
            self.status_label.config(text=f"Error: Failed to save Excel: {e}")

if __name__ == "__main__":
    import os
    os.environ["TK_SILENCE_DEPRECATION"] = "1"
    root = tk.Tk()
    app = BangaApp(root)
    root.mainloop()
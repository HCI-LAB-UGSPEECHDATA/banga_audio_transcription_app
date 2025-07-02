import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
import pandas as pd
import uuid

BASE_DIR = Path(__file__).parent

class TranscriptionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Akan Transcription App")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        
        # Configure main window grid
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        # Initialize variables
        self.theme_index = 0
        self.phrase_index = 0
        self.alt_phrase_index = 1
        self.alt_phrase_data = {}
        self.themes = []
        self.phrases = {}
        
        # Load data
        self.load_themes_and_phrases()
        self.initialize_alt_phrase_data()
        self.load_existing_metadata()

        # Apply modern styling
        self.setup_styles()
        
        # Create UI
        self.create_ui()
        
        # Initialize selections
        self.initialize_selections()

    def setup_styles(self):
        """Configure modern styling for the application"""
        style = ttk.Style()
        
        # Configure notebook style for tabs
        style.configure('TNotebook', tabposition='n')
        style.configure('TNotebook.Tab', padding=[12, 8])
        
        # Configure frame styles
        style.configure('Card.TFrame', relief='raised', borderwidth=1)
        style.configure('Header.TLabelframe', font=('Segoe UI', 10, 'bold'))
        
        # Configure button styles
        style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'))
        style.configure('Success.TButton', foreground='white')
        
    def create_ui(self):
        """Create the main user interface"""
        # Main container with padding
        self.main_container = ttk.Frame(self.root, padding="20")
        self.main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.main_container.grid_rowconfigure(1, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)

        # Header section
        self.create_header()
        
        # Main content using notebook for better organization
        self.create_notebook()
        
        # Status bar at bottom
        self.create_status_bar()

    def create_header(self):
        """Create application header with title and controls"""
        header_frame = ttk.Frame(self.main_container)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.grid_columnconfigure(1, weight=1)
        
        # App title
        title_label = ttk.Label(header_frame, text="Akan Transcription App", 
                               font=('Segoe UI', 16, 'bold'))
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Quick save button in header
        save_button = ttk.Button(header_frame, text="💾 Save to Excel", 
                                style='Primary.TButton', command=self.save_to_excel)
        save_button.grid(row=0, column=2, sticky=tk.E, padx=(10, 0))

    def create_notebook(self):
        """Create notebook with tabs for better organization"""
        self.notebook = ttk.Notebook(self.main_container)
        self.notebook.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Main transcription tab
        self.main_tab = ttk.Frame(self.notebook, padding="15")
        self.notebook.add(self.main_tab, text="📝 Transcription")
        
        # Data overview tab
        self.overview_tab = ttk.Frame(self.notebook, padding="15")
        self.notebook.add(self.overview_tab, text="📊 Overview")
        
        # Setup main tab content
        self.setup_main_tab()
        self.setup_overview_tab()

    def setup_main_tab(self):
        """Setup the main transcription tab"""
        self.main_tab.grid_rowconfigure(2, weight=1)
        self.main_tab.grid_columnconfigure(0, weight=1)
        
        # Selection panel
        self.create_selection_panel()
        
        # Transcription panel
        self.create_transcription_panel()
        
        # History panel
        self.create_history_panel()

    def create_selection_panel(self):
        """Create the theme/phrase selection panel"""
        selection_frame = ttk.LabelFrame(self.main_tab, text="📂 Selection", 
                                       style='Header.TLabelframe', padding="15")
        selection_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        selection_frame.grid_columnconfigure(1, weight=1)
        selection_frame.grid_columnconfigure(3, weight=3)  # Give phrase column more weight
        selection_frame.grid_columnconfigure(5, weight=1)
        
        # Theme selection
        ttk.Label(selection_frame, text="Theme:", font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10))
        
        self.theme_var = tk.StringVar(value=self.themes[0] if self.themes else "No Themes")
        self.theme_combo = ttk.Combobox(selection_frame, textvariable=self.theme_var, 
                                       state="readonly", width=15)  # Reduced from 20 to 15
        self.theme_combo['values'] = self.themes
        self.theme_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 20))
        self.theme_combo.bind('<<ComboboxSelected>>', self.update_theme)
        
        # Phrase selection
        ttk.Label(selection_frame, text="Phrase:", font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=2, sticky=tk.W, padx=(0, 10))
        
        self.phrase_var = tk.StringVar()
        self.phrase_combo = ttk.Combobox(selection_frame, textvariable=self.phrase_var, 
                                        state="readonly", width=50)  # Increased from 30 to 50
        self.phrase_combo.grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(0, 20))
        self.phrase_combo.bind('<<ComboboxSelected>>', self.update_phrase)
        
        # Alternate phrase selection
        ttk.Label(selection_frame, text="Alternative:", font=('Segoe UI', 9, 'bold')).grid(
            row=0, column=4, sticky=tk.W, padx=(0, 10))
        
        self.alt_phrase_var = tk.StringVar()
        self.alt_phrase_combo = ttk.Combobox(selection_frame, textvariable=self.alt_phrase_var, 
                                           state="readonly", width=12)  # Reduced from 15 to 12
        self.alt_phrase_combo.grid(row=0, column=5, sticky=(tk.W, tk.E))
        self.alt_phrase_combo.bind('<<ComboboxSelected>>', self.update_alt_phrase)

    def create_transcription_panel(self):
        """Create the transcription input panel"""
        trans_frame = ttk.LabelFrame(self.main_tab, text="✏️ Transcription", 
                                   style='Header.TLabelframe', padding="15")
        trans_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        trans_frame.grid_rowconfigure(1, weight=1)
        trans_frame.grid_columnconfigure(0, weight=1)
        
        # Current phrase display
        self.current_phrase_label = ttk.Label(trans_frame, text="Current Phrase: ", 
                                            font=('Segoe UI', 9, 'italic'),
                                            foreground='gray')
        self.current_phrase_label.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Text input with frame
        input_frame = ttk.Frame(trans_frame)
        input_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        input_frame.grid_rowconfigure(0, weight=1)
        input_frame.grid_columnconfigure(0, weight=1)
        
        self.transcription_text = tk.Text(input_frame, height=4, font=('Segoe UI', 10),
                                        wrap=tk.WORD, relief='solid', borderwidth=1)
        self.transcription_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for text area
        scrollbar = ttk.Scrollbar(input_frame, orient="vertical", command=self.transcription_text.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.transcription_text.configure(yscrollcommand=scrollbar.set)
        
        # Button frame - only save button now
        button_frame = ttk.Frame(trans_frame)
        button_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        ttk.Button(button_frame, text="💾 Save Transcription", 
                  style='Primary.TButton',
                  command=self.save_transcription).grid(row=0, column=0, sticky=tk.E)

    def create_history_panel(self):
        """Create the transcription history panel"""
        history_frame = ttk.LabelFrame(self.main_tab, text="📜 History", 
                                     style='Header.TLabelframe', padding="15")
        history_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        history_frame.grid_rowconfigure(1, weight=1)
        history_frame.grid_columnconfigure(0, weight=1)
        
        # Header with button
        history_header = ttk.Frame(history_frame)
        history_header.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        history_header.grid_columnconfigure(0, weight=1)
        
        ttk.Button(history_header, text="+ New Alternative", 
                  command=self.add_alt_phrase).grid(row=0, column=1, sticky=tk.E)
        
        # Treeview for better display
        columns = ('Alternative', 'Transcription', 'Length')
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show='headings', height=8)
        
        # Configure columns - reduced width for Length column
        self.history_tree.heading('Alternative', text='Alternative')
        self.history_tree.heading('Transcription', text='Transcription')
        self.history_tree.heading('Length', text='Length')
        
        self.history_tree.column('Alternative', width=80, minwidth=80)
        self.history_tree.column('Transcription', width=450, minwidth=200)
        self.history_tree.column('Length', width=60, minwidth=50)  # Reduced from 80 to 60
        
        self.history_tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.history_tree.bind('<<TreeviewSelect>>', self.on_history_select)
        
        # Scrollbar for treeview
        tree_scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_tree.yview)
        tree_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.history_tree.configure(yscrollcommand=tree_scrollbar.set)

    def setup_overview_tab(self):
        """Setup the overview tab"""
        self.overview_tab.grid_rowconfigure(0, weight=1)
        self.overview_tab.grid_columnconfigure(0, weight=1)
        
        # Stats frame
        stats_frame = ttk.LabelFrame(self.overview_tab, text="📈 Statistics", 
                                   style='Header.TLabelframe', padding="15")
        stats_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.stats_text = tk.Text(stats_frame, height=20, font=('Segoe UI', 10),
                                state='disabled', wrap=tk.WORD)
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Update stats initially
        self.update_stats()

    def create_status_bar(self):
        """Create status bar at bottom"""
        status_frame = ttk.Frame(self.main_container)
        status_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        status_frame.grid_columnconfigure(1, weight=1)
        
        ttk.Label(status_frame, text="Status:").grid(row=0, column=0, sticky=tk.W)
        
        self.status_label = ttk.Label(status_frame, text="Ready to transcribe", 
                                    relief='sunken', padding="5")
        self.status_label.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))

    def load_themes_and_phrases(self):
        """Load themes and phrases from Excel file"""
        phrases_path = BASE_DIR / "FocusGroup_questions_v1.xlsx"
        if phrases_path.exists():
            try:
                xl = pd.ExcelFile(phrases_path)
                self.themes = xl.sheet_names
                self.phrases = {}
                for theme in self.themes:
                    df = pd.read_excel(phrases_path, sheet_name=theme)
                    self.phrases[theme] = [str(row[0]).strip() for row in df.values if str(row[0]).strip()]
                    if not self.phrases[theme]:
                        self.phrases[theme] = ["Default Phrase"]
            except Exception as e:
                messagebox.showerror("Error", f"Error loading themes and phrases: {e}")
                self.themes = ["Default Theme"]
                self.phrases = {"Default Theme": ["Default Phrase"]}
        else:
            self.themes = ["Default Theme"]
            self.phrases = {"Default Theme": ["Default Phrase"]}

    def initialize_alt_phrase_data(self):
        """Initialize alternate phrase data structure"""
        for theme in self.themes:
            for phrase in self.phrases[theme]:
                if (theme, phrase) not in self.alt_phrase_data:
                    self.alt_phrase_data[(theme, phrase)] = []
                if not self.alt_phrase_data[(theme, phrase)]:
                    self.alt_phrase_data[(theme, phrase)].append({
                        "alt_phrase": 1,
                        "transcription": ""
                    })

    def load_existing_metadata(self):
        """Load existing transcription data"""
        metadata_dir = BASE_DIR / "metadata"
        excel_path = metadata_dir / "transcriptions.xlsx"
        if excel_path.exists():
            try:
                df = pd.read_excel(excel_path)
                df = df.fillna("")
                for _, row in df.iterrows():
                    theme = str(row["Theme"]).strip()
                    phrase = str(row["Phrase"]).strip()
                    if (theme, phrase) not in self.alt_phrase_data:
                        self.alt_phrase_data[(theme, phrase)] = []
                    alt_phrase_cols = [col for col in df.columns if col.startswith("Alternate_Phrase_") and col.endswith("_Transcription")]
                    for trans_col in alt_phrase_cols:
                        alt_num = int(trans_col.split("_")[2])
                        transcription = str(row.get(trans_col, ""))
                        if transcription:
                            entry_exists = False
                            for entry in self.alt_phrase_data[(theme, phrase)]:
                                if entry["alt_phrase"] == alt_num:
                                    entry["transcription"] = transcription
                                    entry_exists = True
                                    break
                            if not entry_exists:
                                self.alt_phrase_data[(theme, phrase)].append({
                                    "alt_phrase": alt_num,
                                    "transcription": transcription
                                })
            except Exception as e:
                messagebox.showwarning("Warning", f"Error loading existing data: {e}")

    def initialize_selections(self):
        """Initialize UI selections"""
        if self.themes:
            self.theme_var.set(self.themes[0])
            self.update_phrase_combo()
            if self.phrases[self.themes[0]]:
                self.phrase_var.set(self.phrases[self.themes[0]][0])
                self.update_alt_phrase_combo()
                self.alt_phrase_var.set("Alternative 1")
                self.update_current_phrase_display()
                self.update_history()
                self.update_transcription_field()

    def update_phrase_combo(self):
        """Update phrase combobox values"""
        if not self.themes:
            return
        current_theme = self.themes[self.theme_index]
        phrases = self.phrases[current_theme]
        self.phrase_combo['values'] = phrases

    def update_alt_phrase_combo(self):
        """Update alternate phrase combobox values"""
        if not self.themes:
            return
        current_theme = self.themes[self.theme_index]
        current_phrase = self.phrases[current_theme][self.phrase_index]
        alt_phrases = [f"Alternative {entry['alt_phrase']}" for entry in self.alt_phrase_data[(current_theme, current_phrase)]]
        if not alt_phrases:
            alt_phrases = ["Alternative 1"]
        self.alt_phrase_combo['values'] = alt_phrases

    def update_current_phrase_display(self):
        """Update the current phrase display"""
        if self.themes and self.phrase_index < len(self.phrases[self.themes[self.theme_index]]):
            current_phrase = self.phrases[self.themes[self.theme_index]][self.phrase_index]
            self.current_phrase_label.config(text=f"Current Phrase: {current_phrase}")

    def update_theme(self, *args):
        """Handle theme selection change"""
        if not self.themes:
            return
        self.theme_index = self.themes.index(self.theme_var.get())
        self.phrase_index = 0
        self.update_phrase_combo()
        current_theme = self.themes[self.theme_index]
        if self.phrases[current_theme]:
            self.phrase_var.set(self.phrases[current_theme][0])
        self.update_alt_phrase_combo()
        self.alt_phrase_index = 1
        self.alt_phrase_var.set("Alternative 1")
        self.update_current_phrase_display()
        self.update_history()
        self.update_transcription_field()

    def update_phrase(self, *args):
        """Handle phrase selection change"""
        if not self.themes:
            return
        current_theme = self.themes[self.theme_index]
        self.phrase_index = self.phrases[current_theme].index(self.phrase_var.get())
        current_phrase = self.phrases[current_theme][self.phrase_index]
        if (current_theme, current_phrase) not in self.alt_phrase_data:
            self.alt_phrase_data[(current_theme, current_phrase)] = [{"alt_phrase": 1, "transcription": ""}]
        self.update_alt_phrase_combo()
        self.alt_phrase_index = 1
        self.alt_phrase_var.set("Alternative 1")
        self.update_current_phrase_display()
        self.update_history()
        self.update_transcription_field()

    def update_alt_phrase(self, *args):
        """Handle alternate phrase selection change"""
        selected = self.alt_phrase_var.get()
        if selected.startswith("Alternative "):
            self.alt_phrase_index = int(selected.split()[1])
            self.update_transcription_field()

    def update_transcription_field(self):
        """Update transcription text field"""
        self.transcription_text.delete("1.0", tk.END)
        current_theme = self.themes[self.theme_index]
        current_phrase = self.phrases[current_theme][self.phrase_index]
        for entry in self.alt_phrase_data[(current_theme, current_phrase)]:
            if entry["alt_phrase"] == self.alt_phrase_index and entry["transcription"]:
                self.transcription_text.insert("1.0", entry["transcription"])
                break

    def on_history_select(self, event):
        """Handle history selection"""
        selection = self.history_tree.selection()
        if not selection:
            return
        item = self.history_tree.item(selection[0])
        version_text = item['values'][0]
        version_num = int(version_text.split()[1])
        self.alt_phrase_index = version_num
        self.alt_phrase_var.set(f"Alternative {version_num}")
        self.update_transcription_field()
        self.status_label.config(text=f"Selected Alternative {version_num}")

    def add_alt_phrase(self):
        """Add new alternate phrase version"""
        if not self.themes:
            return
        current_theme = self.themes[self.theme_index]
        current_phrase = self.phrases[current_theme][self.phrase_index]
        max_alt = max(entry["alt_phrase"] for entry in self.alt_phrase_data[(current_theme, current_phrase)]) if self.alt_phrase_data[(current_theme, current_phrase)] else 0
        new_alt_num = max_alt + 1
        self.alt_phrase_data[(current_theme, current_phrase)].append({
            "alt_phrase": new_alt_num,
            "transcription": ""
        })
        self.update_alt_phrase_combo()
        self.alt_phrase_var.set(f"Alternative {new_alt_num}")
        self.alt_phrase_index = new_alt_num
        self.update_transcription_field()
        self.update_history()
        self.status_label.config(text=f"Created Alternative {new_alt_num}")

    def save_transcription(self):
        """Save current transcription"""
        transcription = self.transcription_text.get("1.0", tk.END).strip()
        if not transcription:
            messagebox.showwarning("Warning", "Transcription cannot be empty")
            return
        
        current_theme = self.themes[self.theme_index]
        current_phrase = self.phrases[current_theme][self.phrase_index]
        entry_exists = False
        for entry in self.alt_phrase_data[(current_theme, current_phrase)]:
            if entry["alt_phrase"] == self.alt_phrase_index:
                entry["transcription"] = transcription
                entry_exists = True
                break
        if not entry_exists:
            self.alt_phrase_data[(current_theme, current_phrase)].append({
                "alt_phrase": self.alt_phrase_index,
                "transcription": transcription
            })
        
        self.status_label.config(text=f"✅ Saved Version {self.alt_phrase_index}")
        self.transcription_text.delete("1.0", tk.END)
        self.update_history()
        self.update_stats()

    def update_history(self):
        """Update history display"""
        # Clear existing items
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        current_theme = self.themes[self.theme_index]
        current_phrase = self.phrases[current_theme][self.phrase_index]
        for entry in self.alt_phrase_data[(current_theme, current_phrase)]:
            version = f"Version {entry['alt_phrase']}"
            transcription = entry["transcription"] if entry["transcription"] else "No transcription"
            length = len(entry["transcription"]) if entry["transcription"] else 0
            # Truncate long transcriptions for display
            display_trans = transcription[:50] + "..." if len(transcription) > 50 else transcription
            self.history_tree.insert("", "end", values=(version, display_trans, length))

    def update_stats(self):
        """Update statistics display"""
        self.stats_text.config(state='normal')
        self.stats_text.delete("1.0", tk.END)
        
        stats = []
        stats.append("TRANSCRIPTION STATISTICS\n" + "="*30 + "\n")
        
        total_themes = len(self.themes)
        total_phrases = sum(len(self.phrases[theme]) for theme in self.themes)
        total_transcriptions = 0
        total_chars = 0
        
        for key in self.alt_phrase_data:
            for entry in self.alt_phrase_data[key]:
                if entry["transcription"]:
                    total_transcriptions += 1
                    total_chars += len(entry["transcription"])
        
        stats.append(f"Total Themes: {total_themes}")
        stats.append(f"Total Phrases: {total_phrases}")
        stats.append(f"Completed Transcriptions: {total_transcriptions}")
        stats.append(f"Total Characters: {total_chars:,}")
        
        if total_transcriptions > 0:
            stats.append(f"Average Length: {total_chars // total_transcriptions} characters")
        
        completion_rate = (total_transcriptions / total_phrases * 100) if total_phrases > 0 else 0
        stats.append(f"Completion Rate: {completion_rate:.1f}%")
        
        stats.append("\n" + "BY THEME" + "\n" + "-"*20)
        for theme in self.themes:
            theme_phrases = len(self.phrases[theme])
            theme_transcriptions = 0
            for phrase in self.phrases[theme]:
                for entry in self.alt_phrase_data.get((theme, phrase), []):
                    if entry["transcription"]:
                        theme_transcriptions += 1
            stats.append(f"{theme}: {theme_transcriptions}/{theme_phrases} phrases completed")
        
        self.stats_text.insert("1.0", "\n".join(stats))
        self.stats_text.config(state='disabled')

    def save_to_excel(self):
        """Save all data to Excel file"""
        try:
            metadata_dir = BASE_DIR / "metadata"
            metadata_dir.mkdir(exist_ok=True)
            excel_path = metadata_dir / "transcriptions.xlsx"
            
            all_data = []
            max_alt_phrases = max(len(self.alt_phrase_data[key]) for key in self.alt_phrase_data) if self.alt_phrase_data else 1
            
            for theme, phrase in sorted(self.alt_phrase_data.keys()):
                phrase_data = {"Theme": theme, "Phrase": phrase}
                transcriptions = self.alt_phrase_data[(theme, phrase)]
                for alt in range(1, max_alt_phrases + 1):
                    entry_for_alt_phrase = None
                    for entry in transcriptions:
                        if entry["alt_phrase"] == alt:
                            entry_for_alt_phrase = entry
                            break
                    if entry_for_alt_phrase:
                        phrase_data[f"Alternate_Phrase_{alt}_Transcription"] = entry_for_alt_phrase["transcription"]
                    else:
                        phrase_data[f"Alternate_Phrase_{alt}_Transcription"] = ""
                all_data.append(phrase_data)
            
            df = pd.DataFrame(all_data)
            columns = ["Theme", "Phrase"] + [col for col in df.columns if col.startswith("Alternate_Phrase_")]
            df = df[columns]
            df.to_excel(excel_path, index=False)
            
            self.status_label.config(text=f"✅ Data saved to {excel_path}")
            self.update_stats()  # Refresh stats after save
            messagebox.showinfo("Success", f"Data successfully saved to:\n{excel_path}")
            
        except Exception as e:
            error_msg = f"Failed to save Excel file: {e}"
            self.status_label.config(text=f"❌ {error_msg}")
            messagebox.showerror("Error", error_msg)

if __name__ == "__main__":
    import os
    os.environ["TK_SILENCE_DEPRECATION"] = "1"
    root = tk.Tk()
    app = TranscriptionApp(root)
    root.mainloop()
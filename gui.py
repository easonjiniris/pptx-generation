import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import threading
from slide import generate_slide, regenerate_slide
from dotenv import load_dotenv, set_key

# Fix DPI scaling on Windows
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass


WIDTH = 1600
HEIGHT = 900


class TextRedirector:
    """Redirects stdout to a tkinter Text widget."""

    def __init__(self, widget):
        self.widget = widget

    def write(self, string):
        self.widget.config(state=tk.NORMAL)
        self.widget.insert(tk.END, string)
        self.widget.see(tk.END)
        self.widget.config(state=tk.DISABLED)
        self.widget.update_idletasks()

    def flush(self):
        pass


class SlideGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Slide Generator")
        self.root.geometry(f"{WIDTH}x{HEIGHT}")

        self.docx_path = None
        self.dir = os.path.dirname(os.path.abspath(__file__))
        self.response_path = os.path.join(self.dir, "output", "response.txt")
        self.output_path = os.path.join(self.dir, "output", "output.pptx")
        self.env_path = os.path.join(self.dir, ".env")
        load_dotenv(self.env_path)

        self._create_widgets()

    def _create_widgets(self):
        # Top frame for 3x2 grid
        top_frame = tk.Frame(self.root, padx=10, pady=10)
        top_frame.pack(fill=tk.X)

        # Row 1: File browser and Generate button
        self.browse_button = tk.Button(
            top_frame,
            text="Browse DOCX File",
            command=self._browse_file,
            width=20,
            height=2
        )
        self.browse_button.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.generate_button = tk.Button(
            top_frame,
            text="Generate",
            command=self._generate,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.generate_button.grid(
            row=0, column=1, padx=5, pady=5, sticky="nsew")

        # Row 2: Regenerate and View Response buttons
        self.view_response_button = tk.Button(
            top_frame,
            text="View Response",
            command=self._view_response,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.view_response_button.grid(
            row=1, column=0, padx=5, pady=5, sticky="nsew")

        self.regenerate_button = tk.Button(
            top_frame,
            text="Regenerate",
            command=self._regenerate,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.regenerate_button.grid(
            row=1, column=1, padx=5, pady=5, sticky="nsew")

        # Row 3: Open Output and Settings buttons
        self.open_output_button = tk.Button(
            top_frame,
            text="Open Output",
            command=self._open_output,
            width=20,
            height=2,
            state=tk.DISABLED
        )
        self.open_output_button.grid(
            row=2, column=0, padx=5, pady=5, sticky="nsew")

        self.settings_button = tk.Button(
            top_frame,
            text="Settings",
            command=self._open_settings,
            width=20,
            height=2
        )
        self.settings_button.grid(
            row=2, column=1, padx=5, pady=5, sticky="nsew")

        # Configure grid weights for equal sizing
        top_frame.grid_columnconfigure(0, weight=1)
        top_frame.grid_columnconfigure(1, weight=1)

        # File path label
        self.file_label = tk.Label(
            self.root,
            text="No file selected",
            fg="gray",
            anchor="w",
            padx=10
        )
        self.file_label.pack(fill=tk.X)

        # Separator
        separator = tk.Frame(self.root, height=2, bd=1, relief=tk.SUNKEN)
        separator.pack(fill=tk.X, padx=10, pady=5)

        # Output text area
        output_frame = tk.Frame(self.root, padx=10, pady=5)
        output_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(output_frame, text="Output Log:", anchor="w").pack(fill=tk.X)

        self.output_text = scrolledtext.ScrolledText(
            output_frame,
            wrap=tk.WORD,
            height=15,
            state=tk.DISABLED
        )
        self.output_text.pack(fill=tk.BOTH, expand=True)

        # Redirect stdout to the text widget
        sys.stdout = TextRedirector(self.output_text)

        print("Application started. Please select a DOCX file to begin.")

    def _browse_file(self):
        """Open file browser to select DOCX file."""
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )

        if file_path:
            self.docx_path = file_path
            self.file_label.config(
                text=f"Selected: {os.path.basename(file_path)}", fg="black")
            self.generate_button.config(state=tk.NORMAL)
            print(f"Selected file: {file_path}")

    def _generate(self):
        """Generate slides from the selected DOCX file."""
        if not self.docx_path:
            messagebox.showerror("Error", "Please select a DOCX file first.")
            return

        # Disable buttons during generation
        self._set_buttons_state(tk.DISABLED)

        def run_generation():
            try:
                generate_slide(self.docx_path)
                # Enable buttons after successful generation
                self.root.after(
                    0, lambda: self.regenerate_button.config(state=tk.NORMAL))
                self.root.after(
                    0, lambda: self.view_response_button.config(state=tk.NORMAL))
                self.root.after(
                    0, lambda: self.open_output_button.config(state=tk.NORMAL))
            except Exception as e:
                print(f"\nError during generation: {str(e)}")
                messagebox.showerror("Generation Error", str(e))
            finally:
                self.root.after(0, lambda: self._set_buttons_state(
                    tk.NORMAL, keep_regenerate=True))

        # Run in separate thread to prevent UI freezing
        thread = threading.Thread(target=run_generation, daemon=True)
        thread.start()

    def _regenerate(self):
        """Regenerate slides from the saved response.txt file."""
        if not os.path.exists(self.response_path):
            messagebox.showerror(
                "Error", "response.txt not found. Please generate slides first.")
            return

        # Disable buttons during regeneration
        self._set_buttons_state(tk.DISABLED)

        def run_regeneration():
            try:
                print(f"\n{'='*50}")
                print("Starting slide regeneration...")
                print(f"{'='*50}\n")
                regenerate_slide(self.response_path)
                print(f"\n{'='*50}")
                print("Regeneration complete!")
                print(f"{'='*50}\n")
            except Exception as e:
                print(f"\nError during regeneration: {str(e)}")
                messagebox.showerror("Regeneration Error", str(e))
            finally:
                self.root.after(0, lambda: self._set_buttons_state(
                    tk.NORMAL, keep_regenerate=True))

        # Run in separate thread to prevent UI freezing
        thread = threading.Thread(target=run_regeneration, daemon=True)
        thread.start()

    def _view_response(self):
        """Open response.txt in the default text editor."""
        if not os.path.exists(self.response_path):
            messagebox.showwarning(
                "File Not Found", "response.txt does not exist yet. Please generate slides first.")
            return

        try:
            if sys.platform == "win32":
                os.startfile(self.response_path)
            elif sys.platform == "darwin":
                os.system(f"open '{self.response_path}'")
            else:
                os.system(f"xdg-open '{self.response_path}'")
            print(f"Opened {self.response_path}")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Could not open response.txt: {str(e)}")

    def _open_output(self):
        """Open the output PowerPoint file."""
        if not os.path.exists(self.output_path):
            messagebox.showwarning(
                "File Not Found", "output.pptx does not exist yet. Please generate slides first.")
            return

        try:
            if sys.platform == "win32":
                os.startfile(self.output_path)
            elif sys.platform == "darwin":
                os.system(f"open '{self.output_path}'")
            else:
                os.system(f"xdg-open '{self.output_path}'")
            print(f"Opened {self.output_path}")
        except Exception as e:
            messagebox.showerror(
                "Error", f"Could not open output.pptx: {str(e)}")

    def _open_settings(self):
        """Open the settings window."""
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry(f"{WIDTH//2}x{HEIGHT}")

        # Center the settings window
        settings_window.transient(self.root)
        settings_window.grab_set()

        # Main frame with padding
        main_frame = tk.Frame(settings_window, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 1. Model Selection
        model_label = tk.Label(main_frame, text="Model:",
                               font=("Arial", 10, "bold"))
        model_label.pack(anchor="w", pady=(0, 5))

        model_var = tk.StringVar(value=os.getenv("MODEL_NAME", "gpt-4o-mini"))
        model_dropdown = ttk.Combobox(
            main_frame,
            textvariable=model_var,
            values=["gpt-4o-mini", "gpt-4.1-mini"],
            state="readonly",
            width=40
        )
        model_dropdown.pack(fill=tk.X, pady=(0, 15))

        # 2. API Key Entry
        api_key_label = tk.Label(
            main_frame, text="API Key:", font=("Arial", 10, "bold"))
        api_key_label.pack(anchor="w", pady=(0, 5))

        api_key_var = tk.StringVar(value=os.getenv("OPENAI_API_KEY", ""))
        api_key_entry = tk.Entry(
            main_frame, textvariable=api_key_var, width=40)
        api_key_entry.pack(fill=tk.X, pady=(0, 15))

        # 3. Additional Prompt
        prompt_label = tk.Label(
            main_frame, text="Additional Prompt:", font=("Arial", 10, "bold"))
        prompt_label.pack(anchor="w", pady=(0, 5))

        prompt_text = scrolledtext.ScrolledText(
            main_frame,
            wrap=tk.WORD,
            height=10,
            width=40
        )
        prompt_text.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        prompt_text.insert("1.0", os.getenv("ADDITIONAL_PROMPT", ""))

        # Button frame
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        def save_settings():
            """Save settings to .env file."""
            try:
                # Ensure .env file exists
                if not os.path.exists(self.env_path):
                    with open(self.env_path, 'w') as f:
                        pass

                # Save settings
                set_key(self.env_path, "MODEL_NAME", model_var.get())
                set_key(self.env_path, "OPENAI_API_KEY", api_key_var.get())
                set_key(self.env_path, "ADDITIONAL_PROMPT",
                        prompt_text.get("1.0", "end-1c"))

                messagebox.showinfo("Success", "Settings saved successfully!")
                settings_window.destroy()
            except Exception as e:
                messagebox.showerror(
                    "Error", f"Failed to save settings: {str(e)}")

        save_button = tk.Button(
            button_frame,
            text="Save",
            command=save_settings,
            width=15,
            bg="#4CAF50",
            fg="white"
        )
        save_button.pack(side=tk.LEFT, padx=(0, 5))

        cancel_button = tk.Button(
            button_frame,
            text="Cancel",
            command=settings_window.destroy,
            width=15
        )
        cancel_button.pack(side=tk.LEFT)

    def _set_buttons_state(self, state, keep_regenerate=False):
        """Enable or disable all buttons."""
        self.browse_button.config(state=state)

        # Only enable generate button if a file is selected
        if state == tk.NORMAL and self.docx_path:
            self.generate_button.config(state=state)
        else:
            self.generate_button.config(state=tk.DISABLED)

        # Keep regenerate button state if requested
        if not keep_regenerate:
            self.regenerate_button.config(state=state)

        self.view_response_button.config(state=state)

        # Keep open output button state if requested (same as regenerate)
        if not keep_regenerate:
            self.open_output_button.config(state=state)

        # Settings button is always enabled
        self.settings_button.config(state=state)


def main():
    root = tk.Tk()
    app = SlideGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

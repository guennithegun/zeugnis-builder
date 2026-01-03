import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkcalendar import Calendar
import pikepdf
import os
import re
import threading
import time
from datetime import datetime

# --- CONFIGURATION ---
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class ZeugnisApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Setup
        self.title("Zeugnis Copy Tool V7")
        self.geometry("900x850")

        # --- VARIABLES ---
        self.source_folder = ctk.StringVar()
        self.template_file = ctk.StringVar()
        self.target_folder = ctk.StringVar()
        self.class_name = ctk.StringVar()
        self.school_year = ctk.StringVar(value="2024/25")

        # Default Date (Today)
        self.sign_date = ctk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))

        # UI Layout
        self.create_ui()

    def create_ui(self):
        # 1. HEADER
        self.header_frame = ctk.CTkFrame(self, corner_radius=0, fg_color="transparent")
        self.header_frame.pack(fill="x", padx=20, pady=(20, 10))

        ctk.CTkLabel(
            self.header_frame,
            text="Zeugnis Manager",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            self.header_frame,
            text="Automatischer Kopier-Assistent",
            font=ctk.CTkFont(size=14),
            text_color="gray",
        ).pack(side="left", padx=15, pady=(8, 0))

        # 2. MAIN SCROLLABLE AREA
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=10)

        # --- CARD 1: FILES ---
        self.create_section_label(self.main_frame, "1. Dateien")
        self.file_card = ctk.CTkFrame(self.main_frame, corner_radius=15)
        self.file_card.pack(fill="x", pady=(5, 20))

        self.create_file_row(
            self.file_card, "Ausgangsordner:", self.source_folder, is_folder=True
        )
        self.create_file_row(
            self.file_card, "Zeugnis Vorlage:", self.template_file, is_folder=False
        )
        self.create_file_row(
            self.file_card, "Zielordner:", self.target_folder, is_folder=True
        )

        # --- CARD 2: SETTINGS ---
        self.create_section_label(self.main_frame, "2. Weitere Einstellungen")
        self.settings_card = ctk.CTkFrame(self.main_frame, corner_radius=15)
        self.settings_card.pack(fill="x", pady=(5, 20))

        # Year
        ctk.CTkLabel(self.settings_card, text="Schuljahr:").grid(
            row=0, column=0, padx=20, pady=20, sticky="w"
        )
        years = ["2023/24", "2024/25", "2025/26", "2026/27"]
        self.year_cb = ctk.CTkComboBox(
            self.settings_card,
            values=years,
            variable=self.school_year,
            width=150,
            corner_radius=10,
        )
        self.year_cb.grid(row=0, column=1, padx=10, pady=20, sticky="w")

        # Class
        ctk.CTkLabel(self.settings_card, text="Klasse (z.B. 4b):").grid(
            row=0, column=2, padx=20, pady=20, sticky="w"
        )
        ctk.CTkEntry(
            self.settings_card,
            textvariable=self.class_name,
            width=100,
            corner_radius=10,
        ).grid(row=0, column=3, padx=10, pady=20, sticky="w")

        # Date Picker
        ctk.CTkLabel(self.settings_card, text="Datum (Unterschrift):").grid(
            row=1, column=0, padx=20, pady=(0, 20), sticky="w"
        )

        date_frame = ctk.CTkFrame(self.settings_card, fg_color="transparent")
        date_frame.grid(row=1, column=1, padx=10, pady=(0, 20), sticky="w")

        self.date_entry = ctk.CTkEntry(
            date_frame, textvariable=self.sign_date, width=110, corner_radius=10
        )
        self.date_entry.pack(side="left", padx=(0, 5))

        ctk.CTkButton(
            date_frame,
            text="📅",
            width=40,
            corner_radius=10,
            fg_color="#34495e",
            command=self.open_calendar_popup,
        ).pack(side="left")

        # --- BUTTON ---
        self.start_btn = ctk.CTkButton(
            self.main_frame,
            text="ZEUGNISSE KOPIEREN",
            font=ctk.CTkFont(size=16, weight="bold"),
            height=50,
            corner_radius=25,
            fg_color="#2ecc71",
            hover_color="#27ae60",
            command=self.start_thread,
        )
        self.start_btn.pack(fill="x", pady=10)

        # --- LOGS ---
        self.create_section_label(self.main_frame, "Logs")
        self.log_area = ctk.CTkTextbox(
            self.main_frame, height=150, corner_radius=15, font=("Consolas", 12)
        )
        self.log_area.pack(fill="x", pady=(5, 20))
        self.log_area.configure(state="disabled")

    def create_section_label(self, parent, text):
        ctk.CTkLabel(parent, text=text, font=ctk.CTkFont(size=14, weight="bold")).pack(
            anchor="w", padx=5
        )

    def create_file_row(self, parent, label_text, var, is_folder):
        row = ctk.CTkFrame(parent, fg_color="transparent")
        row.pack(fill="x", padx=10, pady=10)

        ctk.CTkLabel(row, text=label_text, width=130, anchor="w").pack(
            side="left", padx=10
        )
        ctk.CTkEntry(row, textvariable=var, corner_radius=10).pack(
            side="left", fill="x", expand=True, padx=5
        )

        cmd = self.select_dir if is_folder else self.select_file
        ctk.CTkButton(
            row, text="Suchen", width=80, corner_radius=10, command=lambda: cmd(var)
        ).pack(side="left", padx=5)

    def select_dir(self, var):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    def select_file(self, var):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            var.set(path)

    def log(self, message):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.configure(state="disabled")

    # --- UPDATED CALENDAR LOGIC ---
    def open_calendar_popup(self):
        top = ctk.CTkToplevel(self)
        top.title("Datum auswählen")
        top.geometry("300x250")
        top.attributes("-topmost", True)

        # Styles for better visibility
        # selectbackground: The circle color (Green)
        # selectforeground: The text color inside the circle (White)
        # headersbackground: Dark bar for Mo/Tu/We...
        # headersforeground: White text for headers
        cal = Calendar(
            top,
            selectmode="day",
            date_pattern="yyyy-mm-dd",
            cursor="hand2",
            selectbackground="#2ecc71",  # Green Circle
            selectforeground="white",  # White Text
            headersbackground="#34495e",  # Dark Header
            headersforeground="white",
            bordercolor="gray",
            normalbackground="white",
            weekendbackground="white",
        )
        cal.pack(pady=20)

        def set_date():
            selected = cal.get_date()
            self.sign_date.set(selected)
            top.destroy()

        ctk.CTkButton(
            top,
            text="Übernehmen",
            command=set_date,
            fg_color="#2ecc71",
            text_color="white",
        ).pack(pady=10)

    # ------------------------------

    def start_thread(self):
        t = threading.Thread(target=self.run_process)
        t.start()

    def transfer_data(
        self, source_path, template_path, output_path, s_year, s_class, s_date
    ):
        try:
            pdf_source = pikepdf.Pdf.open(source_path)
            pdf_target = pikepdf.Pdf.open(template_path)

            try:
                xfa_source = pdf_source.Root.AcroForm.XFA
                xfa_target = pdf_target.Root.AcroForm.XFA
            except AttributeError:
                self.log(f"SKIP: {os.path.basename(source_path)} is not an XFA form.")
                return False

            source_idx = -1
            target_idx = -1

            for i, item in enumerate(xfa_source):
                if item == "datasets":
                    source_idx = i + 1
                    break
            for i, item in enumerate(xfa_target):
                if item == "datasets":
                    target_idx = i + 1
                    break

            if source_idx != -1 and target_idx != -1:
                data_packet = xfa_source[source_idx]
                raw_xml = data_packet.read_bytes()
                xml_str = raw_xml.decode("utf-8", errors="ignore")

                # 1. Fix English
                if "TextfeldEnglisch" in xml_str:
                    xml_str = xml_str.replace("TextfeldEnglisch", "TextfeldÄsthetik")

                # 2. Fix Year
                if s_year:
                    pattern = r"(<Schuljahr\s*[^>]*>)(.*?)(</Schuljahr\s*>)"
                    if re.search(pattern, xml_str, re.DOTALL):
                        xml_str = re.sub(
                            pattern, f"\\g<1>{s_year}\\g<3>", xml_str, flags=re.DOTALL
                        )
                    elif "<Schuljahr" in xml_str:
                        xml_str = re.sub(
                            r"<Schuljahr\s*/>",
                            f"<Schuljahr>{s_year}</Schuljahr>",
                            xml_str,
                        )

                # 3. Fix Class
                if s_class:
                    pattern = r"(<Klasse\s*[^>]*>)(.*?)(</Klasse\s*>)"
                    if re.search(pattern, xml_str, re.DOTALL):
                        xml_str = re.sub(
                            pattern, f"\\g<1>{s_class}\\g<3>", xml_str, flags=re.DOTALL
                        )
                    elif "<Klasse" in xml_str:
                        xml_str = re.sub(
                            r"<Klasse\s*/>", f"<Klasse>{s_class}</Klasse>", xml_str
                        )

                # 4. Fix Date
                if s_date:
                    pattern_date = (
                        r"(<DatumsUhrzeitfeld1\s*[^>]*>)(.*?)(</DatumsUhrzeitfeld1\s*>)"
                    )
                    if re.search(pattern_date, xml_str, re.DOTALL):
                        xml_str = re.sub(
                            pattern_date,
                            f"\\g<1>{s_date}\\g<3>",
                            xml_str,
                            flags=re.DOTALL,
                        )
                    elif "<DatumsUhrzeitfeld1" in xml_str:
                        xml_str = re.sub(
                            r"<DatumsUhrzeitfeld1\s*/>",
                            f"<DatumsUhrzeitfeld1>{s_date}</DatumsUhrzeitfeld1>",
                            xml_str,
                        )

                new_data_stream = pdf_target.make_stream(xml_str.encode("utf-8"))
                xfa_target[target_idx] = new_data_stream

                if "/Perms" in pdf_target.Root:
                    del pdf_target.Root["/Perms"]
                if (
                    "/AcroForm" in pdf_target.Root
                    and "/SigFlags" in pdf_target.Root.AcroForm
                ):
                    del pdf_target.Root.AcroForm["/SigFlags"]

                pdf_target.save(output_path)
                return True
            else:
                self.log(f"ERROR: No datasets found in {os.path.basename(source_path)}")
                return False

        except Exception as e:
            self.log(f"CRITICAL ERROR: {str(e)}")
            return False

    def run_process(self):
        src = self.source_folder.get()
        tmpl = self.template_file.get()
        out = self.target_folder.get()
        s_year = self.school_year.get()
        s_class = self.class_name.get()
        s_date = self.sign_date.get()

        if not src or not tmpl or not out:
            self.log("ERROR: Please select all three paths!")
            return

        files = [f for f in os.listdir(src) if f.lower().endswith(".pdf")]
        if not files:
            self.log("ERROR: No PDF files found in source folder.")
            return

        self.start_btn.configure(state="disabled", text="Verarbeite...")
        self.log("-" * 40)
        self.log(f"Starting... (Year: {s_year}, Class: {s_class}, Date: {s_date})")

        start_time = time.time()

        count = 0
        for filename in files:
            source_full = os.path.join(src, filename)
            output_full = os.path.join(out, filename)

            self.log(f"Processing: {filename}...")
            if self.transfer_data(
                source_full, tmpl, output_full, s_year, s_class, s_date
            ):
                self.log(" -> Done.")
                count += 1
            else:
                self.log(" -> FAILED.")

        end_time = time.time()
        duration = end_time - start_time

        self.log("-" * 40)
        self.log(f"COMPLETED in {duration:.2f} seconds.")
        self.start_btn.configure(state="normal", text="ZEUGNISSE KOPIEREN")


if __name__ == "__main__":
    app = ZeugnisApp()
    app.mainloop()

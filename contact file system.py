import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
import pandas as pd
import pywhatkit as kit
import threading
import time
import random
import os

class WhatsAppCRMApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk WhatsApp Marketing Software")
        self.root.geometry("950x800")
        
        # Cross-Platform Database Path setup (Saves to user's home directory)
        self.db_path = os.path.join(os.path.expanduser("~"), "whatsapp_crm_contacts.db")
        self.init_db()
        
        # UI Styling Elements
        style = ttk.Style()
        style.theme_use('clam')
        
        # -------------------------------------------------------------
        # PANEL 1: CONTACT MANAGEMENT (ADD MANUAL & EXCEL IMPORT)
        # -------------------------------------------------------------
        management_frame = ttk.LabelFrame(root, text=" Feature 1: Add & Import Contacts ", padding=12)
        management_frame.pack(fill="x", padx=15, pady=8)
        
        # Row 1: Manual Input Fields
        ttk.Label(management_frame, text="Name:").grid(row=0, column=0, sticky="w", padx=5)
        self.name_entry = ttk.Entry(management_frame, width=15)
        self.name_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(management_frame, text="Phone (+91 Helper):").grid(row=0, column=2, sticky="w", padx=5)
        self.phone_entry = ttk.Entry(management_frame, width=15)
        self.phone_entry.grid(row=0, column=3, padx=5, pady=5)
        self.phone_entry.insert(0, "+91") 
        
        ttk.Label(management_frame, text="Group/Tag:").grid(row=0, column=4, sticky="w", padx=5)
        self.tag_entry = ttk.Entry(management_frame, width=12)
        self.tag_entry.grid(row=0, column=5, padx=5, pady=5)
        
        add_btn = ttk.Button(management_frame, text="Save Contact", command=self.save_contact)
        add_btn.grid(row=0, column=6, padx=10)

        # Separator Line
        sep = ttk.Separator(management_frame, orient='vertical')
        sep.grid(row=0, column=7, rowspan=2, sticky='ns', padx=10)

        # Row 2: Excel Import Tools
        excel_label = ttk.Label(management_frame, text="Bulk Data Upload:", font=("Arial", 10, "bold"))
        excel_label.grid(row=0, column=8, padx=5, sticky="w")
        
        excel_btn = ttk.Button(management_frame, text="📂 Import Excel Sheet", command=self.import_excel_file)
        excel_btn.grid(row=0, column=9, padx=5, pady=5, ipadx=5)

        # -------------------------------------------------------------
        # PANEL 2: CONTACT DATABASE VIEW (CENTER JUSTIFIED)
        # -------------------------------------------------------------
        view_frame = ttk.LabelFrame(root, text=" Feature 3: Saved Contact Database ", padding=10)
        view_frame.pack(fill="both", expand=True, padx=15, pady=5)
        
        scroll_y = ttk.Scrollbar(view_frame, orient="vertical")
        scroll_y.pack(side="right", fill="y")
        
        self.contact_table = ttk.Treeview(view_frame, columns=("ID", "Name", "Phone", "Group Tag"), show="headings", yscrollcommand=scroll_y.set)
        scroll_y.config(command=self.contact_table.yview)
        
        # Center-justified Headings and Rows
        self.contact_table.heading("ID", text="Database ID", anchor="center")
        self.contact_table.heading("Name", text="Full Name", anchor="center")
        self.contact_table.heading("Phone", text="WhatsApp Phone Number", anchor="center")
        self.contact_table.heading("Group Tag", text="Assigned Group/Tag", anchor="center")
        
        self.contact_table.column("ID", width=80, anchor="center")
        self.contact_table.column("Name", width=220, anchor="center") 
        self.contact_table.column("Phone", width=200, anchor="center")
        self.contact_table.column("Group Tag", width=150, anchor="center")
        self.contact_table.pack(fill="both", expand=True, side="left")
        
        table_controls = ttk.Frame(view_frame)
        table_controls.pack(side="right", fill="y", padx=5)
        
        refresh_btn = ttk.Button(table_controls, text="🔄 Refresh\nList", command=self.refresh_table_view)
        refresh_btn.pack(pady=5, fill="x")
        
        delete_btn = ttk.Button(table_controls, text="❌ Delete\nSelected", command=self.delete_selected_contact)
        delete_btn.pack(pady=5, fill="x")

        # -------------------------------------------------------------
        # PANEL 3: FILTER & MESSAGE CAMPAIGN SETUP
        # -------------------------------------------------------------
        campaign_frame = ttk.LabelFrame(root, text=" Feature 2: Set Filter & Compose Broadcast ", padding=12)
        campaign_frame.pack(fill="x", padx=15, pady=5)
        
        filter_layout = ttk.Frame(campaign_frame)
        filter_layout.pack(fill="x", pady=5)
        ttk.Label(filter_layout, text="🎯 Target Group Tag:").pack(side="left", padx=5)
        self.filter_entry = ttk.Entry(filter_layout, width=20)
        self.filter_entry.pack(side="left", padx=5)
        
        check_btn = ttk.Button(filter_layout, text="Count Audience Size", command=self.check_count)
        check_btn.pack(side="left", padx=10)
        self.count_label = ttk.Label(filter_layout, text="0 contacts targeted", font=("Arial", 9, "italic"))
        self.count_label.pack(side="left", padx=5)

        ttk.Label(campaign_frame, text="💬 Input Message Body:").pack(anchor="w", padx=5, pady=(5,2))
        self.msg_text = tk.Text(campaign_frame, height=4, wrap="word")
        self.msg_text.pack(fill="x", padx=5, pady=5)
        
        media_layout = ttk.Frame(campaign_frame)
        media_layout.pack(fill="x", pady=5)
        ttk.Label(media_layout, text="🖼️ Attachment (Poster / Video Path):").pack(side="left", padx=5)
        self.media_entry = ttk.Entry(media_layout, width=50)
        self.media_entry.pack(side="left", padx=5, fill="x", expand=True)
        browse_btn = ttk.Button(media_layout, text="Browse...", command=self.browse_file)
        browse_btn.pack(side="left", padx=5)

        # -------------------------------------------------------------
        # PANEL 4: STATUS TRACKER & ENGINE LAUNCH
        # -------------------------------------------------------------
        status_frame = ttk.LabelFrame(root, text=" Execution & Live Logs ", padding=10)
        status_frame.pack(fill="x", padx=15, pady=8)
        
        self.log_box = tk.Text(status_frame, height=5, bg="#1e1e1e", fg="#00ff00", font=("Consolas", 10))
        self.log_box.pack(fill="x", side="left", expand=True, padx=5)
        
        send_btn = ttk.Button(status_frame, text="🚀 Launch\nBroadcast", command=self.start_broadcast_thread)
        send_btn.pack(side="right", padx=10, fill="y", ipady=10)
        
        self.refresh_table_view()

    def init_db(self):
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS recipients (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                phone TEXT UNIQUE NOT NULL,
                tag TEXT NOT NULL
            )
        """)
        conn.commit()
        conn.close()

    def refresh_table_view(self):
        for item in self.contact_table.get_children():
            self.contact_table.delete(item)
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT id, name, phone, tag FROM recipients ORDER BY id DESC")
        rows = cursor.fetchall()
        conn.close()
        for row in rows:
            self.contact_table.insert("", tk.END, values=row)

    def save_contact(self):
        name = self.name_entry.get().strip()
        phone = self.phone_entry.get().strip()
        tag = self.tag_entry.get().strip().lower()
        
        if not name or not phone or not tag:
            messagebox.showwarning("Incomplete Fields", "Please populate Name, Phone, and Tag fields.")
            return
            
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("INSERT INTO recipients (name, phone, tag) VALUES (?, ?, ?)", (name, phone, tag))
            conn.commit()
            conn.close()
            self.name_entry.delete(0, tk.END)
            self.tag_entry.delete(0, tk.END)
            self.refresh_table_view()
        except sqlite3.IntegrityError:
            messagebox.showerror("Error", f"Phone number {phone} already exists!")

    def import_excel_file(self):
        file_path = filedialog.askopenfilename(title="Select Contacts Excel File", filetypes=[("Excel Files", "*.xlsx *.xls")])
        if not file_path:
            return
        try:
            df = pd.read_excel(file_path)
            df.columns = [str(col).strip().lower() for col in df.columns]
            required_cols = ['name', 'number', 'category']
            if not all(col in df.columns for col in required_cols):
                messagebox.showerror("Format Error", "Excel columns must be named exactly:\n'name', 'number', and 'category'")
                return
                
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            success_count, duplicate_count = 0, 0
            
            for index, row in df.iterrows():
                raw_name = str(row['name']).strip()
                raw_num = str(row['number']).strip()
                raw_tag = str(row['category']).strip().lower()
                
                if raw_num.endswith('.0'):
                    raw_num = raw_num[:-2]
                if not raw_name or not raw_num or raw_num == "nan":
                    continue
                
                if not raw_num.startswith('+'):
                    if raw_num.startswith('91') and len(raw_num) > 10:
                        formatted_num = "+" + raw_num
                    else:
                        formatted_num = "+91" + raw_num
                else:
                    formatted_num = raw_num
                
                try:
                    cursor.execute("INSERT INTO recipients (name, phone, tag) VALUES (?, ?, ?)", (raw_name, formatted_num, raw_tag))
                    success_count += 1
                except sqlite3.IntegrityError:
                    duplicate_count += 1
                    continue
            conn.commit()
            conn.close()
            self.refresh_table_view()
            messagebox.showinfo("Import Success", f"Imported {success_count} contacts.\nSkipped {duplicate_count} duplicates.")
        except Exception as error:
            messagebox.showerror("File Error", f"Could not read Excel file:\n{error}")

    def delete_selected_contact(self):
        selected_item = self.contact_table.selection()
        if not selected_item:
            messagebox.showwarning("No Selection", "Please select a row first.")
            return
        confirm = messagebox.askyesno("Confirm Delete", "Permanently delete this contact?")
        if confirm:
            item_data = self.contact_table.item(selected_item)
            record_id = item_data['values'][0]
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute("DELETE FROM recipients WHERE id = ?", (record_id,))
            conn.commit()
            conn.close()
            self.refresh_table_view()

    def check_count(self):
        tag = self.filter_entry.get().strip().lower()
        if not tag:
            self.count_label.config(text="Please type a tag name.")
            return 0
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM recipients WHERE tag = ?", (tag,))
        count = cursor.fetchone()[0]
        conn.close()
        self.count_label.config(text=f"📊 Target Size: {count} contacts.")
        return count

    def browse_file(self):
        filename = filedialog.askopenfilename(title="Select Media File", filetypes=[("All Files", "*.*")])
        if filename:
            self.media_entry.delete(0, tk.END)
            self.media_entry.insert(0, filename)

    def write_log(self, text):
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)

    def start_broadcast_thread(self):
        thread = threading.Thread(target=self.run_broadcast_engine, daemon=True)
        thread.start()

    def run_broadcast_engine(self):
        tag = self.filter_entry.get().strip().lower()
        msg = self.msg_text.get("1.0", tk.END).strip()
        media = self.media_entry.get().strip()
        
        if not tag or not msg:
            messagebox.showerror("Execution Aborted", "Target Group and Message are required.")
            return
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        cursor.execute("SELECT name, phone FROM recipients WHERE tag = ?", (tag,))
        rows = cursor.fetchall()
        conn.close()
        
        if not rows:
            self.write_log(f"⚠ Aborted: No entries match '{tag}'")
            return
            
        self.write_log(f"🏁 Processing {len(rows)} recipients...")
        
        for name, phone in rows:
            personalized_msg = f"Hey {name},\n\n{msg}"
            self.write_log(f"📤 Sending payload to {name} ({phone})...")
            
            try:
                if media and os.path.exists(media):
                    kit.sendwhats_image(receiver=phone, img_path=media, caption=personalized_msg, wait_time=20, tab_close=True, close_time=3)
                else:
                    kit.sendwhatmsg_instantly(phone_no=phone, message=personalized_msg, wait_time=15, tab_close=True, close_time=3)
                
                self.write_log(f"✅ Dispatched to {name}!")
                delay_timer = 20 + random.randint(5, 12)
                self.write_log(f"⏳ Sleeping {delay_timer} seconds for safety...")
                time.sleep(delay_timer)
                
            except Exception as failure:
                self.write_log(f"❌ Error for {name}: {failure}")
                continue
        self.write_log("🏁 Campaign transmission finished.")

if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppCRMApp(root)
    root.mainloop()

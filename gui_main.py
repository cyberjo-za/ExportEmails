import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import aspose.email as ae
import json
import os
from datetime import datetime

SESSION_FILE = "export_session.json"

class IMAPExporterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("IMAP to PST Exporter")
        self.root.geometry("600x700")
        self.root.resizable(False, False)
        
        # Dark mode colors
        self.bg_color = "#1e1e1e"
        self.fg_color = "#e0e0e0"
        self.accent_color = "#0d7377"
        self.entry_bg = "#2d2d2d"
        self.border_color = "#404040"
        
        self.root.configure(bg=self.bg_color)
        
        # Configure styles
        self.setup_styles()
        
        # State variables
        self.exporting = False
        self.export_thread = None
        self.output_file = tk.StringVar()
        
        # Build GUI
        self.build_gui()
        
    def setup_styles(self):
        """Configure ttk styles for dark mode."""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure colors
        style.configure('TLabel', background=self.bg_color, foreground=self.fg_color)
        style.configure('TFrame', background=self.bg_color)
        style.configure('TEntry', fieldbackground=self.entry_bg, foreground=self.fg_color, 
                       borderwidth=1, relief='solid')
        style.configure('TButton', background=self.accent_color, foreground=self.fg_color)
        style.map('TButton', background=[('active', '#0a5a63')])
        style.configure('TProgressbar', background=self.accent_color, troughcolor=self.entry_bg)
        
    def build_gui(self):
        """Build the GUI layout."""
        # Main frame
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Title
        title_label = ttk.Label(main_frame, text="IMAP to PST Exporter", 
                               font=("Arial", 18, "bold"))
        title_label.pack(pady=(0, 20))
        
        # Server details frame
        details_frame = ttk.Frame(main_frame)
        details_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(details_frame, text="IMAP Server:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        self.server_entry = ttk.Entry(details_frame, width=40)
        self.server_entry.insert(0, "mail.juliesproperties.co.za")
        self.server_entry.grid(row=0, column=1, sticky=tk.EW, padx=(10, 0))
        
        ttk.Label(details_frame, text="Port:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        self.port_entry = ttk.Entry(details_frame, width=40)
        self.port_entry.insert(0, "993")
        self.port_entry.grid(row=1, column=1, sticky=tk.EW, padx=(10, 0))
        
        ttk.Label(details_frame, text="Email:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        self.email_entry = ttk.Entry(details_frame, width=40)
        self.email_entry.insert(0, "craig@juliesproperties.co.za")
        self.email_entry.grid(row=2, column=1, sticky=tk.EW, padx=(10, 0))
        
        ttk.Label(details_frame, text="Password:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(details_frame, width=40, show="•")
        self.password_entry.grid(row=3, column=1, sticky=tk.EW, padx=(10, 0))
        
        details_frame.columnconfigure(1, weight=1)
        
        # Output file frame
        file_frame = ttk.Frame(main_frame)
        file_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(file_frame, text="Output File:", font=("Arial", 10)).pack(side=tk.LEFT)
        self.file_label = ttk.Label(file_frame, text="Not selected", foreground="#888888")
        self.file_label.pack(side=tk.LEFT, padx=(10, 10), fill=tk.X, expand=True)
        
        file_btn = ttk.Button(file_frame, text="Browse", command=self.select_file)
        file_btn.pack(side=tk.RIGHT)
        
        # Export button
        export_btn = ttk.Button(main_frame, text="Export", command=self.start_export)
        export_btn.pack(fill=tk.X, pady=(0, 15))
        
        # Progress bar
        self.progress = ttk.Progressbar(main_frame, mode='determinate', length=400)
        self.progress.pack(fill=tk.X, pady=(0, 10))
        
        # Progress text
        self.progress_label = ttk.Label(main_frame, text="Ready", font=("Arial", 9))
        self.progress_label.pack(fill=tk.X, pady=(0, 15))
        
        # Status log frame
        log_frame = ttk.Frame(main_frame)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        ttk.Label(log_frame, text="Status Log:", font=("Arial", 10)).pack(anchor=tk.W)
        
        # Text widget with scrollbar
        scrollbar = ttk.Scrollbar(log_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.log_text = tk.Text(log_frame, height=12, width=60, bg=self.entry_bg, 
                               fg=self.fg_color, insertbackground=self.fg_color,
                               relief='solid', borderwidth=1, wrap=tk.WORD,
                               yscrollcommand=scrollbar.set)
        self.log_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.log_text.yview)
        self.log_text.config(state=tk.DISABLED)
        
    def select_file(self):
        """Open file dialog to select output PST file."""
        file = filedialog.asksaveasfilename(
            defaultextension=".pst",
            filetypes=[("PST Files", "*.pst"), ("All Files", "*.*")]
        )
        if file:
            self.output_file.set(file)
            self.file_label.config(text=os.path.basename(file), foreground=self.fg_color)
    
    def log(self, message):
        """Add message to log text widget."""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.root.update_idletasks()
    
    def start_export(self):
        """Start export in separate thread."""
        if self.exporting:
            messagebox.showwarning("Warning", "Export already in progress")
            return
        
        # Validate inputs
        server = self.server_entry.get().strip()
        port = self.port_entry.get().strip()
        email = self.email_entry.get().strip()
        password = self.password_entry.get()
        output_file = self.output_file.get()
        
        if not all([server, port, email, password, output_file]):
            messagebox.showerror("Error", "Please fill in all fields and select output file")
            return
        
        try:
            port = int(port)
        except ValueError:
            messagebox.showerror("Error", "Port must be a number")
            return
        
        self.exporting = True
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)
        self.progress.config(value=0)
        self.progress_label.config(text="Starting export...")
        
        # Start export in background thread
        self.export_thread = threading.Thread(
            target=self.export_emails,
            args=(server, port, email, password, output_file),
            daemon=True
        )
        self.export_thread.start()
    
    def export_emails(self, server, port, email, password, output_file):
        """Export emails from IMAP to PST."""
        try:
            self.log(f"Connecting to {server}...")
            
            # Create IMAP client
            imap_client = ae.clients.imap.ImapClient()
            imap_client.host = server
            imap_client.port = port
            imap_client.username = email
            imap_client.password = password
            imap_client.security_options = ae.clients.SecurityOptions.AUTO
            imap_client.timeout = 300000
            
            self.log("✓ Connected successfully")
            
            # Get mailbox info and all folders
            mailbox_info = imap_client.mailbox_info
            folders = imap_client.list_folders()
            self.log(f"✓ Found {len(folders)} folder(s)")
            
            # Create PST file
            if os.path.exists(output_file):
                os.remove(output_file)
            
            # Create PST
            pst = ae.storage.pst.PersonalStorage.create(output_file, ae.storage.pst.FileFormatVersion.UNICODE)
            
            self.log("✓ Starting export...\n")
            
            total_all_messages = 0
            processed = 0
            
            # First pass: count total messages
            for folder in folders:
                try:
                    imap_client.select_folder(folder.name)
                    messages = imap_client.list_messages()
                    total_all_messages += len(messages)
                except:
                    continue
            
            self.log(f"✓ Total messages to export: {total_all_messages}\n")
            
            # Second pass: export all folders
            for folder in folders:
                try:
                    folder_name = folder.name
                    imap_client.select_folder(folder_name)
                    messages = imap_client.list_messages()
                    
                    if len(messages) == 0:
                        continue
                    
                    self.log(f"\nExporting folder: {folder_name} ({len(messages)} messages)")
                    
                    # Create folder in PST
                    pst_folder = pst.root_folder.add_sub_folder(folder_name)
                    
                    # Export each message in this folder
                    for i, msg_info in enumerate(messages):
                        try:
                            email_msg = imap_client.fetch_message(msg_info.unique_id)
                            mapi_msg = ae.mapi.MapiMessage.from_mail_message(email_msg)
                            pst_folder.add_message(mapi_msg)
                            
                            processed += 1
                            progress = (processed / total_all_messages) * 100 if total_all_messages > 0 else 0
                            
                            self.progress.config(value=progress)
                            
                            if (i + 1) % 10 == 0 or (i + 1) == len(messages):
                                subject = msg_info.subject[:40] if msg_info.subject else "(No subject)"
                                self.progress_label.config(
                                    text=f"Progress: [{processed}/{total_all_messages}] ({progress:.1f}%)"
                                )
                                self.log(f"  [{processed}/{total_all_messages}] {subject}")
                            
                            self.root.update_idletasks()
                            
                        except Exception as e:
                            self.log(f"  ✗ Failed to export message: {str(e)}")
                            continue
                
                except Exception as e:
                    self.log(f"✗ Failed to export folder '{folder_name}': {str(e)}")
                    continue
            
            # PST is automatically saved when created
            
            self.log(f"\n✓ Successfully exported {processed} emails from {len(folders)} folders")
            self.log(f"✓ Saved to: {output_file}")
            self.progress_label.config(text="Export completed successfully!")
            messagebox.showinfo("Success", f"Exported {processed} emails to {output_file}")
            
        except Exception as e:
            error_msg = f"Export failed: {str(e)}"
            self.log(f"✗ {error_msg}")
            self.progress_label.config(text="Export failed")
            messagebox.showerror("Error", error_msg)
        
        finally:
            self.exporting = False

if __name__ == "__main__":
    root = tk.Tk()
    app = IMAPExporterGUI(root)
    root.mainloop()
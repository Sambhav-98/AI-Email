import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import imaplib
import email
from email.header import decode_header
import ollama
from datetime import datetime, timedelta
import threading
from tkcalendar import DateEntry  # Requires pip install tkcalendar

class EmailAIAssistant:
    def __init__(self, root):
        self.root = root
        self.root.title("Email AI Assistant")
        self.root.geometry("1100x1200")  # Increased height to accommodate new prompt bar
        
        # Configuration variables
        self.email = "sambhavkc30853@gmail.com"
        self.password = "nige ysvx evdm vjez"
        self.imap_server = "imap.gmail.com"
        self.model_name = "llama3.2:3b"
        self.max_emails = 10
        self.email_data = {}  # Dictionary to store email data
        self.current_email_id = None  # Track currently selected email
        
        # Create UI
        self.create_widgets()
        
        # Initialize Ollama
        threading.Thread(target=self.initialize_ollama, daemon=True).start()
    
    def create_widgets(self):
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Left panel - Controls
        control_frame = ttk.LabelFrame(main_frame, text="Controls", padding="10")
        control_frame.pack(side=tk.LEFT, fill=tk.Y, padx=5, pady=5)
        
        # Connection settings
        ttk.Label(control_frame, text="IMAP Server:").grid(row=0, column=0, sticky=tk.W)
        self.server_entry = ttk.Entry(control_frame)
        self.server_entry.grid(row=0, column=1, pady=5)
        self.server_entry.insert(0, self.imap_server)
        
        ttk.Label(control_frame, text="Email:").grid(row=1, column=0, sticky=tk.W)
        self.email_entry = ttk.Entry(control_frame)
        self.email_entry.grid(row=1, column=1, pady=5)
        self.email_entry.insert(0, self.email)
        
        ttk.Label(control_frame, text="Password:").grid(row=2, column=0, sticky=tk.W)
        self.password_entry = ttk.Entry(control_frame, show="*")
        self.password_entry.grid(row=2, column=1, pady=5)
        self.password_entry.insert(0, self.password)
        
        ttk.Label(control_frame, text="Max Emails:").grid(row=3, column=0, sticky=tk.W)
        self.max_emails_spin = ttk.Spinbox(control_frame, from_=1, to=50, width=5)
        self.max_emails_spin.grid(row=3, column=1, pady=5)
        self.max_emails_spin.set(self.max_emails)
        
        # Date range filter
        ttk.Label(control_frame, text="Date Range Filter:").grid(row=4, column=0, columnspan=2, pady=(10,0), sticky=tk.W)
        
        ttk.Label(control_frame, text="From:").grid(row=5, column=0, sticky=tk.W)
        self.date_from = DateEntry(control_frame)
        self.date_from.grid(row=5, column=1, pady=5)
        
        ttk.Label(control_frame, text="To:").grid(row=6, column=0, sticky=tk.W)
        self.date_to = DateEntry(control_frame)
        self.date_to.grid(row=6, column=1, pady=5)
        
        # Search options
        self.search_option = tk.StringVar(value="unseen")
        ttk.Radiobutton(control_frame, text="Unread Only", variable=self.search_option, value="unseen").grid(row=7, column=0, columnspan=2, sticky=tk.W)
        ttk.Radiobutton(control_frame, text="All Emails", variable=self.search_option, value="all").grid(row=8, column=0, columnspan=2, sticky=tk.W)
        
        # Action buttons
        ttk.Button(control_frame, text="Fetch Emails", command=self.fetch_emails_thread).grid(row=9, column=0, columnspan=2, pady=10)
        ttk.Button(control_frame, text="Clear Results", command=self.clear_results).grid(row=10, column=0, columnspan=2, pady=5)
        
        # Right panel - Results
        result_frame = ttk.LabelFrame(main_frame, text="Email Processing Results", padding="10")
        result_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Progress bar frame
        self.progress_frame = ttk.Frame(result_frame)
        self.progress_frame.pack(fill=tk.X, pady=5)
        
        self.progress_label = ttk.Label(self.progress_frame, text="Ready")
        self.progress_label.pack(side=tk.TOP, fill=tk.X)
        
        self.progress_bar = ttk.Progressbar(self.progress_frame, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X)
        
        # Email list treeview
        self.email_tree = ttk.Treeview(result_frame, columns=("from", "subject", "date"), show="headings")
        self.email_tree.heading("from", text="From")
        self.email_tree.heading("subject", text="Subject")
        self.email_tree.heading("date", text="Date")
        self.email_tree.column("from", width=150)
        self.email_tree.column("subject", width=250)
        self.email_tree.column("date", width=120)
        self.email_tree.pack(fill=tk.BOTH, expand=True, pady=5)
        self.email_tree.bind("<<TreeviewSelect>>", self.show_email_details)
        
        # Details frame
        detail_frame = ttk.Frame(result_frame)
        detail_frame.pack(fill=tk.BOTH, expand=True)
        
        # Email details tabs
        self.notebook = ttk.Notebook(detail_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Original email tab
        original_tab = ttk.Frame(self.notebook)
        self.original_text = scrolledtext.ScrolledText(original_tab, wrap=tk.WORD)
        self.original_text.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(original_tab, text="Original")
        
        # Summary tab
        summary_tab = ttk.Frame(self.notebook)
        self.summary_text = scrolledtext.ScrolledText(summary_tab, wrap=tk.WORD)
        self.summary_text.pack(fill=tk.BOTH, expand=True)
        self.notebook.add(summary_tab, text="Summary")
        
        # Reply tab
        reply_tab = ttk.Frame(self.notebook)
        reply_frame = ttk.Frame(reply_tab)
        reply_frame.pack(fill=tk.BOTH, expand=True)
        
        self.reply_text = scrolledtext.ScrolledText(reply_frame, wrap=tk.WORD)
        self.reply_text.pack(fill=tk.BOTH, expand=True)
        
        # Prompt bar for reply modifications
        prompt_frame = ttk.Frame(reply_frame)
        prompt_frame.pack(fill=tk.X, pady=(5, 0))
        
        ttk.Label(prompt_frame, text="Modify Reply:").pack(side=tk.LEFT, padx=(0, 5))
        self.prompt_entry = ttk.Entry(prompt_frame)
        self.prompt_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.prompt_entry.bind("<Return>", self.modify_reply)
        
        ttk.Button(prompt_frame, text="Update", command=self.modify_reply).pack(side=tk.LEFT)
        
        self.notebook.add(reply_tab, text="Suggested Reply")
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN)
        self.status_bar.pack(fill=tk.X)
    
    def initialize_ollama(self):
        """Ensure the model is available locally"""
        try:
            self.update_progress("Checking Ollama model...", 0)
            ollama.show(self.model_name)
            self.update_progress("Ollama model ready", 100)
        except ollama.ResponseError:
            self.update_progress(f"Downloading {self.model_name}...", 0)
            stream = ollama.pull(self.model_name, stream=True)
            
            # Track download progress
            total_size = 0
            downloaded = 0
            for progress in stream:
                if 'total' in progress:
                    total_size = progress['total']
                if 'completed' in progress:
                    downloaded = progress['completed']
                
                if total_size > 0:
                    percent = int((downloaded / total_size) * 100)
                    self.update_progress(f"Downloading {self.model_name}... {percent}%", percent)
            
            self.update_progress("Ollama model downloaded and ready", 100)
    
    def update_progress(self, message, value):
        """Update progress bar and label"""
        self.progress_label.config(text=message)
        self.progress_bar['value'] = value
        self.root.update_idletasks()
    
    def fetch_emails_thread(self):
        """Start email fetching in a separate thread"""
        self.update_progress("Fetching emails...", 0)
        threading.Thread(target=self.fetch_and_process_emails, daemon=True).start()
    
    def fetch_and_process_emails(self):
        """Fetch and process emails"""
        try:
            # Update config from UI
            self.imap_server = self.server_entry.get()
            self.email = self.email_entry.get()
            self.password = self.password_entry.get()
            self.max_emails = int(self.max_emails_spin.get())
            
            # Clear previous results
            self.clear_results()
            
            # Fetch emails
            emails = self.fetch_unread_emails()
            
            if not emails:
                self.update_progress("No emails found matching criteria", 100)
                return
            
            # Process emails
            total_emails = len(emails)
            for idx, email in enumerate(emails, 1):
                progress = int((idx / total_emails) * 100)
                self.update_progress(f"Processing email {idx} of {total_emails}...", progress)
                
                summary = self.generate_summary(email)
                reply = self.generate_reply(email)
                
                # Add to treeview
                item_id = self.email_tree.insert("", tk.END, values=(
                    email["from"],
                    email["subject"],
                    email["date"]
                ))
                
                # Store full data in our dictionary
                self.email_data[item_id] = {
                    "original": email,
                    "summary": summary,
                    "reply": reply
                }
                
            self.update_progress(f"Processed {total_emails} emails", 100)
            
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.update_progress("Error processing emails", 0)
    
    def fetch_unread_emails(self):
        """Connect to IMAP server and fetch emails based on filters"""
        mail = imaplib.IMAP4_SSL(self.imap_server)
        mail.login(self.email, self.password)
        mail.select("inbox")
        
        # Build search query based on options
        search_query = []
        
        # Add date range filter
        date_from = self.date_from.get_date()
        date_to = self.date_to.get_date() + timedelta(days=1)  # Include entire end day
        
        # Format dates for IMAP search (DD-MMM-YYYY)
        date_from_str = date_from.strftime("%d-%b-%Y")
        date_to_str = date_to.strftime("%d-%b-%Y")
        
        search_query.append(f'(SINCE "{date_from_str}" BEFORE "{date_to_str}")')
        
        # Add unread filter if selected
        if self.search_option.get() == "unseen":
            search_query.append("UNSEEN")
        
        # Combine search criteria
        search_criteria = " ".join(search_query)
        
        self.update_progress(f"Searching with criteria: {search_criteria}", 10)
        
        status, messages = mail.search(None, search_criteria)
        if status != "OK":
            return []
        
        email_ids = messages[0].split()
        emails = []
        
        # Only process up to max_emails
        total_to_process = min(len(email_ids), self.max_emails)
        for idx, email_id in enumerate(email_ids[:self.max_emails]):
            progress = 10 + int((idx / total_to_process) * 40)  # 10-50% for fetching
            self.update_progress(f"Fetching email {idx+1} of {total_to_process}...", progress)
            
            _, msg_data = mail.fetch(email_id, "(RFC822)")
            raw_email = msg_data[0][1]
            email_message = email.message_from_bytes(raw_email)
            emails.append(self.parse_email(email_message))
        
        mail.close()
        mail.logout()
        return emails
    
    def parse_email(self, msg):
        """Extract relevant information from email"""
        subject, encoding = decode_header(msg["subject"])[0]
        if isinstance(subject, bytes):
            subject = subject.decode(encoding or "utf-8")
        
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    body = part.get_payload(decode=True).decode(errors="replace")
                    break
        else:
            body = msg.get_payload(decode=True).decode(errors="replace")
        
        return {
            "subject": subject,
            "from": msg["from"],
            "date": msg["date"],
            "body": body
        }
    
    def generate_summary(self, email_content):
        """Use Ollama to generate email summary"""
        prompt = f"""
        Summarize this email in 2-3 bullet points:
        
        Subject: {email_content['subject']}
        From: {email_content['from']}
        Body: {email_content['body'][:2000]}
        """
        
        response = ollama.chat(
            model=self.model_name,
            messages=[{"role": "user", "content": prompt}]
        )
        return response["message"]["content"]
    
    def generate_reply(self, email_content):
        """Use Ollama to draft a reply"""
        prompt = f"""
        Draft a professional response to this email. 
        Keep it concise (2-3 sentences max).
        
        Original email:
        Subject: {email_content['subject']}
        From: {email_content['from']}
        Body: {email_content['body'][:2000]}
        """
        
        response = ollama.chat(
            model=self.model_name,
            messages=[{"role": "user", "content": prompt}]
        )
        return response["message"]["content"]
    
    def show_email_details(self, event):
        """Show details of selected email"""
        selected_item = self.email_tree.focus()
        if not selected_item or selected_item not in self.email_data:
            return
        
        self.current_email_id = selected_item
        email_data = self.email_data[selected_item]
        
        # Clear previous content
        self.original_text.delete(1.0, tk.END)
        self.summary_text.delete(1.0, tk.END)
        self.reply_text.delete(1.0, tk.END)
        
        # Insert new content
        original_content = f"From: {email_data['original']['from']}\n"
        original_content += f"Date: {email_data['original']['date']}\n"
        original_content += f"Subject: {email_data['original']['subject']}\n\n"
        original_content += email_data['original']['body']
        
        self.original_text.insert(tk.END, original_content)
        self.summary_text.insert(tk.END, email_data['summary'])
        self.reply_text.insert(tk.END, email_data['reply'])
    
    def modify_reply(self, event=None):
        """Modify the reply based on user prompt"""
        if not self.current_email_id:
            messagebox.showwarning("Warning", "No email selected")
            return
        
        prompt = self.prompt_entry.get()
        if not prompt:
            messagebox.showwarning("Warning", "Please enter a modification prompt")
            return
        
        original_email = self.email_data[self.current_email_id]['original']
        current_reply = self.reply_text.get(1.0, tk.END)
        
        modification_prompt = f"""
        Here's the original email:
        Subject: {original_email['subject']}
        From: {original_email['from']}
        Body: {original_email['body'][:2000]}
        
        Here's the current draft reply:
        {current_reply}
        
        Please modify the reply according to these instructions: {prompt}
        """
        
        try:
            self.update_progress("Generating modified reply...", 50)
            response = ollama.chat(
                model=self.model_name,
                messages=[{"role": "user", "content": modification_prompt}]
            )
            
            modified_reply = response["message"]["content"]
            
            # Update the reply text and stored data
            self.reply_text.delete(1.0, tk.END)
            self.reply_text.insert(tk.END, modified_reply)
            self.email_data[self.current_email_id]['reply'] = modified_reply
            
            self.update_progress("Reply modified successfully", 100)
            self.prompt_entry.delete(0, tk.END)  # Clear the prompt entry
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to modify reply: {str(e)}")
            self.update_progress("Error modifying reply", 0)
    
    def clear_results(self):
        """Clear all results"""
        self.email_tree.delete(*self.email_tree.get_children())
        self.original_text.delete(1.0, tk.END)
        self.summary_text.delete(1.0, tk.END)
        self.reply_text.delete(1.0, tk.END)
        self.email_data = {}
        self.current_email_id = None
        self.update_progress("Ready", 0)
        self.status_var.set("Ready")

if __name__ == "__main__":
    root = tk.Tk()
    app = EmailAIAssistant(root)
    root.mainloop()
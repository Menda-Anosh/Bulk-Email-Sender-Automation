import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import smtplib
import os
from email.message import EmailMessage
# Global mode tracker
is_dark_mode = False

# Browse file function
def browse_file(entry_widget):
    path = filedialog.askopenfilename()
    if path:
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, path)

# Browse attachment
def browse_attachment():
    path = filedialog.askopenfilename()
    if path:
        attachment_path.set(path)

# Toggle Dark Mode
def toggle_dark_mode():
    global is_dark_mode
    is_dark_mode = not is_dark_mode

    bg = "#2e2e2e" if is_dark_mode else "#eef5f9"
    fg = "white" if is_dark_mode else "black"
    for widget in root.winfo_children():
        try:
            widget.configure(bg=bg, fg=fg)
        except:
            pass
    root.configure(bg=bg)

# Preview email content
def preview_email():
    try:
        with open(entry_subject.get(), 'r', encoding='utf-8') as f:
            subject = f.read()
        with open(entry_body.get(), 'r', encoding='utf-8') as f:
            body = f.read()

        preview = f"Subject:\n{subject}\n\nBody:\n{body}"
        messagebox.showinfo("Email Preview", preview)
    except Exception as e:
        messagebox.showerror("Preview Error", str(e))

# Send emails function
def send_bulk_emails():
    sender = entry_email.get()
    password = entry_password.get()
    file = entry_file.get()
    sub_file = entry_subject.get()
    body_file = entry_body.get()
    attachment = attachment_path.get()

    try:
        with open(sub_file, 'r', encoding='utf-8') as f:
            subject = f.read()
        with open(body_file, 'r', encoding='utf-8') as f:
            body = f.read()

        # ✅ Improved Excel/CSV reading with proper engine handling
        if file.endswith('.xlsx'):
            try:
                df = pd.read_excel(file, engine='openpyxl')
            except Exception as e:
                raise ValueError("The selected file is not a valid .xlsx Excel file. Please check the file and try again.") from e
        elif file.endswith('.xls'):
            try:
                df = pd.read_excel(file, engine='xlrd')
            except Exception as e:
                raise ValueError("The selected file is not a valid .xls Excel file. Please check the file and try again.") from e
        elif file.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            raise ValueError("Only .csv, .xls or .xlsx files are supported.")

        if 'emails' not in df.columns:
            raise ValueError("Missing 'emails' column in the file.")

        recipients = df['emails'].tolist()

        confirm = messagebox.askyesno("Confirmation", f"Send email to {len(recipients)} recipients?")
        if not confirm:
            return

        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender, password)

        for email in recipients:
            msg = EmailMessage()
            msg['Subject'] = subject
            msg['From'] = sender
            msg['To'] = email
            msg.set_content(body)

            # Attach file if given
            if attachment and os.path.isfile(attachment):
                with open(attachment, 'rb') as f:
                    file_data = f.read()
                    file_name = os.path.basename(attachment)
                msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

            server.send_message(msg)
            log_area.insert(tk.END, f"✅ Sent to {email}\n")
            log_area.see(tk.END)

        server.quit()
        messagebox.showinfo("Success", "All emails sent successfully!")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# GUI Setup
root = tk.Tk()
root.title("📧 Bulk Email Sender Pro")
root.geometry("670x600")
root.configure(bg="#eef5f9")

# Entry Fields
tk.Label(root, text="Your Gmail:", bg="#eef5f9").pack()
entry_email = tk.Entry(root, width=60)
entry_email.pack()

tk.Label(root, text="App Password:", bg="#eef5f9").pack()
entry_password = tk.Entry(root, width=60, show="*")
entry_password.pack()

tk.Label(root, text="Select Email File (.csv/.xlsx/.xls):", bg="#eef5f9").pack()
frame_file = tk.Frame(root)
entry_file = tk.Entry(frame_file, width=45)
tk.Button(frame_file, text="Browse", command=lambda: browse_file(entry_file)).pack(side=tk.RIGHT)
entry_file.pack(side=tk.LEFT)
frame_file.pack(pady=4)

tk.Label(root, text="Subject Text File:", bg="#eef5f9").pack()
frame_subject = tk.Frame(root)
entry_subject = tk.Entry(frame_subject, width=45)
tk.Button(frame_subject, text="Browse", command=lambda: browse_file(entry_subject)).pack(side=tk.RIGHT)
entry_subject.pack(side=tk.LEFT)
frame_subject.pack(pady=4)

tk.Label(root, text="Body Text File:", bg="#eef5f9").pack()
frame_body = tk.Frame(root)
entry_body = tk.Entry(frame_body, width=45)
tk.Button(frame_body, text="Browse", command=lambda: browse_file(entry_body)).pack(side=tk.RIGHT)
entry_body.pack(side=tk.LEFT)
frame_body.pack(pady=4)

# Attachment
tk.Label(root, text="Attach File (optional):", bg="#eef5f9").pack()
attachment_path = tk.StringVar()
frame_attach = tk.Frame(root)
entry_attach = tk.Entry(frame_attach, textvariable=attachment_path, width=45)
tk.Button(frame_attach, text="Browse", command=browse_attachment).pack(side=tk.RIGHT)
entry_attach.pack(side=tk.LEFT)
frame_attach.pack(pady=4)

# Buttons
btn_frame = tk.Frame(root, bg="#eef5f9")
tk.Button(btn_frame, text="Preview Email", command=preview_email, bg="#004aad", fg="white").pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="Send Emails", command=send_bulk_emails, bg="green", fg="white").pack(side=tk.LEFT, padx=10)
tk.Button(btn_frame, text="🌙 Toggle Dark Mode", command=toggle_dark_mode, bg="black", fg="white").pack(side=tk.LEFT, padx=10)
btn_frame.pack(pady=10)

# Log Box
tk.Label(root, text="Log:", bg="#eef5f9").pack()
log_area = scrolledtext.ScrolledText(root, width=80, height=12)
log_area.pack()

# Launch GUI
root.mainloop()

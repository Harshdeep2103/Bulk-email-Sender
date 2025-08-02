import smtplib
import pandas as pd
from tkinter import *
from tkinter import filedialog, messagebox
from email.message import EmailMessage
import os

def send_emails():
    sender_email = sender_entry.get()
    app_password = password_entry.get()
    subject = subject_entry.get()
    body = body_text.get("1.0", END).strip()
    resume_path = resume_path_var.get()
    excel_path = excel_path_var.get()

    if not all([sender_email, app_password, subject, body, resume_path, excel_path]):
        messagebox.showerror("Error", "All fields are required!")
        return

    try:
        df = pd.read_excel(excel_path)
        hr_emails = df['Email'].dropna().tolist()
    except Exception as e:
        messagebox.showerror("Error", f"Could not read Excel: {e}")
        return

    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(sender_email, app_password)
    except Exception as e:
        messagebox.showerror("Login Failed", f"Could not log in: {e}")
        return

    for email in hr_emails:
        msg = EmailMessage()
        msg['Subject'] = subject
        msg['From'] = sender_email
        msg['To'] = email
        msg.set_content(body)

        with open(resume_path, 'rb') as f:
            file_data = f.read()
            file_name = os.path.basename(resume_path)
            msg.add_attachment(file_data, maintype='application', subtype='pdf', filename=file_name)

        try:
            server.send_message(msg)
        except Exception as e:
            messagebox.showwarning("Warning", f"Could not send to {email}: {e}")

    server.quit()
    messagebox.showinfo("Success", "Emails sent successfully!")

def browse_resume():
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    resume_path_var.set(path)

def browse_excel():
    path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
    excel_path_var.set(path)

# GUI
root = Tk()
root.title("Bulk Resume Email Sender")
root.geometry("600x500")

Label(root, text="Your Gmail:").pack()
sender_entry = Entry(root, width=50)
sender_entry.pack()

Label(root, text="Gmail App Password:").pack()
password_entry = Entry(root, width=50, show="*")
password_entry.pack()

Label(root, text="Email Subject:").pack()
subject_entry = Entry(root, width=50)
subject_entry.pack()

Label(root, text="Email Body:").pack()
body_text = Text(root, height=10, width=60)
body_text.pack()

resume_path_var = StringVar()
Button(root, text="Select Resume (PDF)", command=browse_resume).pack()
Entry(root, textvariable=resume_path_var, width=50).pack()

excel_path_var = StringVar()
Button(root, text="Select Excel File", command=browse_excel).pack()
Entry(root, textvariable=excel_path_var, width=50).pack()

Button(root, text="Send Emails", command=send_emails, bg="green", fg="white").pack(pady=10)

root.mainloop()

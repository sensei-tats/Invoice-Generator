def run_gui():
    root = tk.Tk()
    root.title("Invoice Generator")
    root.geometry("400x300")
    root.configure(bg="#f0f0f0")

    title = tk.Label(root, text="🧾 Invoice Generator", font=("Helvetica", 18, "bold"), bg="#f0f0f0")
    title.pack(pady=20)

    # Move this up here so it's defined before handlers use it
    status_label = tk.Label(root, text="", font=("Helvetica", 10), bg="#f0f0f0", fg="black")
    status_label.pack(pady=20)

    def handle_send():
        try:
            main(dry_run=False)
            status_label.config(text="✅ Invoices sent successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Something went wrong:\n{str(e)}")

    def handle_dry_run():
        try:
            main(dry_run=True)
            status_label.config(text="🔍 Dry run completed. Check console logs.")
        except Exception as e:
            messagebox.showerror("Error", f"Dry run failed:\n{str(e)}")

    send_btn = tk.Button(root, text="Send Invoices", command=handle_send, font=("Helvetica", 12), width=20, bg="#4caf50", fg="white")
    send_btn.pack(pady=10)

    dry_run_btn = tk.Button(root, text="Dry Run (Preview)", command=handle_dry_run, font=("Helvetica", 12), width=20, bg="#2196f3", fg="white")
    dry_run_btn.pack(pady=10)

    exit_btn = tk.Button(root, text="Exit", command=root.quit, font=("Helvetica", 12), width=20, bg="#f44336", fg="white")
    exit_btn.pack(pady=10)

    root.mainloop()

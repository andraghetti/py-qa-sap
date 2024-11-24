import tkinter as tk

import customizing

root = tk.Tk()
root.geometry("400x400")

# Create the vertical scrollable frame
scrollable_frame = customizing.ScrollableFrame(root)
scrollable_frame.pack(fill=tk.BOTH, expand=True)

# Add sample widgets to the inner frame
for i in range(30):
    tk.Label(scrollable_frame.inner_frame, text=f"Label {i+1}", bg="#D8E6EC", font=("Arial", 12)).pack(pady=5)

root.mainloop()


import tkinter as tk

class ScrollableFrame(tk.Frame):
    def __init__(self, master=None, **kwargs):
        tk.Frame.__init__(self, master, **kwargs)

        # Create a canvas and add it to the frame
        self.canvas = tk.Canvas(self, borderwidth=0, highlightthickness=0, background='#F0F8FF')

        # Create a frame inside the canvas to hold the widgets
        self.inner_frame = tk.Frame(self.canvas, background='#11F8FF')

        # Add a horizontal scrollbar and link it to the canvas
        self.h_scrollbar = tk.Scrollbar(self, orient="horizontal", command=self.canvas.xview, background='#F0F8FF')
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set)

        # Pack the canvas and scrollbar into the frame
        self.canvas.pack(side="top", fill=tk.BOTH, expand=tk.TRUE)
        self.h_scrollbar.pack(side="bottom", fill="x")

        # Add the inner frame to the canvas
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw", height=400, width=400)

        # Configure the canvas to update the scroll region when the frame size changes
        self.inner_frame.bind("<Configure>", self.on_inner_frame_configure)

    def on_inner_frame_configure(self, event):
        # Update the scroll region to encompass the inner frame
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def add_widget(self, widget, **kwargs):
        # Add a widget to the inner frame
        widget.place(**kwargs)
        
        # Update the canvas scroll region
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

# Create the main window
root = tk.Tk()

# Create a ScrollableFrame instance
scrollable_frame = ScrollableFrame(root, background='#10F8AA')
scrollable_frame.pack(fill=tk.BOTH, expand=tk.TRUE)

# Add a label to the inner frame
label = tk.Label(scrollable_frame.inner_frame, text="Hello, World!", bg="white")
scrollable_frame.add_widget(label, x=10, y=10)

# Start the tkinter event loop
root.mainloop()

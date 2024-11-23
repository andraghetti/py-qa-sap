import tkinter as tk

class LabelSelectorApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Label Selector")
        self.root.geometry("600x400")

        self.selected_label = None

        # Left frame for labels
        self.left_frame = tk.Frame(self.root, bg="lightblue", width=200)
        self.left_frame.pack(side="left", fill="y")

        # Right frame for display
        self.right_frame = tk.Frame(self.root, bg="white", width=400)
        self.right_frame.pack(side="right", fill="both", expand=True)

        # Label for displaying content on the right
        self.display_label = tk.Label(
            self.right_frame, 
            text="Selected content will appear here", 
            bg="white", 
            font=("Arial", 14), 
            wraplength=350, 
            justify="center"
        )
        self.display_label.pack(expand=True)

        # Add labels to the left frame
        self.label_texts = ["Label 1", "Label 2", "Label 3", "Label 4"]
        self.labels = []

        for text in self.label_texts:
            label = tk.Label(self.left_frame, text=text, font=("Arial", 12), bg="lightgray", relief="ridge")
            label.pack(pady=5, padx=10, fill="x")
            label.bind("<Button-1>", self.on_label_click)  # Bind left-click event
            self.labels.append(label)

        # Add a button to display content
        self.display_button = tk.Button(
            self.left_frame, 
            text="Display Selected Label", 
            command=self.display_selected_label, 
            bg="lightgreen", 
            font=("Arial", 12)
        )
        self.display_button.pack(pady=20, padx=10, fill="x")

    def on_label_click(self, event):
        # Deselect the previously selected label
        if self.selected_label:
            self.selected_label.config(bg="lightgray")

        # Select the new label
        self.selected_label = event.widget
        self.selected_label.config(bg="yellow")

    def display_selected_label(self):
        # Show the content of the selected label in the right frame
        if self.selected_label:
            content = self.selected_label.cget("text")
            self.display_label.config(text=f"Selected Label Content:\n{content}")
        else:
            self.display_label.config(text="No label selected!")


# Main application
if __name__ == "__main__":
    root = tk.Tk()
    app = LabelSelectorApp(root)
    root.mainloop()

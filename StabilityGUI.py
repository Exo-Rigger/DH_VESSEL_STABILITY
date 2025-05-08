import tkinter as tk
from tkinter import filedialog
import stability_program  # Import your stability program script

def browse_stability_file():
    stabdir = r'C:\Users\wgrin\Desktop\Stabiliteit'
    filepath = filedialog.askopenfilename(initialdir=stabdir, filetypes=[("Excel files", "*.xlsx")])
    stability_file_entry.delete(0, tk.END)
    stability_file_entry.insert(0, filepath)

def browse_manifest_file():
    stabdir = r'C:\Users\wgrin\Desktop\Stabiliteit\Manifesten'
    filepath = filedialog.askopenfilename(initialdir=stabdir, filetypes=[("PDF files", "*.pdf")])
    manifest_file_entry.delete(0, tk.END)
    manifest_file_entry.insert(0, filepath)

def browse_save_location():
    laaddir = r'C:\Users\wgrin\Desktop\Laadplannen'
    save_path = filedialog.asksaveasfilename(initialdir=laaddir, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    save_location_entry.delete(0, tk.END)
    save_location_entry.insert(0, save_path)

def process_files():
    stability_file = stability_file_entry.get()
    manifest_file = manifest_file_entry.get()
    save_location = save_location_entry.get()

    modified_loading_plan = stability_program.process_stability(stability_file, manifest_file)

    if save_location:
        modified_loading_plan.save(save_location)
        print(f"Modified loading plan saved to {save_location}")


# Create the main window
root = tk.Tk()
root.title("Loading Plan App")
root.geometry("600x400")  # Set the initial window size to 600x400 pixels

# Stability Calculation File
stability_file_label = tk.Label(root, text="Selecteer de stabiliteitsberekening.")
stability_file_label.pack()

stability_file_entry = tk.Entry(root)
stability_file_entry.pack()

stability_file_button = tk.Button(root, text="Browse", command=browse_stability_file)
stability_file_button.pack()

# Manifest PDF File
manifest_file_label = tk.Label(root, text="Selecteer het manifest.")
manifest_file_label.pack()

manifest_file_entry = tk.Entry(root)
manifest_file_entry.pack()

manifest_file_button = tk.Button(root, text="Browse", command=browse_manifest_file)
manifest_file_button.pack()

# Save Location
save_location_label = tk.Label(root, text="Select locatie om op te slaan.:")
save_location_label.pack()

save_location_entry = tk.Entry(root)
save_location_entry.pack()

save_location_button = tk.Button(root, text="Browse", command=browse_save_location)
save_location_button.pack()

# Process Files Button
process_button = tk.Button(root, text="Process Files", command=process_files)
process_button.pack()

# Run the GUI event loop
root.mainloop()

import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
from pdf2image import convert_from_path
import pytesseract


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Data Extractor")

        # Add a button to open a PDF file
        self.open_button = tk.Button(self.root, text="Open PDF", command=self.open_pdf)
        self.open_button.pack(pady=10)

        # Add buttons for navigation
        self.prev_button = tk.Button(self.root, text="Previous Page", command=self.prev_page)
        self.prev_button.pack(side="left", padx=10)
        self.next_button = tk.Button(self.root, text="Next Page", command=self.next_page)
        self.next_button.pack(side="right", padx=10)

        # Add a button to trigger extraction to Excel
        self.extract_button = tk.Button(self.root, text="Extract to Excel", command=self.extract_and_export)
        self.extract_button.pack(pady=10)

        self.canvas = tk.Canvas(self.root, width=600, height=800)
        self.canvas.pack(fill="both", expand=True)

        self.rectangles = []  # List to store rectangle coordinates
        self.rectangle_ids = []  # List to store canvas rectangle IDs
        self.start_x = None
        self.start_y = None
        self.pdf_path = None
        self.images = None  # To hold images of the PDF pages
        self.current_page = 0  # Current page index
        self.pil_image = None  # Original image
        self.display_image = None  # Resized image for display

    def open_pdf(self):
        # Open file dialog to select a PDF
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.load_pdf(file_path)

    def load_pdf(self, pdf_path):
        try:
            # Convert PDF pages to images with a higher DPI (e.g., 300)
            self.images = convert_from_path(pdf_path, dpi=300)
            self.current_page = 0
            self.display_pdf_page(self.current_page)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")

    def display_pdf_page(self, page_number):
        if not self.images:
            return

        # Load the specified page and display as an image
        self.pil_image = self.images[page_number]  # Store the original image
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Resize the image to fit the canvas while maintaining aspect ratio
        self.display_image = self.pil_image.copy()
        self.display_image.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)

        # Convert the resized PIL image to a Tkinter PhotoImage object
        self.img_tk = ImageTk.PhotoImage(self.display_image)

        # Clear the canvas before adding the new image
        self.canvas.delete("all")

        # Create an image on the canvas
        self.canvas.create_image(0, 0, anchor="nw", image=self.img_tk)

        # Redraw rectangles for the current page
        for rect in self.rectangles:
            self.rectangle_ids.append(self.canvas.create_rectangle(*rect, outline="red"))

        # Bind mouse events to select multiple regions
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        # Start the rectangle selection
        self.start_x = event.x
        self.start_y = event.y

    def on_mouse_drag(self, event):
        # Update the rectangle while dragging the mouse
        if self.rectangle_ids:
            self.canvas.delete(self.rectangle_ids[-1])  # Remove the last drawn rectangle
        self.rectangle_ids.append(self.canvas.create_rectangle(self.start_x, self.start_y, event.x, event.y, outline="red"))

    def on_button_release(self, event):
        # Store the selected area coordinates
        coords = (self.start_x, self.start_y, event.x, event.y)
        self.rectangles.append(coords)  # Store the coordinates
        print(f"Selected Area: {coords}")

    def prev_page(self):
        if self.images and self.current_page > 0:
            self.current_page -= 1
            self.rectangles = []  # Clear previous rectangles
            self.rectangle_ids = []  # Clear canvas rectangle IDs
            self.display_pdf_page(self.current_page)

    def next_page(self):
        if self.images and self.current_page < len(self.images) - 1:
            self.current_page += 1
            self.rectangles = []  # Clear previous rectangles
            self.rectangle_ids = []  # Clear canvas rectangle IDs
            self.display_pdf_page(self.current_page)

    def extract_and_export(self):
        if not self.rectangles:
            messagebox.showerror("Error", "Please select at least one region!")
            return

        extracted_data = []

        for page_num, image in enumerate(self.images):
            for rect in self.rectangles:
                if isinstance(rect, tuple) and len(rect) == 4:
                    # Scale the rectangle coordinates to match the original image
                    scale_x = image.width / self.display_image.width
                    scale_y = image.height / self.display_image.height
                    x0 = int(rect[0] * scale_x)
                    y0 = int(rect[1] * scale_y)
                    x1 = int(rect[2] * scale_x)
                    y1 = int(rect[3] * scale_y)

                    # Crop the image to the selected area
                    cropped_image = image.crop((x0, y0, x1, y1))

                    # Convert the cropped image to text using OCR (Tesseract)
                    text = self.extract_text_from_image(cropped_image)

                    if text:
                        extracted_data.append({"Page": page_num + 1, "Extracted Data": text.strip(), "Region": (x0, y0, x1, y1)})
                else:
                    print(f"Invalid rectangle: {rect}")

        if extracted_data:
            print("Extracted Text from all pages:")
            for entry in extracted_data:
                print(f"Page {entry['Page']} (Region {entry['Region']}):\n{entry['Extracted Data']}\n")

            # Export the extracted data to Excel
            self.export_to_excel(extracted_data)
        else:
            messagebox.showerror("Error", "No text found in the selected area on any page!")

    def extract_text_from_image(self, image):
        try:
            # Convert the image to grayscale
            image = image.convert("L")

            # Apply thresholding to improve OCR accuracy
            image = image.point(lambda x: 0 if x < 128 else 255, "1")

            # Use Tesseract OCR to extract text from the image
            text = pytesseract.image_to_string(image)
            return text
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text from image: {str(e)}")
            return ""

    def export_to_excel(self, extracted_data):
        # Convert the extracted data to a pandas DataFrame
        df = pd.DataFrame(extracted_data)

        try:
            # Ask the user where to save the Excel file
            output_file = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
            if output_file:
                # Save the data to the selected Excel file
                df.to_excel(output_file, index=False)
                messagebox.showinfo("Success", "Data exported to Excel successfully.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data to Excel: {str(e)}")


if __name__ == "__main__":
    # Create the main Tkinter window
    root = tk.Tk()
    app = PDFExtractorApp(root)
    root.mainloop()
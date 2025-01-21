import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import pandas as pd
from pdf2image import convert_from_path


class PDFExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Data Extractor")

        # Add a button to open a PDF file
        self.open_button = tk.Button(self.root, text="Open PDF", command=self.open_pdf)
        self.open_button.pack(pady=10)

        # Add a button to trigger extraction to Excel
        self.extract_button = tk.Button(self.root, text="Extract to Excel", command=self.extract_and_export)
        self.extract_button.pack(pady=10)

        self.canvas = tk.Canvas(self.root, width=600, height=800)
        self.canvas.pack(fill="both", expand=True)

        self.rectangles = []  # List to store multiple selected areas
        self.start_x = None
        self.start_y = None
        self.pdf_path = None
        self.images = None  # To hold images of the PDF pages
        self.current_page = 0  # Current page index

    def open_pdf(self):
        # Open file dialog to select a PDF
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path = file_path
            self.load_pdf(file_path)

    def load_pdf(self, pdf_path):
        # Convert PDF pages to images
        self.images = convert_from_path(pdf_path)

        # Display the first page as an image
        self.display_pdf_page(0)  # Display the first page

    def display_pdf_page(self, page_number):
        # Load the specified page and display as an image
        pil_image = self.images[page_number]

        # Get the size of the canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()

        # Resize the image to fit the canvas while maintaining aspect ratio
        pil_image.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)

        # Convert the resized PIL image to a Tkinter PhotoImage object
        self.img_tk = ImageTk.PhotoImage(pil_image)

        # Clear the canvas before adding the new image
        self.canvas.delete("all")

        # Create an image on the canvas
        self.canvas.create_image(0, 0, anchor="nw", image=self.img_tk)

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
        if self.rectangles:
            self.canvas.delete(self.rectangles[-1])  # Remove the last drawn rectangle
        self.rectangles.append(self.canvas.create_rectangle(self.start_x, self.start_y, event.x, event.y, outline="red"))

    def on_button_release(self, event):
        # Store the selected area and stop binding mouse events
        # Get the coordinates of the last rectangle created
        coords = self.canvas.coords(self.rectangles[-1])
        if len(coords) == 4:
            x0, y0, x1, y1 = coords  # Unpack the coordinates into variables
            self.rectangles[-1] = (x0, y0, x1, y1)  # Update the last rectangle with final coordinates
            print(f"Selected Area: {(x0, y0, x1, y1)}")
        else:
            print(f"Unexpected coordinates format: {coords}")

    def extract_and_export(self):
        # Ensure there are selected regions
        if not self.rectangles:
            messagebox.showerror("Error", "Please select at least one region!")
            return

        # Initialize a list to store extracted data from all pages
        extracted_data = []

        # Loop through all images (pages) and extract text from each selected region
        for page_num, image in enumerate(self.images):
            for rect in self.rectangles:
                # Check that the rect is a tuple and has four items
                if isinstance(rect, tuple) and len(rect) == 4:
                    x0, y0, x1, y1 = rect
                    # Crop the image to the selected area
                    cropped_image = image.crop((x0, y0, x1, y1))

                    # Convert the cropped image to text using OCR (Tesseract)
                    text = self.extract_text_from_image(cropped_image)

                    if text:
                        # Collect the extracted text and associate it with the page number and region
                        extracted_data.append({"Page": page_num + 1, "Extracted Data": text.strip(), "Region": rect})
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
        # Use Tesseract OCR to extract text from the image
        try:
            import pytesseract
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

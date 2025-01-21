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

        self.canvas = tk.Canvas(self.root, width=600, height=800)
        self.canvas.pack(fill="both", expand=True)

        self.rect = None
        self.start_x = None
        self.start_y = None
        self.selected_area = None
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

        # Bind mouse events to select a region
        self.canvas.bind("<ButtonPress-1>", self.on_button_press)
        self.canvas.bind("<B1-Motion>", self.on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self.on_button_release)

    def on_button_press(self, event):
        # Start the rectangle selection
        self.start_x = event.x
        self.start_y = event.y
        if self.rect:
            self.canvas.delete(self.rect)
        self.rect = self.canvas.create_rectangle(self.start_x, self.start_y, self.start_x, self.start_y, outline="red")

    def on_mouse_drag(self, event):
        # Update the rectangle while dragging the mouse
        self.canvas.coords(self.rect, self.start_x, self.start_y, event.x, event.y)

    def on_button_release(self, event):
        # Capture the selected area and stop binding mouse events
        self.selected_area = (self.start_x, self.start_y, event.x, event.y)
        print(f"Selected Area: {self.selected_area}")

        # Remove the rectangle from the canvas
        self.canvas.delete(self.rect)

        # Extract text from the selected area across all pages
        self.extract_text_from_pdf()

    def extract_text_from_pdf(self):
        # Ensure the user has selected an area
        if not self.selected_area:
            messagebox.showerror("Error", "Please select a region first!")
            return

        # Initialize a list to store extracted data from all pages
        extracted_data = []

        # Loop through all images (pages) and extract text from the selected region
        x0, y0, x1, y1 = self.selected_area
        for page_num, image in enumerate(self.images):
            # Crop the image to the selected area
            cropped_image = image.crop((x0, y0, x1, y1))

            # Convert the cropped image to text using OCR (Tesseract)
            text = self.extract_text_from_image(cropped_image)

            if text:
                # Collect the extracted text and associate it with the page number
                extracted_data.append({"Page": page_num + 1, "Extracted Data": text.strip()})

        if extracted_data:
            print("Extracted Text from all pages:")
            for entry in extracted_data:
                print(f"Page {entry['Page']}:\n{entry['Extracted Data']}\n")

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

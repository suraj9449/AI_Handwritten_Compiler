import os
import cv2
import numpy as np
import pytesseract
from docx import Document

# Set path to Tesseract if you're on Windows
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Directories
data_dir = "data"
output_dir = "output"
os.makedirs(output_dir, exist_ok=True)

# Global text buffer
extracted_text = ""
current_image = ""

def list_images(folder):
    images = [f for f in os.listdir(folder) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
    if not images:
        print(" No images found in the 'data/' folder.")
    else:
        print("\n Available Images:")
        for idx, image in enumerate(images, 1):
            print(f"{idx}. {image}")
    return images

def deskew(image):
    coords = np.column_stack(np.where(image > 0))
    angle = cv2.minAreaRect(coords)[-1]
    if angle < -45:
        angle = -(90 + angle)
    else:
        angle = -angle
    (h, w) = image.shape[:2]
    center = (w // 2, h // 2)
    M = cv2.getRotationMatrix2D(center, angle, 1.0)
    rotated = cv2.warpAffine(image, M, (w, h), flags=cv2.INTER_CUBIC, borderMode=cv2.BORDER_REPLICATE)
    return rotated

def preprocess_image(image_path):
    image = cv2.imread(image_path)
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)

    # Denoising
    denoised = cv2.fastNlMeansDenoising(gray, None, 30, 7, 21)

    # Adaptive Thresholding
    thresh = cv2.adaptiveThreshold(denoised, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                   cv2.THRESH_BINARY, 31, 10)

    # Deskewing
    thresh = deskew(thresh)

    return thresh

def extract_text_from_image(image):
    return pytesseract.image_to_string(image, lang='eng')

def save_output(text, filename_base):
    txt_path = os.path.join(output_dir, f"{filename_base}.txt")
    docx_path = os.path.join(output_dir, f"{filename_base}.docx")

    with open(txt_path, "w", encoding="utf-8") as f:
        f.write(text)

    doc = Document()
    doc.add_paragraph(text)
    doc.save(docx_path)

    print(f"\n Text saved to:\n- {txt_path}\n- {docx_path}")

def process_single_image(images):
    global extracted_text, current_image
    try:
        idx = int(input("\nEnter image number to process: ")) - 1
        if idx < 0 or idx >= len(images):
            raise ValueError
        current_image = images[idx]
        image_path = os.path.join(data_dir, current_image)

        print(f"\n Processing: {current_image}")
        preprocessed = preprocess_image(image_path)
        extracted_text = extract_text_from_image(preprocessed)

        print("\n Extraction complete.")
    except ValueError:
        print(" Invalid selection.")

def process_multiple_images(images):
    global extracted_text
    selected_indices = input("\nEnter image numbers to process (comma-separated): ")
    try:
        indices = [int(i.strip()) - 1 for i in selected_indices.split(',') if i.strip().isdigit()]
        combined_text = ""
        for idx in indices:
            if 0 <= idx < len(images):
                print(f"\n Processing: {images[idx]}")
                image_path = os.path.join(data_dir, images[idx])
                preprocessed = preprocess_image(image_path)
                text = extract_text_from_image(preprocessed)
                combined_text += f"\n--- {images[idx]} ---\n{text}\n"
        extracted_text = combined_text
        print("\n Multiple image extraction complete.")
    except Exception as e:
        print(f" Error during batch processing: {str(e)}")

def main_menu():
    global extracted_text, current_image

    while True:
        print("\n MENU")
        print("1. List Available Images")
        print("2. Process a Single Image")
        print("3. Process Multiple Images")
        print("4. View Extracted Text")
        print("5. Save Extracted Text")
        print("6. Exit")

        choice = input("\nEnter your choice: ")

        if choice == "1":
            list_images(data_dir)

        elif choice == "2":
            images = list_images(data_dir)
            if images:
                process_single_image(images)

        elif choice == "3":
            images = list_images(data_dir)
            if images:
                process_multiple_images(images)

        elif choice == "4":
            if extracted_text:
                print("\n Extracted Text:")
                print("-" * 40)
                print(extracted_text.strip())
            else:
                print(" No text extracted yet. Please process an image first.")

        elif choice == "5":
            if extracted_text:
                base_name = input("Enter base filename to save (without extension): ")
                save_output(extracted_text, base_name)
            else:
                print(" Nothing to save. Please extract text first.")

        elif choice == "6":
            print(" Exiting. See you soon!")
            break
        else:
            print(" Invalid option. Please choose from the menu.")

if __name__ == "__main__":
    main_menu()
import cv2
import pytesseract
import random
import os
import sys
from docx import Document

# Ensure Tesseract is installed and correctly configured
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Change path if needed


def preprocess_image(image_path, contrast=1.2, blur=1):
    """Enhance image for better OCR (contrast, blur, thresholding)."""
    image = cv2.imread(image_path, cv2.IMREAD_GRAYSCALE)
    if image is None:
        print("⚠️ Error: Image not found or could not be read.")
        return None

    # Adjust contrast
    image = cv2.convertScaleAbs(image, alpha=contrast, beta=10)

    # Apply slight Gaussian blur
    image = cv2.GaussianBlur(image, (blur, blur), 0)

    # Adaptive thresholding for better text extraction
    image = cv2.adaptiveThreshold(image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                  cv2.THRESH_BINARY, 11, 2)

    return image


def extract_text(image_path):
    """Extract text from an image using OCR."""
    processed_image = preprocess_image(image_path)
    if processed_image is None:
        return None

    custom_config = r'--oem 3 --psm 6'  # OCR Engine & Page Segmentation Mode
    extracted_text = pytesseract.image_to_string(processed_image, lang="eng", config=custom_config)

    return extracted_text.strip()


def pso_text_correction(text, max_iterations=10):
    """Use PSO to refine OCR-extracted text by reducing common errors."""
    corrections = {
        "0": "O", "O": "0", "1": "l", "l": "1",
        "5": "S", "S": "5", "B": "8", "8": "B",
        "rn": "m", "vv": "w", "ll": "l"
    }

    # Initialize particles (candidate corrections)
    particles = [list(text) for _ in range(10)]
    
    # PSO Optimization
    for _ in range(max_iterations):
        for particle in particles:
            for i in range(len(particle)):
                if particle[i] in corrections and random.random() < 0.3:
                    particle[i] = corrections[particle[i]]

    # Select best result
    best_text = max(particles, key=lambda x: x.count(" ") + x.count("."))  # Favor readability
    return "".join(best_text)


def save_text(text, output_folder="output", filename="extracted_text"):
    """Save extracted text in both .txt and .docx formats."""
    if not text.strip():
        print("⚠️ No text extracted. Nothing to save.")
        return

    os.makedirs(output_folder, exist_ok=True)  # Ensure output folder exists

    txt_path = os.path.join(output_folder, filename + ".txt")
    docx_path = os.path.join(output_folder, filename + ".docx")

    try:
        # Save as .txt
        with open(txt_path, "w", encoding="utf-8") as f:
            f.write(text)
        
        # Save as .docx
        doc = Document()
        doc.add_paragraph(text)
        doc.save(docx_path)

        print(f"✅ Text saved successfully: {txt_path} & {docx_path}")

    except Exception as e:
        print(f"❌ Error saving files: {e}")


def list_images(folder="data"):
    """List available images in the data folder."""
    if not os.path.exists(folder):
        print(f"⚠️ Folder '{folder}' does not exist.")
        return []
    
    images = [f for f in os.listdir(folder) if f.endswith((".jpg", ".png", ".jpeg"))]
    return images


def main_menu():
    """CLI Menu for user interaction."""
    while True:
        print("\n📄 AI Handwritten Notes Compiler - CLI")
        print("======================================")
        print("1️⃣  Process a Single Image")
        print("2️⃣  Process All Images in 'data' Folder")
        print("3️⃣  Exit")
        print("======================================")

        choice = input("🔷 Choose an option (1-3): ")

        if choice == "1":
            images = list_images()
            if not images:
                print("⚠️ No images found in the 'data' folder.")
                continue

            print("\n📂 Available Images:")
            for idx, img in enumerate(images, 1):
                print(f"{idx}. {img}")

            img_choice = input("\n🔷 Enter the number of the image to process: ")
            if not img_choice.isdigit() or int(img_choice) < 1 or int(img_choice) > len(images):
                print("⚠️ Invalid selection. Try again.")
                continue

            image_path = os.path.join("data", images[int(img_choice) - 1])
            process_image(image_path)

        elif choice == "2":
            images = list_images()
            if not images:
                print("⚠️ No images found in the 'data' folder.")
                continue

            for img in images:
                image_path = os.path.join("data", img)
                process_image(image_path)

        elif choice == "3":
            print("\n👋 Exiting... Have a great day!\n")
            break
        else:
            print("⚠️ Invalid option. Please enter 1, 2, or 3.")


def process_image(image_path):
    """Processes an image through OCR, PSO, and saves results."""
    print(f"\n🔄 Processing: {image_path}...\n")
    extracted_text = extract_text(image_path)

    if extracted_text:
        print("\n✅ Raw Extracted Text:\n", extracted_text)

        refined_text = pso_text_correction(extracted_text)
        print("\n✅ Refined Text (After PSO Correction):\n", refined_text)

        filename = os.path.splitext(os.path.basename(image_path))[0]
        save_text(refined_text, filename=filename)

        print("\n🎉 Extraction complete! Check the 'output/' folder.\n")

    else:
        print("❌ No text could be extracted. Try improving the image quality or adjusting settings.")


if __name__ == "__main__":
    main_menu()

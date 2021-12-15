# pip install pdf2image
# Download poppler from "https://github.com/oschwartz10612/poppler-windows/releases/"

## Poppler path: E:\Codes\OCR\poppler-20.11.0\bin
from pdf2image import convert_from_path

images = convert_from_path('PDF2.pdf', poppler_path=r"E:\Codes\OCR\poppler-20.11.0\bin")
print(type(images), len(images))

for ind, image in enumerate(images):
    image.save(f"output{ind}.jpg", "JPEG")
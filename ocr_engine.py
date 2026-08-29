import asyncio
import os
import winrt.windows.media.ocr as ocr
import winrt.windows.graphics.imaging as imaging
import winrt.windows.storage as storage
import winrt.windows.foundation as foundation
import winrt.windows.globalization as globalization

class InternalWinOCR:
    def __init__(self):
        print("Initializing Windows OCR Engine...")
        self.engine = ocr.OcrEngine.try_create_from_user_profile_languages()
        if not self.engine:
            print("Warning: Windows OCR Engine could not be created.")
        else:
            print(f"Windows OCR Engine initialized for language: {self.engine.recognizer_language.language_tag}")
            
    async def extract_text_async(self, image_path: str) -> str:
        if not self.engine:
            return ""
        try:
            file = await storage.StorageFile.get_file_from_path_async(os.path.abspath(image_path))
            stream = await file.open_async(storage.FileAccessMode.READ)
            decoder = await imaging.BitmapDecoder.create_async(stream)
            bitmap = await decoder.get_software_bitmap_async()
            result = await self.engine.recognize_async(bitmap)
            return result.text
        except Exception as e:
            print(f"Windows OCR Error: {e}")
            return ""

class PowerfulOCR:
    def __init__(self):
        self.win_ocr = InternalWinOCR()
        self.easy_reader = None
        self.initialized_easy = False
        
        self.bn_easy_reader = None
        self.initialized_bn_easy = False
        self.use_bengali = False
        
    def init_easyocr(self):
        if not self.initialized_easy:
            try:
                print("Initializing EasyOCR (Hindi + English)... This may take a moment on first run.")
                import easyocr
                # Initialize for Hindi and English
                self.easy_reader = easyocr.Reader(['hi', 'en'], gpu=False)
                self.initialized_easy = True
                print("EasyOCR Initialized successfully.")
            except Exception as e:
                print(f"Failed to initialize EasyOCR: {e}")

    def init_bengali_easyocr(self):
        if not self.initialized_bn_easy:
            try:
                print("Initializing EasyOCR (Bengali + English)... This may take a moment on first run.")
                import easyocr
                self.bn_easy_reader = easyocr.Reader(['bn', 'en'], gpu=False)
                self.initialized_bn_easy = True
                print("EasyOCR Bengali Initialized successfully.")
            except Exception as e:
                print(f"Failed to initialize EasyOCR Bengali: {e}")

    def extract_text(self, image_path: str) -> str:
        text = ""
        
        # 1. If Bengali OCR is requested, use EasyOCR
        if getattr(self, 'use_bengali', False):
            try:
                self.init_bengali_easyocr()
                if self.initialized_bn_easy and self.bn_easy_reader:
                    print(f"Running EasyOCR (Bengali + English) on {image_path}...")
                    results = self.bn_easy_reader.readtext(image_path)
                    text = " ".join([res[1] for res in results])
                    if text:
                        print(f"EasyOCR Bengali extracted {len(text)} characters.")
                        return text
            except Exception as e:
                print(f"EasyOCR Bengali failed: {e}")
        
        # 2. Use Windows OCR primarily as requested (English only / Default)
        try:
            print(f"Running Windows OCR on {image_path}...")
            loop = asyncio.new_event_loop()
            asyncio.set_event_loop(loop)
            text = loop.run_until_complete(self.win_ocr.extract_text_async(image_path))
            loop.close()
            if text:
                print(f"Windows OCR extracted {len(text)} characters.")
        except Exception as e:
            print(f"Windows OCR failed: {e}")

        # Fallback to Tesseract if Windows OCR fails
        if not text:
            try:
                import pytesseract
                from PIL import Image
                # Fallback to Bengali if request was Bengali
                if getattr(self, 'use_bengali', False):
                    text = pytesseract.image_to_string(Image.open(image_path), lang='ben')
                else:
                    text = pytesseract.image_to_string(Image.open(image_path))
                print("Tesseract fallback used.")
            except:
                pass
                
        return text

# Export as WindowsOCR for backward compatibility in main.py
WindowsOCR = PowerfulOCR

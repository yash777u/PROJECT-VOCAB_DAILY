import os
import subprocess
import base64
import streamlit as st

@st.cache_resource(show_spinner=False)
def get_slide_images():
    """
    Ensure the PPTX file is converted to slides and return a list of base64 PNG data.
    Cached for the lifetime of the server process — slides are re-loaded only on restart.
    """
    pptx_path = os.path.join("data", "German.pptx")
    cache_dir = os.path.join("data", "slides_cache")
    
    if not os.path.exists(pptx_path):
        return []
        
    os.makedirs(cache_dir, exist_ok=True)
    
    def get_num(filename):
        try:
            # extract number from 'slide-X.png'
            return int(filename.split("-")[1].split(".")[0])
        except:
            return 99999
            
    # Check if we need to convert
    pptx_mtime = os.path.getmtime(pptx_path)
    slide_files = sorted(
        [f for f in os.listdir(cache_dir) if f.startswith("slide-") and f.endswith(".png")],
        key=get_num
    )
    
    cache_valid = len(slide_files) > 0
    if cache_valid:
        for sf in slide_files:
            sf_path = os.path.join(cache_dir, sf)
            if os.path.getmtime(sf_path) < pptx_mtime or os.path.getsize(sf_path) == 0:
                cache_valid = False
                break
                
    if not cache_valid:
        # Clean up cache dir first
        for f in os.listdir(cache_dir):
            try:
                os.remove(os.path.join(cache_dir, f))
            except:
                pass
                
        # Try converting via LibreOffice
        try:
            subprocess.run([
                "/snap/bin/libreoffice", "--headless", "--convert-to", "pdf",
                "--outdir", cache_dir, pptx_path
            ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            try:
                subprocess.run([
                    "libreoffice", "--headless", "--convert-to", "pdf",
                    "--outdir", cache_dir, pptx_path
                ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception as e:
                print(f"LibreOffice conversion failed: {e}")
                
        pdf_path = os.path.join(cache_dir, "German.pdf")
        if os.path.exists(pdf_path):
            try:
                subprocess.run([
                    "/usr/bin/pdftoppm", "-png", "-r", "150",
                    pdf_path, os.path.join(cache_dir, "slide")
                ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            except Exception:
                try:
                    subprocess.run([
                        "pdftoppm", "-png", "-r", "150",
                        pdf_path, os.path.join(cache_dir, "slide")
                    ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                except Exception as e:
                    print(f"pdftoppm conversion failed: {e}")
            
            # Clean up the PDF file to save space
            try:
                os.remove(pdf_path)
            except:
                pass
                
        slide_files = sorted(
            [f for f in os.listdir(cache_dir) if f.startswith("slide-") and f.endswith(".png")],
            key=get_num
        )
    
    # Read files and convert to base64
    slides_b64 = []
    for f in slide_files:
        path = os.path.join(cache_dir, f)
        try:
            with open(path, "rb") as image_file:
                encoded_string = base64.b64encode(image_file.read()).decode("utf-8")
                slides_b64.append(f"data:image/png;base64,{encoded_string}")
        except Exception as e:
            print(f"Error encoding {path}: {e}")
            
    return slides_b64

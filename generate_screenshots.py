import os
import subprocess
from pdf2image import convert_from_path

def convert_excel_to_images(excel_path, output_dir="images"):
    """
    Converts individual sheets of an Excel file into clean PNG screenshots
    using an open-source headless rendering engine.
    """
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
        
    print(f"📦 Rendering visual frames for: {excel_path}...")
    
    # Step 1: Render the Excel sheets into a temporary PDF layout cleanly
    # Note: Requires LibreOffice installed on your machine/GitHub runner
    try:
        cmd = [
            "libreoffice", "--headless", "--convert-to", "pdf",
            "--outdir", output_dir, excel_path
        ]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL)
        pdf_path = os.path.join(output_dir, excel_path.replace(".xlsx", ".pdf"))
        
        # Step 2: Convert the high-fidelity PDF sheets into crisp PNG screenshots
        pages = convert_from_path(pdf_path, dpi=200)
        
        sheet_names = ["Macro_Financial_Dashboard", "Micro_Operational_Checklist"]
        for idx, page in enumerate(pages):
            # Fallback naming if pages exceed planned sheets
            name = sheet_names[idx] if idx < len(sheet_names) else f"Sheet_Page_{idx+1}"
            image_filename = os.path.join(output_dir, f"{name}.png")
            
            # Save as professional, clear PNG
            page.save(image_filename, "PNG")
            print(f"📸 Screenshot successfully generated: {image_filename}")
            
        # Clean up intermediate temporary files
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
            
    except FileNotFoundError:
        print("\n❌ Error: 'libreoffice' command line utility not detected.")
        print("💡 Alternative Quick Method: Open your generated Excel file, take a screen snippet")
        print("   of the colored cards, save them as 'images/Dashboard_Preview.png' manually.")

if __name__ == "__main__":
    # Point this to your generated workbook suite
    target_workbook = "MSME_Inventory_Consulting_Suite.xlsx"
    if os.path.exists(target_workbook):
        convert_excel_to_images(target_workbook)
    else:
        print(f"❌ Could not find {target_workbook}. Please run your main generation script first.")

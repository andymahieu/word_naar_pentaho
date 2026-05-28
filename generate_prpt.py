import zipfile
import os

def main():
    prpt_filename = "Stortformulier_fixed.prpt"
    extract_dir = "check_prpt"  # extracted from original PRPT
    
    # mimetype content
    mimetype_content = b"application/vnd.pentaho.reporting-archive"
    
    with zipfile.ZipFile(prpt_filename, 'w', zipfile.ZIP_DEFLATED) as zf:
        # First, add mimetype with ZIP_STORED (no compression)
        zf.writestr("mimetype", mimetype_content, compress_type=zipfile.ZIP_STORED)
        
        # Now add the rest of the files
        for root, dirs, files in os.walk(extract_dir):
            for file in files:
                if file == "mimetype":
                    continue  # skip, we already added
                full_path = os.path.join(root, file)
                # Determine the archive name (relative to extract_dir)
                arcname = os.path.relpath(full_path, extract_dir)
                # Read the file content
                with open(full_path, 'rb') as f:
                    data = f.read()
                # Add with deflate compression (default)
                zf.writestr(arcname, data, compress_type=zipfile.ZIP_DEFLATED)
    
    print(f"Generated {prpt_filename}")

if __name__ == "__main__":
    main()

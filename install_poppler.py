#!/usr/bin/env python3
"""
Poppler Installer for Windows
Downloads and installs Poppler utilities required for pdf2image
"""

import urllib.request
import zipfile
import os
import subprocess
from pathlib import Path

def download_and_install_poppler():
    """Download and install Poppler automatically"""
    print("📥 Downloading Poppler for Windows...")
    
    poppler_url = "https://github.com/oschwartz10612/poppler-windows/releases/download/v24.02.0-0/Release-24.02.0-0.zip"
    poppler_dir = Path("C:/poppler")
    temp_zip = Path("poppler.zip")
    
    try:
        # Download
        urllib.request.urlretrieve(poppler_url, temp_zip)
        print("✅ Downloaded Poppler")
        
        # Create directory if it doesn't exist
        poppler_dir.mkdir(exist_ok=True)
        
        # Extract
        with zipfile.ZipFile(temp_zip, 'r') as zip_ref:
            zip_ref.extractall(poppler_dir)
        print(f"✅ Extracted to {poppler_dir}")
        
        # Clean up
        temp_zip.unlink()
        
        # Find the bin directory
        bin_paths = list(poppler_dir.glob("**/bin"))
        if bin_paths:
            bin_path = bin_paths[0]
            print(f"📁 Poppler bin directory: {bin_path}")
            
            # Test if poppler works
            try:
                result = subprocess.run([str(bin_path / "pdftoppm.exe"), "-h"], 
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    print("✅ Poppler is working correctly")
                    print(f"🎯 Add this to your PATH: {bin_path}")
                    return str(bin_path)
                else:
                    print("❌ Poppler test failed")
                    return None
            except Exception as e:
                print(f"❌ Poppler test error: {e}")
                return None
        else:
            print("❌ Poppler bin directory not found after extraction")
            return None
            
    except Exception as e:
        print(f"❌ Failed to download/install Poppler: {e}")
        return None

if __name__ == "__main__":
    bin_path = download_and_install_poppler()
    if bin_path:
        print(f"\n🔧 To fix the PATH issue, run this command as Administrator:")
        print(f'$env:PATH += ";{bin_path}"')
        print("\nOr add it permanently to your system PATH via Windows settings.")
#!/usr/bin/env python3
"""
Installation script for File Converter App dependencies
This script helps install system dependencies and Python packages
"""

import subprocess
import sys
import os
import platform

def run_command(command, description):
    """Run a command and handle errors"""
    print(f"\n🔄 {description}...")
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        print(f"✅ {description} completed successfully")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ {description} failed: {e}")
        print(f"Error output: {e.stderr}")
        return False

def install_python_packages():
    """Install Python packages from requirements.txt"""
    print("\n📦 Installing Python packages...")
    
    # Upgrade pip first
    run_command(f"{sys.executable} -m pip install --upgrade pip", "Upgrading pip")
    
    # Install packages from requirements.txt
    if os.path.exists("requirements.txt"):
        run_command(f"{sys.executable} -m pip install -r requirements.txt", "Installing Python packages")
    else:
        print("❌ requirements.txt not found!")
        return False
    
    return True

def install_system_dependencies():
    """Install system dependencies based on OS"""
    system = platform.system().lower()
    
    if system == "windows":
        print("\n🪟 Windows detected")
        print("📋 Manual installation required for:")
        print("   1. LibreOffice (for PowerPoint to PDF conversion)")
        print("   2. Poppler (for PDF to image conversion)")
        print("\n📥 Download links:")
        print("   LibreOffice: https://www.libreoffice.org/download/download/")
        print("   Poppler: https://github.com/oschwartz10612/poppler-windows/releases/")
        print("\n⚠️  After installing Poppler, add it to your PATH environment variable")
        
    elif system == "linux":
        print("\n🐧 Linux detected")
        # Try different package managers
        if run_command("which apt-get", "Checking for apt-get"):
            run_command("sudo apt-get update", "Updating package list")
            run_command("sudo apt-get install -y poppler-utils libreoffice", "Installing system dependencies")
        elif run_command("which yum", "Checking for yum"):
            run_command("sudo yum install -y poppler-utils libreoffice", "Installing system dependencies")
        elif run_command("which dnf", "Checking for dnf"):
            run_command("sudo dnf install -y poppler-utils libreoffice", "Installing system dependencies")
        else:
            print("❌ No supported package manager found")
            
    elif system == "darwin":  # macOS
        print("\n🍎 macOS detected")
        if run_command("which brew", "Checking for Homebrew"):
            run_command("brew install poppler", "Installing Poppler")
            run_command("brew install --cask libreoffice", "Installing LibreOffice")
        else:
            print("❌ Homebrew not found. Please install Homebrew first:")
            print("   /bin/bash -c \"$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)\"")

def check_installation():
    """Check if key dependencies are installed"""
    print("\n🔍 Checking installation...")
    
    # Check Python packages
    try:
        import streamlit
        import fitz  # PyMuPDF
        import PIL
        print("✅ Python packages installed correctly")
    except ImportError as e:
        print(f"❌ Python package missing: {e}")
        return False
    
    # Check system dependencies
    system = platform.system().lower()
    if system != "windows":
        # Check poppler
        if run_command("pdftoppm -h", "Checking Poppler installation"):
            print("✅ Poppler installed correctly")
        else:
            print("❌ Poppler not found or not in PATH")
    
    return True

def main():
    """Main installation function"""
    print("🚀 File Converter App - Dependency Installation")
    print("=" * 50)
    
    # Install Python packages
    if not install_python_packages():
        print("\n❌ Python package installation failed!")
        return False
    
    # Install system dependencies
    install_system_dependencies()
    
    # Check installation
    check_installation()
    
    print("\n" + "=" * 50)
    print("🎉 Installation completed!")
    print("\n📋 Next steps:")
    print("   1. Ensure all system dependencies are installed")
    print("   2. Run: streamlit run app.py")
    print("   3. Open your browser to the provided URL")
    
    return True

if __name__ == "__main__":
    main()
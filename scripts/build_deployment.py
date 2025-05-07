"""
Script to build a deployment package for the SF132 to SF133 Excel Processor.
"""
import os
import shutil
import sys
import subprocess
import zipfile
from pathlib import Path
from datetime import datetime

# Add parent directory to path
parent_dir = str(Path(__file__).resolve().parent.parent)
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

def create_build_directory():
    """Create build directory structure."""
    build_dir = Path(parent_dir) / "build"
    # Clear previous build if exists
    if build_dir.exists():
        shutil.rmtree(build_dir)
    
    # Create directories
    build_dir.mkdir(exist_ok=True)
    (build_dir / "dist").mkdir(exist_ok=True)
    
    return build_dir

def install_requirements():
    """Install required packages."""
    requirements_file = Path(parent_dir) / "requirements.txt"
    if not requirements_file.exists():
        print("Error: requirements.txt not found")
        sys.exit(1)
        
    print("Installing requirements...")
    subprocess.run([sys.executable, "-m", "pip", "install", "-r", str(requirements_file)])

def build_executable():
    """Build executable using PyInstaller."""
    try:
        # Install PyInstaller if not present
        subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"])
        
        # Build with PyInstaller
        print("Building executable...")
        subprocess.run([
            sys.executable, 
            "-m", 
            "PyInstaller",
            "--name=SF132_to_SF133_Processor",
            "--noconsole",
            "--onefile",
            "--icon=currency_icon.ico",
            "--add-data=forest-dark.tcl;.",
            os.path.join(parent_dir, "main.py")
        ], check=True)
        
        print("Build completed successfully")
    except subprocess.SubprocessError as e:
        print(f"Build failed: {e}")
        sys.exit(1)

def package_distribution():
    """Create ZIP package of distribution."""
    dist_dir = Path(parent_dir) / "dist"
    if not dist_dir.exists():
        print("Error: dist directory not found")
        return
    
    # Create timestamp
    timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
    zip_filename = f"SF132_to_SF133_Processor_{timestamp}.zip"
    
    with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
        # Add all files from dist
        for file_path in dist_dir.glob('**/*'):
            if file_path.is_file():
                zipf.write(
                    file_path, 
                    arcname=file_path.relative_to(dist_dir)
                )
        # Add README
        readme = Path(parent_dir) / "README.md"
        if readme.exists():
            zipf.write(readme, arcname="README.md")
    
    print(f"Distribution packaged to: {zip_filename}")

if __name__ == "__main__":
    print("Starting build process...")
    build_dir = create_build_directory()
    install_requirements()
    build_executable()
    package_distribution()
    print("Build process completed")

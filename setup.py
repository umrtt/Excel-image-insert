#!/usr/bin/env python
"""
Installation and Setup Script - Excel Image Insert
Run this script to verify and set up the application
"""
import sys
import subprocess
from pathlib import Path


def check_python_version():
    """Verify Python version is 3.8 or higher."""
    if sys.version_info < (3, 8):
        print("❌ Python 3.8+ required")
        return False
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor} detected")
    return True


def check_venv():
    """Check if virtual environment is activated."""
    in_venv = sys.prefix != sys.base_prefix
    if not in_venv:
        print("⚠️  Virtual environment not detected")
        print("   Run: python -m venv .venv")
        print("   Then activate it")
        return False
    print("✅ Virtual environment is active")
    return True


def install_requirements():
    """Install required packages."""
    try:
        print("\n📦 Installing dependencies...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "-r", "requirements.txt"])
        print("✅ Dependencies installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install dependencies")
        return False


def verify_imports():
    """Verify all required modules can be imported."""
    required_modules = ['eel', 'xlwings', 'PIL', 'tkinter']
    
    print("\n🔍 Verifying imports...")
    all_ok = True
    
    for module in required_modules:
        try:
            __import__(module)
            print(f"✅ {module} imported successfully")
        except ImportError:
            print(f"❌ Failed to import {module}")
            all_ok = False
    
    return all_ok


def check_directory_structure():
    """Verify directory structure."""
    required_dirs = [
        'app/controllers',
        'app/services',
        'app/models',
        'app/interfaces',
        'app/utils',
        'infrastructure',
        'web',
        'logs'
    ]
    
    print("\n📁 Checking directory structure...")
    all_ok = True
    
    for dir_path in required_dirs:
        if Path(dir_path).exists():
            print(f"✅ {dir_path}/")
        else:
            print(f"❌ {dir_path}/ missing")
            all_ok = False
    
    return all_ok


def main():
    """Run all checks."""
    print("=" * 60)
    print("Excel Image Insert - Setup Verification")
    print("=" * 60)
    
    checks = [
        ("Python Version", check_python_version),
        ("Virtual Environment", check_venv),
        ("Directory Structure", check_directory_structure),
    ]
    
    results = []
    for check_name, check_func in checks:
        try:
            result = check_func()
            results.append((check_name, result))
        except Exception as e:
            print(f"❌ Error in {check_name}: {e}")
            results.append((check_name, False))
    
    # Install requirements if checks pass
    if all(result for _, result in results):
        if install_requirements():
            results.append(("Dependencies", verify_imports()))
    
    # Summary
    print("\n" + "=" * 60)
    print("Setup Summary")
    print("=" * 60)
    
    for check_name, result in results:
        status = "✅ PASS" if result else "❌ FAIL"
        print(f"{check_name:.<40} {status}")
    
    print("=" * 60)
    
    if all(result for _, result in results):
        print("\n🎉 Setup successful! You can now run: python main.py")
        return 0
    else:
        print("\n⚠️  Some checks failed. Please fix the issues above.")
        return 1


if __name__ == '__main__':
    sys.exit(main())

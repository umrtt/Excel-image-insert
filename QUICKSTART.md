# Quick Start Guide

## 🚀 Getting Started in 5 Minutes

### Step 1: Navigate to Project Directory
```bash
cd "11. Excel image insert"
```

### Step 2: Create and Activate Virtual Environment

**Windows (PowerShell):**
```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

**Windows (CMD):**
```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

**macOS/Linux:**
```bash
python -m venv .venv
source .venv/bin/activate
```

### Step 3: Install Dependencies
```bash
pip install -r requirements.txt
```

### Step 4: Run the Application
```bash
python main.py
```

The application will automatically open in your default browser.

---

## 📋 First Run Checklist

- [ ] Python 3.8+ installed
- [ ] Virtual environment activated
- [ ] Dependencies installed
- [ ] Microsoft Excel installed
- [ ] Application running without errors

---

## 🎯 Basic Workflow

1. **Click "Browse"** → Select your Excel file (.xlsx)
2. **Select Sheet** → Choose target sheet from dropdown
3. **Click "Add Images"** → Select one or more images
4. **Enter Cell Position** → Type target cell (e.g., A1)
5. **Click "Insert Images"** → Images will be inserted and saved

---

## 💡 Tips & Tricks

### Tip 1: Auto-sizing
Leave "Auto-fit image to cell size" checked to automatically scale images to cell dimensions.

### Tip 2: Cell References
You can use any valid Excel cell reference:
- Single cells: `A1`, `Z99`
- Mixed: `AA1`, `XFD1048576`

### Tip 3: Multiple Sheets
Each image is associated with a specific sheet. Different images can go to different sheets.

### Tip 4: Merged Cells
The app automatically detects merged cells and fits images to the entire merged range.

---

## ⚙️ Configuration

### Change Port (if 8000 is in use)
Edit `main.py`:
```python
eel.start('index.html', port=8001)  # Change 8000 to 8001
```

### Enable Debug Logging
At the top of `main.py`:
```python
logger.set_level(LogLevel.DEBUG)
```

---

## 🐛 Common Issues

### Issue: "No module named 'eel'"
**Solution:** Reinstall dependencies
```bash
pip install -r requirements.txt --force-reinstall
```

### Issue: "xlwings requires Microsoft Excel"
**Solution:** Install Microsoft Excel or use LibreOffice Calc with xlwings support

### Issue: "Port 8000 already in use"
**Solution:** Change port in main.py to 8001, 8002, etc.

### Issue: Images not appearing in Excel
**Solution:**
1. Verify image format is supported (.png, .jpg, .gif)
2. Check cell position is correct
3. Try a different cell to rule out formatting issues

---

## 📊 Project Statistics

- **Total Python Files:** 12
- **Lines of Code:** 2,500+
- **Architecture Pattern:** Clean Architecture + SOLID
- **Test Coverage:** Ready for unit tests
- **Documentation:** Comprehensive

---

## 🔗 Important Files

| File | Purpose |
|------|---------|
| `main.py` | Application entry point |
| `requirements.txt` | Python dependencies |
| `web/index.html` | User interface |
| `app/controllers/app_controller.py` | API endpoints |
| `app/services/excel_service.py` | Excel business logic |
| `infrastructure/excel_handler.py` | xlwings wrapper |
| `logs/debug_DDMM.txt` | Application logs |

---

## 📚 Documentation

- **README.md** - Full project documentation
- **DEVELOPMENT.md** - Development guide and examples
- **This file** - Quick start guide

---

## 🎓 Learning Resources

### Architecture Pattern
- **Clean Architecture:** Separates concerns into distinct layers
- **SOLID Principles:** Ensures maintainable and testable code
- **Dependency Injection:** Enables easy testing and component swapping

### Technologies Used
- **Eel:** Creates desktop apps with Python + Web technologies
- **xlwings:** Provides Python API for Excel
- **Pillow:** Image processing library

---

## ✅ Next Steps

1. **Run the app** - `python main.py`
2. **Try inserting images** - Follow the workflow above
3. **Check logs** - Open `logs/debug_DDMM.txt` to see detailed logs
4. **Explore code** - Review the architecture in `app/` folder
5. **Extend functionality** - See DEVELOPMENT.md for examples

---

## 🆘 Need Help?

1. Check the **README.md** for comprehensive documentation
2. Review **DEVELOPMENT.md** for code examples
3. Check logs in `logs/` folder for error details
4. Verify all dependencies are installed: `pip list | grep -E "eel|xlwings|Pillow"`

---

**Happy image inserting! 🎉**

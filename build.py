# build.py
import PyInstaller.__main__
import sys

PyInstaller.__main__.run([
    'app.py',
    '--name=DonationReceiptGenerator',
    '--onefile',
    '--windowed',
    '--icon=donation.ico',  # Optional: Add your own .ico file
    '--add-data=.donation_receipt_config.json;.',  # Include config file
    '--hidden-import=docx2pdf',
    '--hidden-import=num2words.lang_DE',
    '--hidden-import=tqdm',
    '--hidden-import=thefuzz',
    '--hidden-import=msoffcrypto',
    '--collect-submodules=docx2pdf',
    '--collect-submodules=num2words',
    '--collect-submodules=thefuzz',
    '--collect-submodules=msoffcrypto',
])
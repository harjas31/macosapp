from setuptools import setup
import os

def get_data_files():
    data_files = []
    # Include all .py files except main2.py
    for file in os.listdir('.'):
        if file.endswith('.py') and file != 'main2.py':
            data_files.append(file)
    return [('', data_files)]  # This puts the files in the root of the .app bundle

APP = ['main2.py']
DATA_FILES = get_data_files()
OPTIONS = {
    'argv_emulation': True,
    'packages': ['certifi', 'customtkinter'],
    'includes': ['tkinter'],
    'excludes': ['client_secret.json'],
    'plist': {
        'CFBundleName': "ProductInfoFetcher",
        'CFBundleDisplayName': "Product Info Fetcher",
        'CFBundleGetInfoString': "Fetches product information from Amazon and Flipkart",
        'CFBundleIdentifier': "com.yourdomain.ProductInfoFetcher",
        'CFBundleVersion': "1.0.0",
        'CFBundleShortVersionString': "1.0",
    },
}

setup(
    app=APP,
    data_files=DATA_FILES,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

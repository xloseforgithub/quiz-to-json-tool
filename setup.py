from setuptools import setup

APP = ['quiz_to_json.py']
OPTIONS = {
    'argv_emulation': True,
    'iconfile': None,  # 你可以替换成你的.icns图标路径
    'plist': {
        'CFBundleName': '题库转JSON工具',
        'CFBundleDisplayName': '题库转JSON工具',
        'CFBundleIdentifier': 'com.yourdomain.quiz_to_json',
        'CFBundleVersion': '1.0',
    },
    'packages': ['openpyxl', 'docx', 'fitz', 'tkinter'],
}

setup(
    app=APP,
    options={'py2app': OPTIONS},
    setup_requires=['py2app'],
)

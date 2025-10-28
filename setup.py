from setuptools import setup, find_packages

setup(
    name="LicensePlateAnalyzer",
    version="1.0.0",
    description="表格拆分工具",
    packages=find_packages(),
    install_requires=[
        "pandas",
        "matplotlib",
        "openpyxl",
        "tkinter"
    ],
    entry_points={
        'console_scripts': [
            'plate-analyzer=pivot_table:main'
        ]
    }
)

# pyinstaller --onefile --windowed --name="表格拆分工具" table_splitting.py
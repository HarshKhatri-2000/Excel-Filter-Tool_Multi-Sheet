from setuptools import setup, find_packages

with open("README.md", "r", encoding="utf-8") as f:
    long_description = f.read()

setup(
    name="Excel Sheet_Filter Multple Sheet",
    version="11.0.0",
    author="HARSH KHATRI",
    author_email="harshkhatri.pro@gmail.com",
    description="Filter & Split Excel files by column values with a simple GUI",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/HarshKhatri-2000/Excel-Filter-Tool_Multi-Sheet",  # ← Put your username
    py_modules=["Excel Sheet_Filter Multple Sheet"],
    python_requires=">=3.7",
    install_requires=[
        "openpyxl>=3.0.0",
    ],
    entry_points={
        "console_scripts": [
            "excel-filter-tool=excel_filter_tool:main",
        ],
    },
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "Intended Audience :: End Users/Desktop",
        "Development Status :: 5 - Production/Stable",
    ],
    keywords="excel filter split openpyxl spreadsheet tool gui",
)

#!/usr/bin/env python3
"""
Setup configuration for Logistics Document Generator package
"""

from setuptools import setup, find_packages
import os
from pathlib import Path

# Read the README file
def read_readme():
    """Read README file for long description"""
    readme_path = Path(__file__).parent / "README.md"
    if readme_path.exists():
        with open(readme_path, "r", encoding="utf-8") as f:
            return f.read()
    return "Logistics Document Generator - Generate shipping placards from Excel data"

# Read requirements
def read_requirements():
    """Read requirements from requirements.txt"""
    req_path = Path(__file__).parent / "requirements.txt"
    if req_path.exists():
        with open(req_path, "r", encoding="utf-8") as f:
            # Filter out comments and empty lines
            requirements = []
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    requirements.append(line)
            return requirements
    return []

setup(
    name="logistics-document-generator",
    version="2.1.0",
    author="Logistics Team",
    author_email="logistics@company.com",
    description="A secure, high-performance logistics document generator for shipping placards",
    long_description=read_readme(),
    long_description_content_type="text/markdown",
    url="https://github.com/your-company/logistics-document-generator",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 5 - Production/Stable",
        "Intended Audience :: End Users/Desktop",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Office/Business",
        "Topic :: Text Processing :: Markup",
    ],
    python_requires=">=3.9",
    install_requires=read_requirements(),
    extras_require={
        "dev": [
            "pytest>=7.0.0",
            "pytest-cov>=4.0.0",
            "black>=23.0.0",
            "flake8>=6.0.0",
            "mypy>=1.0.0",
        ],
        "gui": [
            "dearpygui>=1.11.1",
        ],
    },
    entry_points={
        "console_scripts": [
            "placard-generator=logistics_generator.cli:main",
            "placard-gui=logistics_generator.gui:main",
        ],
    },
    package_data={
        "logistics_generator": [
            "templates/*.docx",
            "data/*.xlsx",
            "data/*.xls",
        ],
    },
    include_package_data=True,
    zip_safe=False,
    keywords="logistics, shipping, placards, documents, excel, word",
    project_urls={
        "Bug Reports": "https://github.com/your-company/logistics-document-generator/issues",
        "Source": "https://github.com/your-company/logistics-document-generator",
        "Documentation": "https://github.com/your-company/logistics-document-generator/wiki",
    },
) 
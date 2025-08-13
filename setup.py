"""
Setup configuration for outlook-mcp package
"""
from setuptools import setup, find_packages
from pathlib import Path

# Read the README for long description
readme_path = Path(__file__).parent / "README.md"
long_description = readme_path.read_text(encoding="utf-8") if readme_path.exists() else ""

setup(
    name="outlook-mcp",
    version="2.0.0",
    author="Connor Larson",
    author_email="",
    description="MCP server for Microsoft Outlook integration - search, read, and filter emails/calendar items",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/cjmlarson/outlook-mcp",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: Microsoft :: Windows",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
        "Topic :: Communications :: Email",
        "Topic :: Office/Business",
    ],
    python_requires=">=3.8",
    install_requires=[
        "mcp>=0.1.0",
        "pywin32>=300",
    ],
    entry_points={
        "console_scripts": [
            "outlook-mcp=outlook_mcp.server:serve",
        ],
    },
    keywords="mcp outlook email calendar automation claude anthropic",
    project_urls={
        "Bug Reports": "https://github.com/cjmlarson/outlook-mcp/issues",
        "Source": "https://github.com/cjmlarson/outlook-mcp",
    },
)
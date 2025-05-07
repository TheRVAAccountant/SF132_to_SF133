"""
Setup script for the SF132 to SF133 reconciliation tool.

This file is provided for backward compatibility with older setuptools versions.
For modern Python packaging, pyproject.toml is used as the primary configuration.
"""

from setuptools import setup, find_packages

setup(
    name="sf132-sf133-recon",
    version="1.0.0",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    install_requires=[
        "openpyxl>=3.0.0",
        "pandas>=1.0.0",
        "psutil",
        "pywin32;platform_system=='Windows'",  # Essential for Windows operation
    ],
    extras_require={
        "dev": ["pytest", "black", "flake8"],
    },
    entry_points={
        "console_scripts": [
            "sf132-sf133-recon=sf132_sf133_recon.main:main",
        ],
    },
    python_requires=">=3.9",
)
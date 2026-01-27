#!/usr/bin/env python3

from setuptools import setup, find_packages

setup(
    name="md2pptx",
    version="6.2.1",
    description="Markdown to Powerpoint Converter",
    author="Martin Packman",
    author_email="martin@md2pptx.com",
    packages=find_packages(where="src"),
    package_dir={"": "src"},
    include_package_data=True,
    install_requires=[
        "python-pptx~=1.0.2",
        "lxml",
    ],
    extras_require={
        "full": [
            "cairosvg",
            "Pillow",
            "graphviz",
        ],
    },
    scripts=["md2pptx"],
    data_files=[
        ("share/md2pptx", ["Martin Template.pptx"]),
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires=">=3.8",
)

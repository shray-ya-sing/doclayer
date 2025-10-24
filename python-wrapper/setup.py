from setuptools import setup, find_packages
import os

# Read the contents of README file
this_directory = os.path.abspath(os.path.dirname(__file__))
with open(os.path.join(this_directory, 'README.md'), encoding='utf-8') as f:
    long_description = f.read()

setup(
    name="doclayer-python",
    version="0.1.0",
    author="Your Name",
    author_email="your.email@example.com",
    description="Python client library for DocLayer OOXML PowerPoint generation",
    long_description=long_description,
    long_description_content_type="text/markdown",
    url="https://github.com/yourusername/doclayer",
    packages=find_packages(),
    classifiers=[
        "Development Status :: 4 - Beta",
        "Intended Audience :: Developers",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Topic :: Office/Business :: Office Suites",
        "Topic :: Software Development :: Libraries :: Python Modules",
    ],
    python_requires=">=3.8",
    install_requires=[
        "pythonnet>=3.0.0",
    ],
    extras_require={
        "dev": ["pytest>=6.0", "black", "flake8", "mypy"],
    },
    include_package_data=True,
    package_data={
        "doclayer_python": ["bin/*.dll", "bin/*.json"],
    },
    keywords="powerpoint pptx openxml office documents ai agents",
    project_urls={
        "Bug Reports": "https://github.com/yourusername/doclayer/issues",
        "Source": "https://github.com/yourusername/doclayer",
        "Documentation": "https://doclayer.readthedocs.io/",
    },
)
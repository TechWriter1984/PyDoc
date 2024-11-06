from setuptools import setup, find_packages

setup(
    name="PyDoc",
    version="0.1.0",
    author="TechWriter1984",
    author_email="oopswow@126.com",
    description="A Python library for processing and translating Word documents.",
    long_description=open("README.md").read(),
    long_description_content_type="text/markdown",
    url="https://github.com/TechWriter1984/PyDoc/",
    packages=find_packages(),
    install_requires=[
        "python-docx",
        "requests"
    ],
    classifiers=[
        "Programming Language :: Python :: 3",
        "License :: OSI Approved :: MIT License",
        "Operating System :: OS Independent",
    ],
    python_requires='>=3.6',
)

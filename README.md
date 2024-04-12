# PDF-Merger

![Python application](https://img.shields.io/badge/python-application-blue.svg)

The PDF Merger application is a desktop tool designed to efficiently combine multiple PDF documents into one. Built with PyQt5 for a seamless GUI experience and leveraging PyPDF2 for backend PDF processing, it simplifies the task of PDF merging with drag-and-drop functionality.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Usage](#usage)
- [Contributing](#contributing)
- [Credits](#credits)
- [License](#license)

## Features

- **Drag-and-Drop Interface**: Easily add PDF files to the application.
- **Progress Indication**: Visually tracks progress with a status bar during the merge process.
- **Simple Merging**: Combine your PDF files into a single document with the click of a button.
- **Interactive List Management**: Review and clear the list of PDFs to be merged as needed.

## Installation

### Prerequisites

- Python (3.6 or newer) installed on your system.
- pip3 (Python package installer).

### Setup

1. Clone the repository to your local machine:

    ```bash
    git clone https://github.com/Dungyy/PDF-Merger.git
    ```

2. Navigate to the cloned directory:

    ```bash
    cd PDF-Merger
    ```
    
3. Use env for Python packages

   ```sh 
   python -m venv env
   source env/bin/activate

4. Install the required Python packages:

    ```bash
    pip install PyQt5 PyPDF2
    ```

    Alternatively, if you have a `requirements.txt` file:

    ```bash
    pip install -r requirements.txt  
    ```

## Usage

To run the PDF Merger application:

1. Navigate to your project directory.
2. Execute the script:

    ```bash
    python merger.py
    ```

## Contributing

We welcome contributions to the PDF Merger project! Please consider contributing in the following ways:

- Reporting a bug
- Discussing the current state of the code
- Submitting a fix
- Proposing new features

## Credits

This software uses the following open-source packages:

- [PyQt5](https://riverbankcomputing.com/software/pyqt/intro)

## License

This project is open-source and available for anyone to use at there own risk.
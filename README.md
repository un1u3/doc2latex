# DOCX to LaTeX Converter

A C++ tool to convert Microsoft Word documents (`.docx`) into LaTeX (`.tex`) files.  
Supports basic formatting: **bold**, *italic*, _underline_, and headings.

---

## Features

- Converts headings to LaTeX `\section`, `\subsection`, etc.
- Handles **bold**, *italic*, and _underline_ formatting.
- Escapes special LaTeX characters automatically.
- Simple command-line interface: input DOCX → output LaTeX.

---

## Requirements

- C++17 compatible compiler
- [libzip](https://libzip.org/) (for reading DOCX)
- [TinyXML2](https://github.com/leethomason/tinyxml2) (for XML parsing)
- CMake (for building)

---

## Project Structure

```
DocxToLatexConverter/
├── src/ # Source code
│ └── doc2latex.cpp
├── include/ # Header files (optional if split)
├── build/ # Build directory (ignored in Git)
├── docs/ # Sample DOCX files
│ └── input.docx
├── CMakeLists.txt
└── README.md

```


---

## Build Instructions

```bash
# Clone the repository
git clone https://github.com/<your-username>/DocxToLatexConverter.git
cd DocxToLatexConverter

# Create build directory and compile
mkdir build && cd build
cmake ..
make
```

### Usages


# Run the converter


`./doc2latex ../docs/input.docx [output.tex]`

# Example:
`./doc2latex ../docs/input.docx output.tex`

# Compile the resulting LaTeX file
`pdflatex output.tex`

<i>If [output.tex] is not provided, the default output file will be output.tex.</i>

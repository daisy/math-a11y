# Math Symbols to SVG using the STIX Two fonts

This project uses Node.js and the opentype.js library for font processing.

## Prerequisites

- [Node.js](https://nodejs.org/) installed on your system

## Setup

1. Clone this repository or simply download the stix_two_math_to_svg.js file into a folder
2. Download the latest STIXTwoMath-Regular.otf font file from https://www.stixfonts.org/
3. Initialize the Node.js project:

```bash
npm init -y
```

4. Install required dependency:

```bash
npm install opentype.js
```

## Project Structure

The project uses the following dependencies:

- `opentype.js`: For font file processing
- `fs`: Node.js built-in module for file system operations

## Running the Script

To run the script:

```bash
node stix_two_math_to_svg.js
```

SVG files will be generated in the same folder.

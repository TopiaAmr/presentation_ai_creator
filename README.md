# Presentation AI Creator

A Python tool for automatically generating PowerPoint presentations from structured data.

## Overview

This project provides a simple yet powerful way to programmatically create PowerPoint presentations using Python. It leverages the `python-pptx` library to convert structured JSON data into professionally formatted slides.

## Features

- Create title slides with subtitles
- Generate content slides with bullet points
- Support for indented bullet points (sub-points)
- Paragraph-based content slides
- Extensible architecture for adding more slide types

## Installation

1. Clone this repository:
   ```
   git clone https://github.com/yourusername/presentation_ai_creator.git
   cd presentation_ai_creator
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

The main functionality is provided by the `create_presentation_from_data()` function in `presentation_gen.py`. Here's a basic example:

```python
from presentation_gen import create_presentation_from_data

# Define your presentation structure
data = {
    "slides": [
        {
            "layout_type": "title_slide", 
            "title": "My Presentation", 
            "subtitle": "Created with AI"
        },
        {
            "layout_type": "title_content", 
            "title": "Key Points", 
            "content": [
                "First important point",
                "Second important point",
                "  Sub-point with indentation"
            ]
        },
        {
            "layout_type": "title_content", 
            "title": "Summary", 
            "content": "This is a paragraph-style slide with continuous text."
        }
    ]
}

# Generate the presentation
presentation = create_presentation_from_data(data)

# Save the presentation
presentation.save("my_presentation.pptx")
```

## Data Structure

The input data should be a dictionary with a "slides" key containing a list of slide definitions. Each slide definition should include:

- `layout_type`: The type of slide layout to use (e.g., "title_slide", "title_content", "blank")
- `title`: The title text for the slide
- Additional fields depending on the slide type:
  - `subtitle` for title slides
  - `content` for content slides (can be a string or a list of strings for bullet points)

## License

[MIT License](LICENSE)

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

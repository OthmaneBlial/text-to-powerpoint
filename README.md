# Text to PowerPoint

Generate PowerPoint presentations from structured text input with customizable templates.

## Features

- Convert text to PowerPoint slides
- Multiple slide templates
- User-friendly GUI
- Cross-platform compatibility

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/text-to-powerpoint.git
   cd text-to-powerpoint
   ```

2. Create and activate a virtual environment:
   ```
   python -m venv myenv
   source myenv/bin/activate  # On Windows: myenv\Scripts\activate
   ```

3. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   - Mac/Linux: `python src/app.py`
   - Windows: `python src\app.py`

2. Enter your slide content in the text area. Use this format:
   - Start with a title slide (two lines starting with `#`)
   - Use `#` for new slide titles
   - Use `-` or `•` for bullet points

Example input:
```
# Introduction to Python
- A powerful, versatile programming language
# Key Python Features
- Easy-to-read syntax
- Large standard library
- Cross-platform compatibility
# Data Structures in Python
- Lists, tuples, and dictionaries
- Flexible and dynamic
- Used for storing and managing data
# Python Functions
- Defined using `def` keyword
- Encourages code reuse
- Supports recursion and lambda functions
# Why Learn Python?
- High demand in various industries
- Excellent for beginners and professionals
- Extensive community and resources
```

3. Choose a template from the dropdown menu.
4. Click "Generate PowerPoint" to create and save your presentation.

## Building Executable (Windows)

To create a standalone executable:

```
python setup.py build
```

The executable will be in the `build` directory.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

MIT

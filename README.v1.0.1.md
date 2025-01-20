# update_tools_v1.0.1.py

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/downloads/) [![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

## Overview

`update_tools_v1.0.1.py` is a Python script designed to automate the updating of various package managers and tools on Windows systems. It sequentially executes update commands for tools like Winget, Chocolatey, Pip, and Npm, logging the outcomes to a file for easy review and debugging.

## Features

- **Automated Updates**: Executes update commands for multiple tools without manual intervention.
- **File Logging**: Outputs the update process to `update_log.txt`, capturing both success and error messages.
- **Extensibility**: Easily add or modify update commands within the script to accommodate additional tools.

## Prerequisites

Ensure that the following tools are installed on your system:

- [Winget](https://learn.microsoft.com/en-us/windows/package-manager/winget/)
- [Chocolatey](https://chocolatey.org/install)
- [Pip](https://pip.pypa.io/en/stable/installation/)
- [Npm](https://docs.npmjs.com/downloading-and-installing-node-js-and-npm)

## Installation

1. **Clone the Repository**:
   ```bash
   git clone https://github.com/EndeavorEverlasting/Python.git
   cd Python
   ```

2. **Verify Python Environment**:
   Ensure you have Python 3.8 or higher installed. You can verify your Python version by running:
   ```bash
   python --version
   ```

## Usage

1. **Run the Script**:
   Execute the script using Python:
   ```bash
   python update_tools_v1.0.1.py
   ```

2. **Review the Log**:
   Upon completion, the script generates an `update_log.txt` file in the same directory. This log contains detailed information about the update process, including success messages and any errors encountered.

## Extending the Script

To add a new tool or modify existing update commands:

1. Open `update_tools_v1.0.1.py` in a text editor.
2. Locate the `update_commands` dictionary.
3. Add or modify entries in the format:
   ```python
   "ToolName": "update_command"
   ```
   For example:
   ```python
   "NewTool": "newtool update --all"
   ```

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your enhancements or bug fixes. Ensure that your code adheres to [PEP 8 standards](https://www.python.org/dev/peps/pep-0008/) and includes appropriate logging for any new functionality.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Acknowledgments

This script was inspired by the need for a unified tool updater on Windows systems, simplifying the maintenance of development environments.

---
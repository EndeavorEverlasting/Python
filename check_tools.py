import subprocess  # Import the subprocess module to run shell commands

# Dictionary of commands to check for the presence and version of various tools
commands = {
    "Python": "python --version",  # Command to check Python version
    "Winget": "winget --version",  # Command to check Winget version
    "Chocolatey": "choco --version",  # Command to check Chocolatey version
    "Scoop": "scoop --version",  # Command to check Scoop version
    "Pip": "pip --version",  # Command to check Pip version
    "Npm": "npm --version",  # Command to check Npm version
}

def check_tool(name, command):
    """
    Executes a shell command to check if a tool is installed and returns its version.
    
    Parameters:
    - name: The name of the tool.
    - command: The shell command to execute.
    
    Returns:
    - A string indicating the tool's version or its installation status.
    """
    try:
        # Execute the command using subprocess.run
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            # If the command was successful, return the tool's version
            return f"{name}: {result.stdout.strip()}"
        else:
            # If the command failed, indicate the tool is not installed
            return f"{name}: Not installed"
    except Exception as e:
        # Catch any exceptions and return an error message
        return f"{name}: Error checking - {e}"

def main():
    """
    Main function to check for the presence of package managers and tools.
    """
    print("Checking for package managers and tools...\n")
    # Iterate over each tool and its command in the commands dictionary
    for name, command in commands.items():
        # Print the result of the check_tool function for each tool
        print(check_tool(name, command))

if __name__ == "__main__":
    # Entry point of the script
    main()


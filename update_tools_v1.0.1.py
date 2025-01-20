import subprocess  # Import the subprocess module to run shell commands

# Dictionary containing commands to update various tools
update_commands = {
    "Winget": "winget upgrade --all -q",  # Command to update all installed apps via Winget
    "Chocolatey": "choco upgrade all -y",  # Command to upgrade all Chocolatey packages
    "Pip": "python -m pip install --upgrade pip",  # Command to upgrade pip itself
}

def run_command(name, command, log_file):
    """
    Executes a shell command to update a tool and handles the output.

    Parameters:
    - name (str): The name of the tool to update.
    - command (str): The shell command to execute for updating the tool.
    - log_file (file object): The file object to write the output to.

    This function writes the result of the update process to the log file, indicating success or failure.
    """
    try:
        log_file.write(f"\nUpdating {name}...\n")  # Log the start of the update process
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            log_file.write(f"{name} updated successfully:\n{result.stdout}\n")
        else:
            log_file.write(f"{name} update failed:\n{result.stderr}\n")
    except Exception as e:
        log_file.write(f"Error updating {name}: {e}\n")

def update_npm(log_file):
    """
    Custom update function for Npm, including a version check for Node.js.

    Parameters:
    - log_file (file object): The file object to write the output to.
    """
    try:
        # Check Node.js version
        log_file.write("\nChecking Node.js version for Npm compatibility...\n")
        node_result = subprocess.run("node -v", shell=True, capture_output=True, text=True)
        if node_result.returncode != 0:
            log_file.write("Node.js is not installed or not found in PATH. Please install Node.js.\n")
            return

        node_version = node_result.stdout.strip().lstrip('v')
        major_version = int(node_version.split('.')[0])

        # Check compatibility with npm@11.0.0
        if major_version < 20 or (major_version >= 21 and major_version < 22):
            log_file.write(f"Node.js version {node_version} is incompatible with npm@11.0.0.\n")
            log_file.write("Please update Node.js to ^20.17.0 or >=22.9.0 before updating Npm.\n")
        else:
            # Update npm
            log_file.write(f"Node.js version {node_version} is compatible. Proceeding with Npm update...\n")
            npm_result = subprocess.run("npm install -g npm", shell=True, capture_output=True, text=True)
            if npm_result.returncode == 0:
                log_file.write(f"Npm updated successfully:\n{npm_result.stdout}\n")
            else:
                log_file.write(f"Npm update failed:\n{npm_result.stderr}\n")
    except Exception as e:
        log_file.write(f"Error updating Npm: {e}\n")

def main():
    """
    Main function to initiate the update process for all tools listed in update_commands.

    This function iterates over each tool and its corresponding update command,
    executing the update and logging the results.
    """
    with open("update_log.txt", "w") as log_file:
        log_file.write("Starting update process for tools...\n\n")  # Log the start of the update process
        
        # Iterate over each tool and its command in the update_commands dictionary
        for name, command in update_commands.items():
            run_command(name, command, log_file)

        # Handle Npm separately due to version requirements
        update_npm(log_file)

        # Log a note for Chocolatey updates
        log_file.write("\nNOTE: If Chocolatey was updated, restart your shell or run 'refreshenv' to apply changes.\n")

        log_file.write("\nUpdate process complete!\n")  # Log the completion of the update process

if __name__ == "__main__":
    main()
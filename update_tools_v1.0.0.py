import subprocess  # Import the subprocess module to run shell commands

# Dictionary containing commands to update various tools
update_commands = {
    "Winget": "winget upgrade --all -q",  # Command to update all installed apps via Winget
    "Chocolatey": "choco upgrade all -y",  # Command to upgrade all Chocolatey packages
    "Pip": "pip install --upgrade pip",  # Command to upgrade pip itself
    "Npm": "npm install -g npm",  # Command to update npm globally
}

def run_command(name, command):
    """
    Executes a shell command to update a tool and handles the output.

    Parameters:
    - name (str): The name of the tool to update.
    - command (str): The shell command to execute for updating the tool.

    This function prints the result of the update process, indicating success or failure.
    """
    try:
        print(f"\nUpdating {name}...")  # Inform the user that the update process has started
        # Run the command using subprocess.run
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            # If the command was successful, print the success message and output
            print(f"{name} updated successfully:\n{result.stdout}")
        else:
            # If the command failed, print the error message and output
            print(f"{name} update failed:\n{result.stderr}")
    except Exception as e:
        # Catch any exceptions and print an error message
        print(f"Error updating {name}: {e}")

def main():
    """
    Main function to initiate the update process for all tools listed in update_commands.
    
    This function iterates over each tool and its corresponding update command,
    executing the update and printing the results.
    """
    print("Starting update process for tools...\n")  # Inform the user that the update process is starting
    # Iterate over each tool and its command in the update_commands dictionary
    for name, command in update_commands.items():
        # Call run_command to execute the update for each tool
        run_command(name, command)
    print("\nUpdate process complete!")  # Inform the user that the update process is complete

if __name__ == "__main__":
    # Entry point of the script
    main()

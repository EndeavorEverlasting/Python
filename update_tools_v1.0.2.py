import subprocess
from datetime import datetime
import platform

# Dictionary containing commands to update various tools
update_commands = {
    "Winget": "winget upgrade --all -q",  # Command to update all installed apps via Winget
    "Chocolatey": "choco upgrade all -y",  # Command to upgrade all Chocolatey packages
    "Pip": "python -m pip install --upgrade pip",  # Command to upgrade pip itself
}

# To store results for summary
results = {"success": [], "failure": []}

def log_message(log_file, message_type, message):
    """
    Writes a message to the log file with a timestamp.

    Parameters:
    - log_file (file object): The file object to write the message to.
    - message_type (str): The type of the message (INFO, SUCCESS, ERROR, DEBUG).
    - message (str): The message to write.
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    log_file.write(f"[{timestamp}] [{message_type}] {message}\n")

def run_command(name, command, log_file):
    """
    Executes a shell command to update a tool and logs detailed output.

    Parameters:
    - name (str): The name of the tool to update.
    - command (str): The shell command to execute for updating the tool.
    - log_file (file object): The file object to write the output to.
    """
    try:
        log_message(log_file, "INFO", f"Starting update for {name}. Command: {command}")
        result = subprocess.run(command, shell=True, capture_output=True, text=True)

        if result.returncode == 0:
            log_message(log_file, "SUCCESS", f"{name} updated successfully.")
            log_message(log_file, "DEBUG", f"Output:\n{result.stdout}")
            results["success"].append(name)
        else:
            log_message(log_file, "ERROR", f"{name} update failed.")
            log_message(log_file, "DEBUG", f"Error Output:\n{result.stderr}")
            results["failure"].append(name)
    except Exception as e:
        log_message(log_file, "ERROR", f"Exception during {name} update: {e}")
        results["failure"].append(name)

def update_npm(log_file):
    """
    Custom update function for Npm, including a version check for Node.js.

    Parameters:
    - log_file (file object): The file object to write the output to.
    """
    try:
        log_message(log_file, "INFO", "Checking Node.js version for Npm compatibility...")
        node_result = subprocess.run("node -v", shell=True, capture_output=True, text=True)
        if node_result.returncode != 0:
            log_message(log_file, "ERROR", "Node.js is not installed or not found in PATH.")
            results["failure"].append("Npm")
            return

        node_version = node_result.stdout.strip().lstrip('v')
        major_version = int(node_version.split('.')[0])

        if major_version < 20 or (major_version >= 21 and major_version < 22):
            log_message(log_file, "ERROR", f"Node.js version {node_version} is incompatible with npm@11.0.0.")
            log_message(log_file, "INFO", "Please update Node.js to ^20.17.0 or >=22.9.0.")
            results["failure"].append("Npm")
        else:
            log_message(log_file, "INFO", f"Node.js version {node_version} is compatible. Proceeding with Npm update...")
            npm_result = subprocess.run("npm install -g npm", shell=True, capture_output=True, text=True)
            if npm_result.returncode == 0:
                log_message(log_file, "SUCCESS", "Npm updated successfully.")
                log_message(log_file, "DEBUG", f"Output:\n{npm_result.stdout}")
                results["success"].append("Npm")
            else:
                log_message(log_file, "ERROR", "Npm update failed.")
                log_message(log_file, "DEBUG", f"Error Output:\n{npm_result.stderr}")
                results["failure"].append("Npm")
    except Exception as e:
        log_message(log_file, "ERROR", f"Exception during Npm update: {e}")
        results["failure"].append("Npm")

def main():
    """
    Main function to initiate the update process for all tools listed in update_commands.

    This function logs system information, iterates over each tool, and provides a summary.
    """
    with open("update_log.txt", "w") as log_file:
        # Log system information
        log_message(log_file, "INFO", f"System Information: {platform.platform()}")
        log_message(log_file, "INFO", f"Python Version: {platform.python_version()}")
        
        # Log start of update process
        log_message(log_file, "INFO", "Starting update process for tools...\n")
        
        # Iterate over each tool and its command in the update_commands dictionary
        for name, command in update_commands.items():
            run_command(name, command, log_file)
        
        # Handle Npm separately due to version requirements
        update_npm(log_file)

        # Provide a summary of results
        log_message(log_file, "INFO", "\nUpdate Summary:")
        log_message(log_file, "INFO", f"Successful updates: {', '.join(results['success']) if results['success'] else 'None'}")
        log_message(log_file, "INFO", f"Failed updates: {', '.join(results['failure']) if results['failure'] else 'None'}")
        log_message(log_file, "INFO", "Update process complete!")

if __name__ == "__main__":
    main()

"""
File: auto_input.py
Author: Lam Wai Taing, Timothy
Date: 2024/07/19
Description: A Python script to run the TOV640 Exceptino Report program and automatically enter the inputs.
"""

import subprocess

# Function to simulate user input
def auto_input(process, inputs):
    """
    Writes the inputs automatically when prompted.

    Args:
        process (Popen[str]): Executes a child program in a new process.
        inputs (list[str]): A list of inputs.

    Returns:
        None
    """
    for input_value in inputs:
        print(f"Sending input: {input_value}")
        process.stdin.write((input_value + '\n'))
        process.stdin.flush()

# Define the inputs you want to provide to the script
inputs = [
    '',
    '',
    '',
    '',
    'EAL',  # Line
    'y',    # LMC
    # 'y',    # RAC
    # 'y',    # LOW S1
    # 'HUH-FOT'    # manually input section
    'UP',    # UP/DN
    ''
]

script_path = 'TOV640 exception generator/TOV640_exception_generator.py'

# Start the process
process = subprocess.Popen(['python', script_path], 
                           stdin=subprocess.PIPE, 
                           stdout=subprocess.PIPE,
                           stderr=subprocess.PIPE,
                           text=True)

# Simulate user input
auto_input(process, inputs)

# Read output
output, error = process.communicate()

print('Output:\n', output)
print('Error:\n', error)

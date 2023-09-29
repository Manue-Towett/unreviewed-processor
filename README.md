# unreviewed-processor
Moves matching rows to the bottom row of the Master file in the "Unreviewed matching products" tab

# Requires
- python 3.11+

# Settings
- There is a configuration file in settings directory:
    - use it to set:
        - input - the folder where the input excel files are located
        - master - the folder where the master excel file is located

# Usage
- Open the command prompt / terminal
- cd into the project folder / directory
- If running for the first time, first install dependencies using the command: 
    ```pip install -r requirements.txt```
- To run the script, use the commands:
    - For Linux/MacOS: 
        ```python3 main.py```
    - For windows: 
        ```python main.py```
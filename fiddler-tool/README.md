# Fiddler Tool

## Overview
The Fiddler Tool is a Python application designed to analyze `.saz` files using various scripts. It provides a graphical user interface (GUI) for users to select their `.saz` files and run analyses that check IP ownership and match hostnames against fetched URLs.

## Project Structure
```
fiddler-tool
├── src
│   ├── main.py            # Entry point of the application
│   ├── index.py           # Logic for processing the .saz file
│   ├── ariSearch.py       # Performs WHOIS lookups and checks IP ownership
│   ├── searchMsDocs.py    # Fetches URLs and checks hostname matches
│   └── utils
│       └── config.py      # Configuration settings
├── data
│   └── output
│       └── .gitkeep       # Placeholder for Git tracking
├── requirements.txt        # Python dependencies
├── .gitignore              # Files to ignore in Git
└── README.md               # Project documentation
```

## Requirements
To run this project, you need to have Python installed on your machine. Additionally, you will need the following Python packages:

- requests
- pandas
- colorama
- openpyxl
- tkinter

You can install the required packages using the following command:

```
pip install -r requirements.txt
```

## Usage
1. Run the application by executing the `main.py` file:
   ```
   python src/main.py
   ```
2. A GUI will appear prompting you to select a `.saz` file.
3. After selecting the file, the application will run the analysis scripts and display the results.

## Contributing
If you would like to contribute to this project, please fork the repository and submit a pull request with your changes.

## License
This project is licensed under the MIT License - see the LICENSE file for details.
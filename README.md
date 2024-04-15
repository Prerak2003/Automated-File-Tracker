# Excel File Updater README
## Introduction
Managing incoming data files in a business system can be a daunting task, especially when dealing with numerous files arriving at regular intervals. It becomes crucial to keep track of new data files and maintain a record of past uploads. The Excel File Updater script simplifies this process by automating the identification and tracking of new files in a designated directory.
If your business system receives data on a periodic basis and it's challenging to keep records of incoming files, this script is an ideal solution. It efficiently identifies newly added files compared to the previous records, making it easy to track and manage uploads.

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Setup](#setup)
3. [Usage](#usage)
4. [Dependencies](#dependencies)
5. [Contributing](#contributing)
6. [License](#license)

## Prerequisites
- Python 3.x installed on your system.
- Ensure `openpyxl` library is installed. If not, you can install it using:
 
## Setup
1. Clone or download this repository to your local machine.
2. Ensure your directory structure matches the one defined in the code or modify the `DIRECTORY_PATH` variable in the script according to your directory structure.
3. Ensure you have a Gmail account with app password enabled for sending emails.

## Usage
1. Place the Excel file you want to update in the specified directory.
2. Run the script `excel_file_updater.py`.
3. The script will update the Excel file with new file paths found in the directory and send an email with the updated Excel file as an attachment.

## Dependencies
- Python 3.x
- `openpyxl` library for handling Excel files.
- `smtplib` and `email` libraries for sending emails.

## Contributing
Contributions are welcome! If you find any issues or have suggestions for improvements, feel free to open an issue or create a pull request.

## License
This project is licensed under the [MIT License](LICENSE).


# Sales Data Automation

Automates the process of extracting sales and inventory data from UCD website. This tool automatically logs in, navigates through the UCD interface, and exports data to Excel files.

## Features

- Automated login to UCD website
- Extracts inventory data
- Extracts monthly supply reports
- Exports data to Excel files
- Automatic ChromeDriver management
- Error screenshots for troubleshooting

## Prerequisites

- Python 3.8 or higher
- Google Chrome browser
- macOS or Windows operating system
- UCD member account credentials

## Installation

1. Clone this repository:
```bash
git clone [your-repository-url]
cd Sales_Data_Automation
```

2. Set up configuration:
```bash
cd config
cp sample.ini config.ini
```

3. Edit `config.ini` with your credentials:
```ini
[Credentials]
website_url = https://www.ucd.com.tw
username = YOUR_UCD_USERNAME # Replace with your UCD username
password = YOUR_UCD_PASSWORD # Replace with your UCD password
[Settings]
timeout = 30 # Increase if you have slow internet
browser = chrome
```

4. Set proper file permissions (macOS only):
```bash
chmod 600 config/config.ini
```

## Usage

1. Activate the virtual environment:
```bash
#On macOS
./activate.sh
#On Windows
.\activate.bat # You'll need to create this for Windows 
```

2. Run the program:
```bash
python3 src/main.py
``` 

3. The program will:
   - Log in to UCD website
   - Extract inventory data
   - Extract monthly supply data
   - Export data to Excel files in the `exports/` directory
   - Keep the browser open for manual interaction
   - Type 'q' in the terminal to quit and close the browser

## Project Structure

Sales_Data_Automation/
├── config/
│ ├── sample.ini
│ ├── config.ini (created by user)
│ └── urls.py
├── src/
│ ├── main.py
│ ├── web_navigator.py
│ └── logger_config.py
├── exports/
│ └── (generated Excel files)
├── error_screenshots/
│ └── (error screenshots if any)
├── venv/
├── activate.sh
└── README.md

## Security Notes

- Never commit `config.ini` to version control
- Keep your UCD credentials secure
- The program uses incognito mode and clears browser data after use
- Screenshots are saved locally for troubleshooting

## Troubleshooting

1. **Configuration Issues**:
   - Verify your UCD credentials
   - Check file permissions on `config.ini`
   - Ensure config file is in the correct location

2. **Browser Issues**:
   - Ensure Chrome is installed
   - Check internet connection
   - Look for error screenshots in `error_screenshots/` directory

3. **Virtual Environment Issues**:
   - Delete `venv` folder and rerun `activate.sh`
   - Verify Python version: `python --version`

## Error Handling

- Screenshots are automatically saved when errors occur
- Check the terminal output for error messages
- Screenshots are saved in `error_screenshots/` with timestamps

## Maintenance

- Regular updates to ChromeDriver are automatic
- Check for Python package updates periodically
- Monitor UCD website changes that might affect automation

## Contributing

1. Fork the repository
2. Create your feature branch
3. Commit your changes
4. Push to the branch
5. Create a new Pull Request

## License

[Your chosen license]

## Contact

[Your contact information]
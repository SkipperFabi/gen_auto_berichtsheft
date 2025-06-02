# Berichtsheft Generator

This script fetches teaching contents from WebUntis and generates a formatted Word document (`.docx`) for your Berichtsheft.

## Features

- Fetches lessons and teaching content from WebUntis for a given date range
- Groups teaching content by subject and day
- Highlights missing teaching content in red
- Outputs a `.docx` file to a configurable directory

## Prerequisites

- [Node.js](https://nodejs.org/) (v16 or newer recommended)
- npm (comes with Node.js)
- Access to your school's WebUntis account

## Setup

1. **Clone the repository:**
   ```bash
   git clone https://github.com/skipperfabi/gen_auto_berichtsheft.git
   cd gen_auto_berichtsheft
   ```

2. Install dependencies:
npm install

3. Create a .env file in the project root:
Copy or rename the .example.env to .env and fill in your credentials and preferences.

- DOCX_PATH (optional): Directory for the output file. If not set, the script directory is used.
- OUTPUT_FILENAME (optional): Name of the output file. Defaults to TeachingContentOverview.docx.
- DEBUG (optional): Set to true for verbose logging.

## Usage
1. Run the script (bash or cmd)
node gen_auto_berichtsheft.js

2. Follow the prompts:
```javscript
    Enter the start date (YYYY-MM-DD)
    Enter the end date (YYYY-MM-DD)
```

3. Find your generated .docx file:
The file will be saved to the directory specified by DOCX_PATH or, if not set, in the script directory.

## Notes
- If a lesson does not have teaching content, the script will insert a red warning in the document: "Kein Lehrstoff durch die Lehrkraft angegeben! Bitte manuell eintragen!"

- If a lesson is cancelled, it will be marked as Entfallen (startTime - endTime) in the document.

- Make sure your WebUntis credentials are correct and have access to the required data.

## Troubleshooting
If you encounter errors related to missing modules, run npm install again.
For authentication or data issues, double-check your .env file.
For further debugging, set DEBUG=true in your .env.

## License
[MIT](https://choosealicense.com/licenses/mit/)
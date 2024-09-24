# Chat-GPT-to-Word-Document-automation

# Project Title:
Automated ChatGPT to Microsoft Word Integration

## Description:
This project automates the process of querying ChatGPT and transferring the response directly into Microsoft Word. Using `pyautogui`, the script opens a web browser, interacts with the ChatGPT interface, retrieves the generated response, and pastes it into a Word document. The filename of the Word document is based on the user’s prompt, ensuring a clear association between the question and the saved document.

## Motivation:
This automation aims to streamline the interaction with ChatGPT by reducing manual steps and increasing efficiency. It allows users to easily save responses for later reference without needing to copy and paste manually.

## Key Features:
1. **User-Friendly Input**: Prompts the user for a question or prompt directly via a pop-up interface.
2. **Automated Browser Control**: Opens the Chrome browser, navigates to ChatGPT, and inputs the user’s question.
3. **Response Retrieval**: Waits for the response to be generated, then copies it to the clipboard.
4. **Word Document Creation**: Opens Microsoft Word, pastes the response, and saves it with the user's prompt as the filename.
5. **Seamless Integration**: No need for manual copying and pasting; the entire process is automated.

## Code Overview:
### Script 1:
This version uses hardcoded coordinates to interact with the applications, including opening the browser and closing the activation wizard in Word.

### Script 2:
This updated version uses hotkeys to enhance reliability across different screen resolutions, improving the script's adaptability and robustness.

## Performance:
The automation process runs effectively with minimal user intervention. The expected run time depends on the response time from ChatGPT and the speed of the user's system.

## Technologies Used:
- **Programming Language**: Python
- **Libraries**: 
  - `pyautogui` for GUI automation
  - `time` for managing delays and waiting times

## Future Improvements:
- Implement error handling to manage cases where applications do not respond as expected.
- Explore additional features such as formatting options in Word.
- Consider integrating options for saving to different file formats or cloud services.

## Usage:
1. **Install Required Libraries**: Ensure that `pyautogui` is installed in your Python environment. You can install it via pip:
   ```bash
   pip install pyautogui

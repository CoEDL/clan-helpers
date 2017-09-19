# Running the CLAN header generator

1. Install Python 2.7 and pyexcel
2. Run the script by either:
	1. Command line - Navigate to this folder in command prompt and type:
	```python generate_headers.py```
	OR
	2. Running the file in Python IDLE
3. At the prompt, select the excel file to use. A sample spreadsheet can be found in the `sample/` folder. It must have a `Sessions` and a `Participants` sheet with at least the following headers:
	- Sessions:
		- Session name
		- Length of audio
		- Date recorded (DD/MM/YYYY)
		- Location info
		- Speakers
		- Language
		- Investigator
		- Activity
		- Transcribed by
	- Participants:
		- Participant Surname
		- Participant First Name
		- CLAN Code
		- Sex
		- Languages (order of proficiency)
		- Usual role in recordings
4. The header files will be generated in an Output folder.

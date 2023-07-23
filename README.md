# Quiz-Formatter
Formats a quiz that was copy/pasted from online into a Word document, be it from a forum or AI generated.

A new, formatted document will be saved in the same directory as the original, named "[original_name]_formatted.docx"

1. Removes all styling (bold, italic, underlines) from the page
2. Homogenizes the font style and size to Arial 12pt
3. Bolds each question
4. If there are multiple choice options, space them below the question and indent them
5. Add space between each question
6. If there is an answer key, move it to the next page

*Note that this script will likely not flawlessly format your document, but it will do most of the heavy lifting.*

Works best with multiple choice quizzes.

	• This script assumes that all questions end with either a "?" or a ":".
	• Questions that do not end in "?" or ":" will not be bolded
	• Any multiple choice options that contain "?" or ":" will not format correctly

To use the script:

	• Ensure python is installed on your computer (https://www.python.org/downloads/)
	• Run the script in an IDE and follow the instructions, or:
	• Open your command line
	• Change the directory to the folder this script is stored using the cd command
        ex. "C:\Users\emily>cd C:\Users\emily\OneDrive\Documents"
	• Type "python [name of this file].py"
        ex. "C:\Users\emily\OneDrive\Documents>python format.py"
	• Follow the instructions

# Examples

UNFORMATTED DOCUMENT:

![Picture of an unformatted document.](https://github.com/roseemi/Quiz-Formatter/blob/main/images/Original_Document.png)


FORMATTED DOCUMENT USING THIS SCRIPT:

![Picture of formated document.](https://github.com/roseemi/Quiz-Formatter/blob/main/images/Formatted_Document.png)


UNFORMATTED DOCUMENT that DOES NOT meet criteria to format correctly:

![Picture of an unformated document.](https://github.com/roseemi/Quiz-Formatter/blob/main/images/Original_Document_2.png)


FORMATTED DOCUMENT USING THIS SCRIPT:

![Picture of an improperly formated document.](https://github.com/roseemi/Quiz-Formatter/blob/main/images/Formatted_Document_2.png)

The issue in this case: The script assumes that "?" is always the end of a sentence. Because the PHP answers use question marks, the formatting fails.

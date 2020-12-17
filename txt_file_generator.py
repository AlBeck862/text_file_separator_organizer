"""
*****DOCUMENTATION*****
Enhancing Computer Skills - Master Word file to text file conversion script
Build Version: 2.3.1
Last Updated: 11/07/2020 (dd/mm/yyyy)
Written by Alexander Becker
Usage: python3 <script name>

User information and disclaimers:
Thank you for using this script.
Refer to the above "usage" statement for information on initializing the script.
Refer to the Marker's Guide for more information on the required Master Word file formatting.
Please note that by using this script you acknowledge that the script may not return an expected result.
Nevertheless, efforts have been made to ensure a smooth and consistent process.
The author is in no way responsible for script glitching, errors, or any other failing.
For questions, comments and concerns, please contact the author at the email address provided by the course instructor.

Known bugs:
(1) The first text file in each folder will contain the name of the student.
(2) When prompted for folder names, if incorrect names are entered and 'yes' is entered on confirmation, the script will crash with a NameError.
*****DOCUMENTATION*****
"""

import docx2txt
import sys
import os

# Help and documentation request
print("If this is your first time using this script, it is highly recommended that you view the help and documentation.")
response = input("Would you like to see the help and documentation? (yes/no): ")

# Help and documentation display
while True:
	if response == "yes":
		print(__doc__)
		print("*****HELP*****")
		print("First, create four folders. One folder for each section.")
		print("For example, W945, W1245, and so on.")
		print("Next, place those folders and the Word document in the same folder as this script.")
		print("For Word document formatting instructions, refer to the Marker's Guide.")
		print("*****HELP*****\n")
		break
	elif response == "no":
		break
	else:
		print("Invalid response. Please enter 'yes' or 'no' without quotation marks.")
		response = input("Would you like to see the help and documentation? (yes/no): ")

# Input Word document name and extension
# External package process: convert docx to processable information
while True:
	try:
		document = input("Enter the name of the Word file including the extension (ex.: test.docx): ")
		text = docx2txt.process(document)
		break
	except FileNotFoundError:
		print("\nIt looks like that file doesn't exist!")
		print("Check that you entered the name correctly and included the extension.")

# Collect assignment number
assg_num = input("Enter the assignment number as an integer: ")
while True:
	try:
		assg_num = int(assg_num)
		assg_num = str(assg_num)
		break
	except ValueError:
		print("\nIt looks like that wasn't an integer!")
		assg_num = input("Enter the assignment number as an integer: ")

# Collect folder names
while True:
	lec1 = input("Enter the first lecture divider, example, W945: ")
	lec2 = input("Enter the second lecture divider, example, W945: ")
	lec3 = input("Enter the third lecture divider, example, W945: ")
	lec4 = input("Enter the fourth lecture divider, example, W945: ")

	print(f"You entered, in order: {lec1}, {lec2}, {lec3}, {lec4}.")
	answer = input("Is this correct? (yes/no): ")

	if answer == "yes":
		break
	elif answer == "no":
		print("You will now be prompted to enter the four folder names again.")
	else:
		print("You did not answer 'yes' or 'no'.")
		print("As a precaution, you will now be prompted to enter the four folder names again.")

print("\nThank you! The script will now attempt to generate the text files.")

series = [lec1,lec2,lec3,lec4]
series_select = ["void",lec1,lec2,lec3,lec4,lec4]

# Course section counter
counter = 0

# Text file counter
tf_counter = 0

# generate a txt file using the stored data from the conversion
with open("masterTXT.txt", "w") as txtf:
	print(text, file=txtf)
	txtf.seek(0)

# store the master txt file data as a list
with open("masterTXT.txt") as txtf:
	data = txtf.readlines()
	end_index = 1

	# cycle through each line
	for idx, line in enumerate(data):

		# Increment the course section counter
		if any(i in line for i in series) or (data[idx] == data[-1]):
			counter += 1

		# Set the folder name to the appropriate section
		folder_name = series_select[counter]

		# Begin a new text file based on specific rules
		# Rule 1: if a section divider is encountered
		# Rule 2: if the end of the document is encountered
		# Rule 3: if five newline characters are encountered in a row
		if (any(i in line for i in series)) or (data[idx] == data[-1]) or (all("\n" == data[idx-int(j)] for j in ["1","2","3","4","5"])):
			start_index = end_index
			end_index = idx

			# Rule 2 case text file naming
			if all("\n" == data[idx-int(j)] for j in ["1","2","3","4","5"]) and (", " in data[start_index]):
				last_name,first_name = data[start_index].split(", ")
				last_name = last_name.replace("\n", "")
				first_name = first_name.replace("\n", "")

			# Rules 1 and 3 cases text file naming
			if ", " in data[start_index+2]:
				last_name,first_name = data[start_index+2].split(", ")
				last_name = last_name.replace("\n", "")
				first_name = first_name.replace("\n", "")
			
			# Ignore segments of insignificant size (don't generate useless files)
			if start_index < (end_index - 1):

					# Create a new list with individual text file contents
					new_doc_list = [data[i] for i in range(start_index+1,end_index)]
					final_text = "".join(new_doc_list)

					# Don't create text files with folder dividers
					if not any(i in final_text for i in series):

						# Generate an individual text file and write in the list information
						with open(__file__.replace("converter2.py","") + folder_name + "/" + "a" + str(assg_num) + "-" + last_name + "-" + first_name + ".txt", "w") as newtxt:
							newtxt.write(final_text)
							tf_counter += 1

os.remove("masterTXT.txt")
print(f"\n*****NOTICE*****\nProcess completed.\n{str(tf_counter)} text files were generated.\n*****NOTICE*****\n")
print("If the number of text files does not match the number of students, you may have used incorrect Word file formatting.")
print("Please consult the documentation for more information regarding correct Word file formatting.")
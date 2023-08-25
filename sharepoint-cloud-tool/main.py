# Import necessary libraries
import os

# Import internal utilities
from sharepoint_tools import get_folder_data, write_to_excel, copy_folder, copy_file, get_folder_files
from sharepoint_ui import display_greetings, provide_url, generate_options, display_options, generate_path, reset_options

# Define global variables
content = []
options = []
new_path = ""

# Define variables
excel_name = "SharePoint_Logs"
mapping_structure = {"Source_Folder_A": "Destination_Folder_A", 
                     "Source_Folder_B": "Destination_Folder_B",
                     "Source_Folder_C": "Destination_Folder_C",
                     "Source_Folder_D": "Destination_Folder_D"
                     }

# Print initial message
display_greetings("start")

# Take path from user input and check if it's valid
user_path = input("Provide directory path: ")
new_path = provide_url(user_path)

# Generate list of subfolders name
content = get_folder_data(new_path)

# Generate options for menu with generated content
generate_options(content)

# Display generated menu
display_options()

# Loop through user options
is_continue = True
while is_continue:

    # Take option from the user
    user_option = int(input("Choose option: "))

    # Exit from program if user press Exit option or press Back at the root level
    if user_option == 0 or (user_option == 1 and new_path == user_path):
        is_continue = False
        display_greetings("end")
        SystemExit

    # Display options if user go back
    elif user_option == 1:
        # Generate path for new selected directory
        generate_path(new_path, content, user_option)
        # Delete all previous options
        reset_options()
        # Read and store in a variable all subdirectories from current directory
        content = get_folder_data(new_path)
        # Generate options for menu with generated content
        generate_options(content)
        # Display generated menu
        display_options()

    # Write to Excel file with specified name at specified location
    elif user_option == 2:
        write_to_excel(excel_name, os.getcwd(), content, 1, 1)

    # Copy the content of the current directory in specified directory
    elif user_option == 3:
        # Ask user to introduce destination folder path and copy folder
        folder_destination = input("Introduce path of the destination folder: ")

        # Iterate through all subfolders name
        for folder_name in content:
            folder_source_path = new_path + "/" + folder_name

            # Set copying path according with mapping structure
            if folder_name in list(mapping_structure.keys()):
                folder_destination_path = folder_destination + "/" + mapping_structure[folder_name]
            else:
                folder_destination_path = folder_destination

            # Copy folder
            copy_folder(folder_source_path, folder_destination_path)

    # Copy files from current directory in specified directory
    elif user_option == 4:
        # Ask user to introduce destination path and copy all files
        files_name = get_folder_files(new_path)
        file_destination = input("Introduce path of the file destination: ")
        for file in files_name:
            file_source = new_path + "/" + file
            copy_file(file, file_source, file_destination)

    # Display options if user choose non-default option
    else:
        # Generate path for new selected directory
        generate_path(new_path, content, user_option)
        # Delete all previous options
        reset_options()
        # Read and store in a variable all subdirectories from current directory
        content = get_folder_data(new_path)
        # Generate options for menu with generated content
        generate_options(content)
        # Display generated menu
        display_options()


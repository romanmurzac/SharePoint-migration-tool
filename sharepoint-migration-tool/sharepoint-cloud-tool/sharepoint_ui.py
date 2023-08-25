# Import necessary libraries
import os

# Import necessary modules
from office365.sharepoint.client_context import ClientContext

# Import internal utilities
from sharepoint_tools import get_folder_data, get_folder_files, copy_folder, copy_file, write_to_excel, read_from_excel, log_to_excel
from credentials import username, password, site_url

# Define global variables
content = []
options = []
new_path = ""


# Define function for menu creation
def generate_options(content: list) -> list:
    """
    Generate menu options for current directory
    :param content: list of subdirectories name in string format
    :return: options: list of menu options
    """
    # Modify global variable
    global options

    # Generate menu options
    options.append("[0] Exit")
    options.append("[1] Back")
    options.append("[2] Store")
    options.append("[3] Copy Folder")
    options.append("[4] Copy File")
    options.append("[5] Structure Copy")
    new_option = ""

    # Append all subdirectories name to the options list
    for index, item in enumerate(content):
        new_option = f"[{index + 6}] {item}"
        options.append(new_option)
    
    # Return modified options variable
    return options


# Define function for menu display
def display_options() -> None:
    """
    Generate menu options for current directory
    :param None
    :return: None
    """
    # Display in console all available options
    for item in options:
        print(item)


# Define function for menu reset
def reset_options() -> None:
    """
    Delete all current options
    :param None
    :return: None
    """
    # Modify global variable
    global options

    # Clear current content of the options variable
    options.clear()


# Define function for display greetings
def display_greetings(state: str) -> None:
    """
    Display specific start and end message
    :param state: specify position of the message in string format
    :return: None
    """
    if state == "start":
        print("**************************")
        print("***** MIGRATION TOOL *****")
        print("**************************\n")
    else:
        print("*******************************")
        print("***** EXIT MIGRATION TOOL *****")
        print("*******************************\n")


# Define function for URL validation
def provide_url(user_input: str) -> str:
    """
    Validate path introduced by user
    :param user_input: introduced path in string format
    :return: user_input: introduced path in string format
    """
    # Loop for existing path providing
    is_not_valid_path = True
    while is_not_valid_path:
        
        # Create connection to the SharePoint and check if path exists
        context = ClientContext(site_url).with_user_credentials(username, password)
        is_valid_folder_path = context.web.get_folder_by_server_relative_url(user_input).get().execute_query().exists

        # Exit from loop if path exists
        if is_valid_folder_path:
            is_not_valid_path = False
    
    # Return valid SharePoint path
    return user_input


# Define function for path generation
def generate_path(path: str, content: list, index: int) -> str:
    """
    Generate path for the selected subdirectory
    :param path: path of the current directory in string format
    :param content: list of subdirectories name in string format
    :param index: chosen integer option from menu
    :return: new_path: path of the desired subdirectory in string format
    """
    # Modify global variable
    global new_path
    
    # Generate path for specific subdirectory depends of the chosen option
    if index > 1:
        new_path = path + "/" + content[index - 6]
    else:
        temp_path = path.split("/")[:-1]
        new_path = "/".join(temp_path)
    
    # Return modified path variable
    return new_path


# Run code as a script
if __name__ == "__main__":

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
    print("\n" + "*" * 10)
    content = get_folder_data(new_path)
    files = get_folder_files(new_path)
    print("*" * 10)

    # Generate options for menu with generated content
    generate_options(content)

    # Display generated menu
    display_options()
    print("*" * 10 + "\n")

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
            print("\n" + "*" * 10)
            content = get_folder_data(new_path)
            files = get_folder_files(new_path)
            print("*" * 10)
            # Generate options for menu with generated content
            generate_options(content)
            # Display generated menu
            display_options()
            print("*" * 10 + "\n")

        # Write to Excel file with specified name at specified location
        elif user_option == 2:
            write_to_excel(excel_name, os.getcwd(), content, 1, 1)

        # Copy the content of the current directory in specified directory
        elif user_option == 3:
            # Ask user to introduce destination folder path and copy folder
            folder_destination = input("Introduce path of the destination folder: ")
            # Copy folder
            copy_folder(new_path, folder_destination)

        # Copy files from current directory in specified directory
        elif user_option == 4:
            # Ask user to introduce destination path and copy all files
            files_name = get_folder_files(new_path)
            file_destination = input("Introduce path of the file destination: ")
            for file in files_name:
                file_source = new_path + "/" + file
                copy_file(file, file_source, file_destination)

        # Copy the content of the current directory in specified directory
        elif user_option == 5:
            # Ask user to introduce destination folder path and copy folder
            folder_destination = input("Introduce path of the destination folder: ")

            # Copy all loose files from main folder
            files_name = get_folder_files(new_path)
            file_destination = folder_destination
            for file in files_name:
                file_source = new_path + "/" + file
                copy_file(file, file_source, file_destination)
                # Store files path in Excel file
                log_to_excel(excel_name, os.getcwd(), file_source, 1, 1)
                log_to_excel(excel_name, os.getcwd(), file_destination, -1, 2)

            # Iterate through all subfolders name
            for folder_name in content:
                folder_source_path = new_path + "/" + folder_name
                # Set copying path according with mapping structure
                if folder_name in list(mapping_structure.keys()):
                    folder_destination_path = folder_destination + "/" + mapping_structure[folder_name]

                    # Copy all loose files from subfolder
                    files_name = get_folder_files(folder_source_path)
                    file_destination = folder_destination_path
                    for file in files_name:
                        file_source = folder_source_path + "/" + file
                        copy_file(file, file_source, file_destination)
                        # Store files path in Excel file
                        log_to_excel(excel_name, os.getcwd(), file_source, 1, 1)
                        log_to_excel(excel_name, os.getcwd(), file_destination, -1, 2)
                    
                    content = get_folder_data(folder_source_path)
                    for sub_dir in content:
                        sub_dir_source_path = folder_source_path + "/" + sub_dir
                        # Copy all subfolders
                        copy_folder(sub_dir_source_path, folder_destination_path)
                        # Store files path in Excel file
                        log_to_excel(excel_name, os.getcwd(), sub_dir_source_path, 1, 1)
                        log_to_excel(excel_name, os.getcwd(), folder_destination_path, -1, 2)
                else:
                    folder_destination_path = folder_destination
                    # Copy folder
                    copy_folder(folder_source_path, folder_destination_path)
                    # Store files path in Excel file
                    log_to_excel(excel_name, os.getcwd(), folder_source_path, 1, 1)
                    log_to_excel(excel_name, os.getcwd(), folder_destination_path, -1, 2)

        # Display options if user choose non-default option
        else:
            # Generate path for new selected directory
            generate_path(new_path, content, user_option)
            # Delete all previous options
            reset_options()
            # Read and store in a variable all subdirectories from current directory
            print("\n" + "*" * 10)
            content = get_folder_data(new_path)
            files = get_folder_files(new_path)
            print("*" * 10)
            # Generate options for menu with generated content
            generate_options(content)
            # Display generated menu
            display_options()
            print("*" * 10 + "\n")
    
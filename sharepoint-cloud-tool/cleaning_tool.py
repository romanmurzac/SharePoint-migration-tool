# Import necessary modules
from office365.sharepoint.client_context import ClientContext

# Import internal utilities
from credentials import username, password, site_url
from sharepoint_tools import get_folder_data, get_folder_files

# Create connection to the SharePoint
context = ClientContext(site_url).with_user_credentials(username, password)


# Define function for deleting folder
def delete_folder(folder_path: str) -> None:
    """
    Delete folder
    :param folder_path: folder path in string format
    :return: None
    """
    # Define folder to be deleted and delete it
    folder_to_delete = context.web.get_folder_by_server_relative_url(folder_path)
    folder_to_delete.delete_object().execute_query()


# Define function for deleting file
def delete_file(file_path: str) -> None:
    """
    Delete file
    :param file_path: file path in string format
    :return: None
    """
    # Define file to be deleted and delete it
    file_to_delete = context.web.get_file_by_server_relative_url(file_path)
    file_to_delete.delete_object().execute_query()


# Define function that compose path
def compose_path(root_path: str, name: str) -> str:
    """
    Generate path
    :param root_path: path of the current directory in string format
    :param name: name in string format
    :return: paths_list: new path in string format
    """
    # Compose new path by concatenating root path and folder / file name
    child_path = root_path + "/" + name

    # Return new path
    return child_path


# Define function that decompose path
def decompose_path(root_path: str) -> str:
    """
    Generate path of parent folder
    :param root_path: path of the current directory in string format
    :return: parent_path: parent paths in string format
    """
    # Compose new path by deleting last part from root path
    raw_parent_path = root_path.split("/")[:-1]
    parent_path = "/".join(raw_parent_path)

    # Return new path
    return parent_path


def delete_directory(absolute_path: str):
    """
    Delete folder content and the folder itself
    :param absolute_path: path of the parent directory in string format
    :return: delete_directory: call itself
    """
    # Change global variable
    global is_content

    # Define folder level to limit delete process
    check_path = decompose_path(root_path)

    # Generate name of folders and files from current folder
    folder_names = get_folder_data(absolute_path)
    file_names = get_folder_files(absolute_path)

    # Delete files and folder if no subfolders
    if not folder_names:
        # Delete files
        if file_names:
            for file in file_names:
                file_path = compose_path(absolute_path, file)
                delete_file(file_path)
                print(f"Deleted file: {file}.")
        # Delete folder
        delete_folder(absolute_path)
        print(f"Deleted folder: {absolute_path.split('/')[:-1]}.")

        # Go one level up in folder structure
        relative_path = decompose_path(absolute_path)

        # Stop running if parent folder was deleted
        if relative_path == check_path:
            is_content = False
            print(f"Directory {root_path.split('/')[:-1]} was deleted.")

        # Return function recursively
        else:
            return delete_directory(relative_path)
    
    # Generate path for first subfolder
    else:
        folder = folder_names[0]
        relative_path = compose_path(absolute_path, folder)

        # Return function recursively
        return delete_directory(relative_path)
    

# Run code as a script
if __name__ == "__main__":

    # Define path of folder to be cleaned
    root_path = "/sites/<enterprise_site>/<parent_directory>/<...>"

    # Set a flag
    is_content = True

    # Run deleting process until parent directory is deleted
    while is_content:
        delete_directory(root_path)

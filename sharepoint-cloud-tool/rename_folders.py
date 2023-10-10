# Import necessary modules
from office365.sharepoint.client_context import ClientContext
from openpyxl import Workbook, load_workbook

# Import internal utilities
from credentials import username, password, site_url

# Create connection to the SharePoint
context = ClientContext(site_url).with_user_credentials(username, password)

# Define function that take all folders path
def get_folders_path(folder_path: str) -> list:
    """
    Return a list with the path of directories from provided URL
    :param folder_path: path of the directory in string format
    :return: folders_path: list with subdirectories path
    """
    # Get relative URL from SharePoint for given folder and load it
    library_root = context.web.get_folder_by_server_relative_path(folder_path)
    context.load(library_root)
    context.execute_query()

    # Retrieve all subfolders from provided folder
    folders = library_root.folders
    context.load(folders)
    context.execute_query()

    # Iterate through all subfolder and save the path
    folders_path = []
    for folder in folders:
        context.load(folder)
        context.execute_query()
        folders_path.append(folder.properties["ServerRelativeUrl"])

    # Return a list of subfolders path
    return folders_path


# Define function that take all subfolders path
def get_subfolders_path(folder_path: str, target_folder: str) -> list:
    """
    Return a list with the path of subdirectories from provided URL
    :param folder_path: path of the directory in string format
    :param target_folder: name of the directory in string format
    :return: folders_path: list with subdirectories path
    """
    # Get relative URL from SharePoint for given folder and load it
    library_root = context.web.get_folder_by_server_relative_path(folder_path)
    context.load(library_root)
    context.execute_query()

    # Retrieve all subfolders from provided folder
    folders = library_root.folders
    context.load(folders)
    context.execute_query()

    # Iterate through all subfolder and save the path
    folders_path = []
    for folder in folders:
        context.load(folder)
        context.execute_query()

        # If subfolder match the name then append it to the list
        if folder.properties["Name"] == target_folder:
            folders_path.append(folder.properties["ServerRelativeUrl"])

    # Return a list of subfolders path
    return folders_path


# Define function that rename subfolder
def rename_folder(folder_path: str, new_folder_name: str) -> None:
    """
    Rename the directory from provided URL with specified name
    :param folder_path: path of the directory in string format
    :param new_folder_name: new name of the directory in string format
    :return: None
    """
    # Get relative URL from SharePoint for given folder and load it
    library_root = context.web.get_folder_by_server_relative_path(folder_path)
    context.load(library_root)
    context.execute_query()

    # Retrieve all subfolders from provided folder
    folders = library_root.folders
    context.load(folders)
    context.execute_query()

    # Iterate through all subfolder and save the path
    folders_path = []
    for folder in folders:
        context.load(folder)
        context.execute_query()
        context.web.get_folder_by_server_relative_path(folder.properties["ServerRelativeUrl"]).rename(new_folder_name).execute_query()


# Define function that create subfolder
def create_folder(folder_path: str, folder_name: str) -> None:
    """
    Create a subdirectory with specific name in the directories from provided URL
    :param folder_path: path of the directory in string format
    :param folder_name: name of the subdirectory in string format
    :return: None
    """
    # Take the parent folder and create subfolder with specific name
    parent_folder = context.web.get_folder_by_server_relative_url(folder_path)
    parent_folder.folders.add(folder_name)
    context.execute_query()


if __name__ == "__main__":

    # Define variables
    root_location = "SharePoint/Site/"
    folders_name = ["Folder_1", "Folder_2", "Folder_3"]
    target_folder_name = "Folder_01"

    # Get all subdirectories
    stores_path = get_folders_path(root_location)
    print("Retrieve all store paths.")
    print(stores_path)
    
    # Get specific subdirectories path from each store
    target_folder_paths = []
    for store_path in stores_path:
        target_folder_path = get_subfolders_path(store_path, target_folder_name)
        target_folder_paths.append(target_folder_path)
        print("Retrieve all specific folder paths.")
        print(target_folder_path)
    
    # Rename specific folder
    for folder_path in target_folder_paths:
        rename_folder(folder_path, target_folder_name)
        print("Rename folder.")
    
    # Create new folders
    for store_path in stores_path:
        for folder_name in folders_name:
            create_folder(store_path, folder_name)
            print("Create folder.")
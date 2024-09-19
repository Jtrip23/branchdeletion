import os
import pandas as pd
from github import Github, BadCredentialsException, UnknownObjectException
import logging

logging.basicConfig(level=logging.INFO)

def disable_branch_protection(repo, branch):
    try:
        branch_obj = repo.get_branch(branch)
        # Removing branch protection if it exists
        branch_obj.remove_protection()
        logging.info(f"Branch protection for '{branch}' has been disabled.")
    except UnknownObjectException:
        logging.info(f"No branch protection found for '{branch}', nothing to disable.")
    except Exception as e:
        logging.error(f"Error disabling branch protection for '{branch}': {e}")

def delete_branch(repo, branch):
    try:
        # Delete the branch using GitHub API
        ref = f"heads/{branch}"
        repo.get_git_ref(ref).delete()  # Use the GitHub API to delete the branch
        logging.info(f"Branch '{branch}' has been deleted successfully.")
        return True
    except Exception as e:
        logging.error(f"Error deleting branch '{branch}': {e}")
        return False

def process_repo(username, token, repo_name):
    branches_to_check = ["feature_cloudhub", "feature_new_cloudhub"]  # Branches to check for conditions
    branches_status = {}

    try:
        g = Github(username, token)
        repo = g.get_repo(repo_name)

        # Check the existence of branches
        for branch in branches_to_check:
            logging.info(f"Checking if branch '{branch}' exists in '{repo_name}'.")
            try:
                repo.get_branch(branch)
                branches_status[branch] = True
                logging.info(f"Branch '{branch}' exists.")
            except UnknownObjectException:
                branches_status[branch] = False
                logging.info(f"Branch '{branch}' does not exist in '{repo_name}'.")

        # Logic to delete branch
        if branches_status.get("feature_cloudhub") and branches_status.get("feature_new_cloudhub"):
            logging.info("Both branches found. Deleting 'feature_cloudhub'.")
            disable_branch_protection(repo, "feature_cloudhub")
            delete_branch(repo, "feature_cloudhub")
        elif branches_status.get("feature_cloudhub"):
            logging.info("'feature_cloudhub' is present but will not be deleted as 'feature_new_cloudhub' is not found.")
        else:
            logging.info("No relevant branches to delete.")

    except BadCredentialsException:
        logging.error("Invalid GitHub credentials. Please check your username and token.")
    except UnknownObjectException:
        logging.error(f"Repository '{repo_name}' not found.")
    except Exception as e:
        logging.error(f"Error processing repository '{repo_name}': {e}")

def create_branches_from_excel(username, token, excel_file, output_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        results = []

        for index, row in df.iterrows():
            repo_name = row['source_repo_name'].strip()  # Ensure no leading/trailing spaces
            logging.info(f"Processing repository: {repo_name}")

            if '/' not in repo_name:
                logging.error(f"Invalid repository name format: '{repo_name}'. It should be 'username/repository_name'.")
                continue
            
            process_repo(username, token, repo_name)

        logging.info(f"Processing complete.")

    except Exception as e:
        logging.error(f"Error reading Excel file or processing branches: {e}")

if __name__ == "__main__":
    username = os.getenv('USERNAME')
    token = os.getenv('TOKEN')
    excel_file = 'repositories.xlsx'  # Path to your input Excel file
    output_file = 'branch_deletion_results.xlsx'  # Path to your output Excel file

    if not (username and token):
        logging.error("GitHub credentials not provided. Set USERNAME and TOKEN environment variables.")
    else:
        create_branches_from_excel(username, token, excel_file, output_file)

import os
import logging
import pandas as pd
from github import Github, BadCredentialsException, UnknownObjectException

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

def process_repo(username, token, repo_name, branches_to_delete):
    try:
        g = Github(username, token)
        
        # Ensure repo_name is in the correct format: username/repository_name
        if '/' not in repo_name:
            repo_name = f"{username}/{repo_name}"  # Prepend the username if not already included
        
        repo = g.get_repo(repo_name)

        results = []

        for branch in branches_to_delete:
            logging.info(f"Checking if branch '{branch}' exists in '{repo_name}'.")
            try:
                repo.get_branch(branch)
                disable_branch_protection(repo, branch)
                success = delete_branch(repo, branch)
                results.append({'branch': branch, 'status': 'deleted' if success else 'failed'})
            except UnknownObjectException:
                logging.info(f"Branch '{branch}' does not exist in '{repo_name}'. Skipping deletion.")
                results.append({'branch': branch, 'status': 'not found'})
            except Exception as e:
                logging.error(f"Error processing branch '{branch}': {e}")
                results.append({'branch': branch, 'status': f'failed - {str(e)}'})

        return results

    except BadCredentialsException:
        logging.error("Invalid GitHub credentials. Please check your username and token.")
        return [{'branch': branch, 'status': 'failed - invalid credentials'} for branch in branches_to_delete]
    except UnknownObjectException:
        logging.error(f"Repository '{repo_name}' not found.")
        return [{'branch': branch, 'status': f'failed - repo not found: {repo_name}'} for branch in branches_to_delete]
    except Exception as e:
        logging.error(f"Error processing repository '{repo_name}': {e}")
        return [{'branch': branch, 'status': f'failed - {str(e)}'} for branch in branches_to_delete]

def delete_branches_in_repositories(username, token, repos, branches_to_delete):
    all_results = []

    for repo_name in repos:
        logging.info(f"Processing repository: {repo_name}")
        repo_results = process_repo(username, token, repo_name, branches_to_delete)
        all_results.extend(repo_results)

    return all_results

def read_repositories_from_excel(excel_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        # Assuming the column that contains the repository names is called 'repo_name'
        repos = df['repo_name'].dropna().tolist()  # Drop NaN values and convert to list
        return repos
    except Exception as e:
        logging.error(f"Error reading repositories from Excel file: {e}")
        return []

if __name__ == "__main__":
    username = os.getenv('USERNAME')
    token = os.getenv('TOKEN')

    # Path to the Excel file containing repository names
    excel_file = 'repositories.xlsx'  # Replace with the path to your Excel file

    # Branches to delete (can be customized)
    branches_to_delete = ['feature_cloudhub', 'development']  # Replace with your list of branches

    if not (username and token):
        logging.error("GitHub credentials not provided. Set USERNAME and TOKEN environment variables.")
    else:
        # Read repositories from Excel file
        repos = read_repositories_from_excel(excel_file)

        if not repos:
            logging.error(f"No repositories found in the Excel file: {excel_file}")
        else:
            results = delete_branches_in_repositories(username, token, repos, branches_to_delete)

            # Output results
            for result in results:
                logging.info(f"Branch '{result['branch']}': {result['status']}")

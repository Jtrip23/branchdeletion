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

def process_repo(username, token, repo_name, branches_to_delete):
    try:
        g = Github(username, token)
        repo = g.get_repo(repo_name)

        results = []

        for branch in branches_to_delete:
            if branch == "feature_cloudhub":
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

            else:
                logging.info(f"Skipping branch '{branch}' (not 'feature_cloudhub').")

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

def create_branches_from_excel(username, token, excel_file, output_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')
        results = []

        for index, row in df.iterrows():
            repo_name = row['source_repo_name'].strip()  # Ensure no leading/trailing spaces
            logging.info(f"Processing repository: {repo_name}")
            branches_to_delete = [branch.strip() for branch in row['branches'].split(',')]
            
            if '/' not in repo_name:
                logging.error(f"Invalid repository name format: '{repo_name}'. It should be 'username/repository_name'.")
                continue
            
            repo_results = process_repo(username, token, repo_name, branches_to_delete)
            results.extend(repo_results)

        # Create a DataFrame for results and save to Excel
        results_df = pd.DataFrame(results)
        results_df.to_excel(output_file, index=False)
        logging.info(f"Results saved to '{output_file}'.")

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

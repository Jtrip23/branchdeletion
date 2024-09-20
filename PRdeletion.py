import os
import pandas as pd
from github import Github, BadCredentialsException, UnknownObjectException
import logging

logging.basicConfig(level=logging.INFO)

def close_pull_request(repo, pr_number):
    try:
        pr = repo.get_pull(pr_number)
        if pr.state == "open":
            pr.edit(state='closed')  # Close the PR if it is open
            logging.info(f"Closed pull request #{pr_number}: '{pr.title}'.")
        else:
            logging.info(f"Pull request #{pr_number}: '{pr.title}' is already closed.")
    except Exception as e:
        logging.error(f"Error closing pull request #{pr_number}: {e}")

def process_repo(username, token, repo_name):
    target_base_branch = "integration_cloudhub"
    
    try:
        g = Github(username, token)
        repo = g.get_repo(repo_name)

        # Fetch all pull requests in the repository
        prs = repo.get_pulls(state='open')  # Get only open PRs
        target_prs = [pr for pr in prs if pr.base.ref == target_base_branch]

        if target_prs:
            for pr in target_prs:
                logging.info(f"Closing pull request '{pr.title}' (#{pr.number}) with base branch '{target_base_branch}'.")
                close_pull_request(repo, pr.number)
        else:
            logging.info(f"No open pull requests with base branch '{target_base_branch}' found in '{repo_name}'.")

    except BadCredentialsException:
        logging.error("Invalid GitHub credentials. Please check your username and token.")
    except UnknownObjectException:
        logging.error(f"Repository '{repo_name}' not found.")
    except Exception as e:
        logging.error(f"Error processing repository '{repo_name}': {e}")

def create_branches_from_excel(username, token, excel_file):
    try:
        df = pd.read_excel(excel_file, engine='openpyxl')

        for index, row in df.iterrows():
            repo_name = row['source_repo_name'].strip()  # Ensure no leading/trailing spaces
            logging.info(f"Processing repository: {repo_name}")

            if '/' not in repo_name:
                logging.error(f"Invalid repository name format: '{repo_name}'. It should be 'username/repository_name'.")
                continue
            
            process_repo(username, token, repo_name)

        logging.info("Processing complete.")

    except Exception as e:
        logging.error(f"Error reading Excel file or processing repositories: {e}")

if __name__ == "__main__":
    username = os.getenv('USERNAME')
    token = os.getenv('TOKEN')
    excel_file = 'repositories.xlsx'  # Path to your input Excel file

    if not (username and token):
        logging.error("GitHub credentials not provided. Set USERNAME and TOKEN environment variables.")
    else:
        create_branches_from_excel(username, token, excel_file)

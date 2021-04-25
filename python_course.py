from typing import Optional, Dict
import exchangelib
import getpass
import ipyparams
import pathlib
import os

Email: str
# interp. An email address as a string
# Examples: 
EM_1 = "cferster@rjc.ca"

Name: str
# interp. A person's full name as registered with their email
# Examples: 
N1 = "Connor Ferster"

UserName: str
# interp. A person's email user name
# Examples:
UN1 = "cferster"


Members: Dict[Email, Name]
# interp: A dict with Email keys and Name values to represent active members of a group
# Examples
M1 = {EM_1: N1}


def get_hub_user() -> Optional[UserName]:
    """
    Returns the UserName of the current user if the function is run
    within a JupyterHub account. Returns None otherwise.
    """
    user = os.getenv('JUPYTERHUB_USER') # Returns None if not found
    return user


def submit_workbook(subject_line: str, recipient: Email):
    """
    Returns None. Assumes being run in a Jupyter environment.

    Emails the current Jupyter Notebook as an attachment through the user's 
    Exchange account in an email with 'subject_line' to 'recipient'.
    """
    user_name = get_hub_user()
    notebook_path = get_notebook_path()
    user_email = user_name.lower() + "rjc.ca"
    account = connect_to_rjc_exchange(user_email)
    message = exchangelib.Message(
        account=account,
        folder=account.sent,
        subject=subject_line,
        body='Workbook submission',
        to_recipients=[exchangelib.Mailbox(email_address=recipient)]
    )
    with open(notebook_path) as nb_file:
        nb_data = nb_file.read()
    attachment = exchangelib.FileAttachment(name=notebook_path.name, content=nb_data)
    message.attach(attachment)
    message.send_and_save()


def get_notebook_path():
    """
    Returns the absolute path of the current notebook .ipynb file
    """
    cwd = pathlib.Path.cwd()
    file_name = ipyparams.notebook_name
    return cwd / file_name


def connect_to_rjc_exchange(email: Email) -> exchangelib.Account:
    """
    Get Exchange account connection with RJC email server
    """
    server = "mail.rjc.ca"
    domain = "RJC"
    username = get_hub_user()
    return connect_to_exchange(server, domain, email, username)


def connect_to_exchange(server: str, domain: str, email: str, username: str) -> exchangelib.Account:
    """
    Get Exchange account connection with server
    """
    credentials = exchangelib.Credentials(username= f'{domain}\\{username}', 
                                          password=getpass.getpass("Exchange pass:"))
    config = exchangelib.Configuration(server=server, credentials=credentials)
    return exchangelib.Account(primary_smtp_address=email, autodiscover=False, 
                   config = config, access_type='delegate')
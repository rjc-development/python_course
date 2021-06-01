from __future__ import annotations
from typing import Optional, Dict
import exchangelib
import getpass
import ipyparams
import pathlib
import os

__version__ = "0.0.8"

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


def submit_workbook(
    subject_line: str,
    recipient: Email,
    username_backup: Optional[str] = None,
    cc_me: bool = False
    ):
    """
    Emails the current notebook to 'recipient' with 'subject_line'.
    If 'cc_me' is True, then the email will also "CC" the current user.

    Uses the RJC Exchange mail server. Passwords are handled securely with
    the getpass library in the Python standard library.

    The current user is known through a query of the user's Jupyter username.
    Since the username is created from the user's RJC email prefix, once the
    username is known, the user's email address is known. This makes things
    convenient.

    'subject_line': a str representing the subject line to be used in the email
    'recipient': the primary recipient of the email, typically the Python course
        administrator.
    'cc_me': a bool indicating whether the current user should be "CC'd" on the
        email.
    'username_backup': If provided, this is used as an override for the current
        user's username. To be used when testing functionality from an account
        that is not linked to an RJC email (e.g. a JupyterHub admin account).
        Not useful in typical operation.
    """
    user_name = get_hub_user()
    if username_backup or not user_name:
        user_name = username_backup
    # if not user_name:
    #     raise EnvironmentError("submit_workbook() must be run from within a Jupyter Hub environment.")
    notebook_path = get_notebook_path()
    user_email = user_name.lower() + "@rjc.ca"
    
    account = connect_to_rjc_exchange(user_email, user_name)
    cc_user = []
    if cc_me:
        cc_user = [user_email]
    message = exchangelib.Message(
        account=account,
        folder=account.sent,
        subject=subject_line,
        body=f'{notebook_path.name} submission',
        to_recipients=[recipient],
        cc_recipients=cc_user,
    )
    with open(notebook_path) as nb_file:
        nb_data = nb_file.read().encode("utf-8")
    attachment = exchangelib.FileAttachment(name=notebook_path.name, content=nb_data)
    message.attach(attachment)
    message.send_and_save()
    print(f"{notebook_path.name} successfully submitted to {recipient}")


def get_notebook_path():
    """
    Returns the absolute path of the current notebook .ipynb file
    """
    cwd = pathlib.Path.cwd()
    file_name = ipyparams.notebook_name
    file_name = file_name.replace("%20", " ")
    return cwd / file_name


def connect_to_rjc_exchange(email: Email, username: UserName) -> exchangelib.Account:
    """
    Get Exchange account connection with RJC email server
    """
    server = "mail.rjc.ca"
    domain = "RJC"
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

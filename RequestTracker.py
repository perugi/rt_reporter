# RequestTracker.py - class for handling the RT REST API. Supports loging in,
# replying to a ticket and logging out.

import requests
import sys
import re

from helpers import exit_script

class RequestTracker:

    def __init__(self, url):
        """API initialization. The url contains the base url of the request
        tracker (e.g. https://rt.cosylab.com/REST/1.0/"""
        
        if not url.endswith("/"):
            url = url = "/"
        self.url = url

        self.session = requests.session()

    def __request(self, api_link, data = None):
        """Make a REST API post or get operation to the specified api.
        The api_link is the end part of the url, containing the ticket and
        operation to be performed (e.g. ticket/100/comment) and is joined with 
        the base url. The data is passed as a dictionary."""

        url = self.url + api_link

        if data:
            response = self.session.post(url, data = data)
        else:
            response = self.session.get(url)

        if response.status_code != 200:
            print(f"Error! Status code received: {response.status_code}.")
            exit_script()

        if '401 Credentials required' in str(response.content):
            print("Credentials error! Please retry.")
            exit_script()
        else:
            return response.content

    def login(self, username, password):
        """Login with the supplied credentials."""

        self.__request('', {'user': username, 'pass': password})

    def logout(self):
        """Logout of the request tracker."""

        self.__request('logout')

    def reply(self, ticket, text = 't', time_worked = 0):
        """Make an RT reply in the specified ticket with the supplied text and 
        time worked."""

        content = f"Text: {text}\nAction: correspond\nTimeWorked: {time_worked}"
        # print(f'ticket/{str(ticket)}/comment', {'content': content})
        self.__request(f'ticket/{str(ticket)}/comment', {'content': content})
        
    def get_name(self, ticket):
        """Return the name of the ticket, specified by the ticket number.""" 

        info = str(self.__request(f'ticket/{str(ticket)}/show'))
        
        name_regex = re.compile(r'Subject: (.*)\\nStatus')
        mo = name_regex.search(info)
        return mo[1]
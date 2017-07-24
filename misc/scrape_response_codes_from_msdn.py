#!/usr/bin/env python3

"""Dirty little script that scrapes the ResponseCode Web page on MSDN.

"""

import re
import sys

import requests
from bs4 import BeautifulSoup


class ResponseCode:
    def __init__(self, response_code, description):
        assert response_code
        assert description
        self.response_code = response_code
        self.description = description

    def __str__(self):
        return "ResponseCode('{}', '{}')".format(self.response_code,
                                                 self.description)


def convert_camel_case(name):
    """Converts name from CamelCase to underscore_name.
    """
    string = re.sub("(.)([A-Z][a-z]+)", r"\1_\2", name)  # first cap
    return re.sub("([a-z0-9])([A-Z])", r"\1_\2", string).lower()  # all cap


def grab_page(url):
    response = requests.get(url)
    if response.status_code != 200:
        raise Exception("Getting '{}' failed: {}".format(url, response.status_code))
    html_doc = response.text
    return html_doc


def response_code_in_page(html_doc):
    """Scrape the page for all the response codes.

    """
    soup = BeautifulSoup(html_doc, "html.parser")

    def find_first_row():
        for table in soup.find_all("table"):
            cell = table.find(string="NoError")
            if cell:
                return cell.find_parent("tr")

    first_row = find_first_row()
    if not first_row:
        raise Exception("Couldn't find a cell with text 'NoError'")

    for table_row in first_row.find_next_siblings("tr"):
        code = table_row.find("td", attrs={"data-th": "Value"}).p.text
        description = table_row.find("td", attrs={"data-th": "Description"}).p.text
        yield ResponseCode(code, description)


def main():
    if len(sys.argv) != 2:
        print("usage: {} URL".format(sys.argv[0]))
        sys.exit(1)

    url = sys.argv[1]
    html_doc = grab_page(url)
    for response_code in response_code_in_page(html_doc):
        print(response_code)


if __name__ == "__main__":
    main()

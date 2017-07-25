#!/usr/bin/env python3

"""Script that scrapes the ResponseCode Web page on MSDN and generates C++ code.

"""

import re
import sys
from textwrap import wrap

import html2text
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

    def sanitize(comment):
        """Strip leading/trailing WS, doubles spaces inside string,
        line-break issues, and Replace * → \li

        """
        comment = re.sub("^\|", "", comment)
        comment = re.sub(" {2}\*", "\li", comment)
        comment = re.sub(" +", " ", comment)
        comment = re.sub("- ", "", comment)
        comment = comment.strip()
        return comment

    def extract_table_row(row, converter):
        val = row.find("td", attrs={"data-th": "Value"}).p.text
        desc_html = row.find("td", attrs={"data-th": "Description"})
        desc = converter.handle(str(desc_html))
        return val, sanitize(desc)

    first_row = find_first_row()
    if not first_row:
        raise Exception("Couldn't find a cell with text 'NoError'")

    h = html2text.HTML2Text(baseurl="", bodywidth=0)
    h.ignore_links = True

    code, description = extract_table_row(first_row, converter=h)
    yield ResponseCode(code, description)

    for table_row in first_row.find_next_siblings("tr"):
        code, description = extract_table_row(table_row, converter=h)
        yield ResponseCode(code, description)


def print_cxx_enum(items: [ResponseCode]):
    def print_line_with_continuation(line, column=80, indentation="        ",
                                     prefix="//! "):
        pref = indentation + prefix
        print('\n'.join(
            ['\n'.join(wrap(block,
                            width=column,
                            initial_indent=pref,
                            subsequent_indent=pref))
             for block in line.splitlines()]))

    indent = "    "
    print(indent + "enum class response_code")
    print(indent + "{")
    no_of_items = len(items)
    count = 0
    indent = indent + indent
    for item in items:
        count = count + 1
        print_line_with_continuation(item.description)
        if count == no_of_items:
            print(indent + convert_camel_case(item.response_code))
        else:
            print(indent + convert_camel_case(item.response_code) + ",")
            print("")
    indent = "    "
    print(indent + "};")


def print_cxx_str_to_enum_function(items):
    i = "    "
    print(i + "inline response_code str_to_response_code(const std::string& str)")
    print(i + "{")

    for item in items:
        print(i + i + "if (str == \"" + item.response_code + "\")")
        print(i + i + "{")
        print(i + i + i + "return response_code::" + convert_camel_case(item.response_code) + ";")
        print(i + i + "}")

    print(i + i + "throw exception(\"Unrecognized response code: \" + str);")
    print(i + "}")


def print_enum_to_str_function(items):
    i = "    "
    print(i + "inline std::string enum_to_str(response_code code)")
    print(i + "{")
    print(i + i + "switch (code)")
    print(i + i + "{")

    for item in items:
        print(i + i + "case response_code::" + convert_camel_case(item.response_code) + ":")
        print(i + i + i + "return \"" + item.response_code + "\";")

    print(i + i + "default:")
    print(i + i + i + "throw exception(\"Unrecognized response code\");")
    print(i + i + "}")
    print(i + "}")


def main():
    if len(sys.argv) != 2:
        print("usage: {} URL".format(sys.argv[0]))
        sys.exit(1)

    url = sys.argv[1]
    html_doc = grab_page(url)
    items = [response_code for response_code in response_code_in_page(html_doc)]
    print_cxx_enum(items)
    print_cxx_str_to_enum_function(items)
    print_enum_to_str_function(items)


if __name__ == "__main__":
    main()

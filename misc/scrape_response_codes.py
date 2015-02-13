#!/usr/bin/env python2

"""Dirty little script that extracts response codes from Appendix A of "Inside
Microsoft Exchange Server 2007 Web Services" published by Microsoft Press and
converts it to C++ code.

There are more than 275 response codes in Exchange Web Services. This script
extracts all of them (along with a little bit of documentation) from a PDF
document, Appendix A "Error Codes", (available from the companion Web page) and
converts it to C++ code.

TODO: we should also scrape the ResponseCode Web page on MSDN to include
additions made in 2013.

"""

import sys
import re
from cStringIO import StringIO
from textwrap import wrap
try:
    import pdfminer
except:
    print("You need to install python-pdfminer for this script")
    sys.exit(1)
from pdfminer.pdfinterp import (PDFResourceManager, PDFPageInterpreter)
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage


class Data(object):
    def __init__(self, response_code):
        assert response_code
        self.response_code = response_code
        self.message = ""
        self.applicable_methods = ""
        self.comments = ""

    def __str__(self):
        return \
"Response Code: '{}' Message: '{}' Applicable Methods: '{}' Comments: '{}'" \
                .format(self.response_code,
                        self.message,
                        self.applicable_methods,
                        self.comments)

    def sanitize(self):
        """Strip leading/trailing WS, doubles spaces inside string, and
        line-break issues.

        """
        self.response_code = self.response_code.strip()
        self.message = self.message.strip()
        self.message = re.sub(' +', ' ', self.message)
        self.message = re.sub('- ', '', self.message)
        self.applicable_methods = self.applicable_methods.strip()
        self.comments = self.comments.strip()
        self.comments = re.sub(' +', ' ', self.comments)
        self.comments = re.sub('- ', '', self.comments)


def convert_camel_case(name):
    """Converts name from CamelCase to camel_case.
    """
    string = re.sub('(.)([A-Z][a-z]+)', r'\1_\2', name)  # first cap
    return re.sub('([a-z0-9])([A-Z])', r'\1_\2', string).lower()  # all cap

def convert(filename, pages=None):
    """Convert given PDF file to text."""
    if not pages:
        pagenums = set()
    else:
        pagenums = set(pages)

    output = StringIO()
    manager = PDFResourceManager()
    converter = TextConverter(manager, output, laparams=LAParams())
    interpreter = PDFPageInterpreter(manager, converter)

    infile = file(filename, 'rb')
    for page in PDFPage.get_pages(infile, pagenums):
        interpreter.process_page(page)
    infile.close()
    converter.close()
    text = output.getvalue()
    output.close
    return text

def interpret(text):
    """Parse given text line by line and try to make sense of it.
    """
    response_code_expr = re.compile(r"^Error\w+$|^NoError$")
    content = None
    in_message_line = False
    in_methods_line = False
    in_comments_line = False

    def is_response_code_line(string):
        match = re.search(response_code_expr, string)
        return match

    count = 0
    next_line = None
    lines = text.split("\n")
    no_of_lines = len(lines)
    for line in lines:
        count = count + 1
        if count == no_of_lines:
            next_line = None
        else:
            next_line = lines[count]
        if is_response_code_line(line):
            content = Data(response_code=line)
            in_message_line = True
        elif in_message_line:
            if line.endswith(".") or \
                    line.endswith("various messages") or \
                    line == "N/A":
                content.message = content.message + line
                in_message_line = False
                in_methods_line = True
            else:
                content.message = content.message + line + " "
        elif in_methods_line:
            if not content.applicable_methods:
                content.applicable_methods = line
            else:
                content.applicable_methods = \
                        content.applicable_methods + line + " "
            in_methods_line = False
            in_comments_line = True
            # look-ahead to support multi-line "Applicable Methods"
            if next_line:
                if (len(next_line) < 31 or next_line.endswith(")")) \
                        and not next_line.endswith("."):
                    in_methods_line = True
                    in_comments_line = False
        elif in_comments_line:
            if not line:
                in_comments_line = False
                content.comment = content.comments.strip()
                yield content
            else:
                if line.startswith("-"):
                    # A bullet point
                    content.comments = content.comments + "\n" + "* " + line[1:]
                    if content.comments.endswith("."):
                        content.comments = content.comments[:-1]
                else:
                    content.comments = content.comments + line + " "

def print_cxx_enum(items):
    def print_line_with_continuation(line, column=80, indentation="        ",
                                     prefix="// "):
        pref = indentation + prefix
        print('\n'.join(
                [ '\n'.join(wrap(block,
                                 width=column,
                                 initial_indent=pref,
                                 subsequent_indent=pref)) \
                        for block in line.splitlines() ]))

    indent = "    "
    print(indent + "enum class response_code")
    print(indent + "{")
    no_of_items = len(items)
    count = 0
    indent = indent + indent
    for item in items:
        count = count + 1
        print_line_with_continuation(item.comments)
        if count == no_of_items:
            print(indent + convert_camel_case(item.response_code))
        else:
            print(indent + convert_camel_case(item.response_code) + ",")
            print("")
    indent = "    "
    print(indent + "};")

def print_cxx_str_to_enum_function(items):
    i = "    "
    last = items.pop()
    print(i + "inline response_code")
    print(i + "string_to_response_code_enum(const std::string& str)")
    print(i + "{")
    print(i + i + "static const std::unordered_map<std::string, response_code> m{")
    for item in items:
        print(i + i + i + "{ \"" + item.response_code + "\", response_code::" \
                + convert_camel_case(item.response_code) + " },")

    print(i + i + i + "{ \"" + last.response_code + "\", response_code::" \
            + convert_camel_case(last.response_code) + " }")
    print(i + i + "};")
    print(i + i + "auto it = m.find(str);")
    print(i + i + "if (it == m.end())")
    print(i + i + "{")
    print(i + i + i + "throw exchange_error(\"Unrecognized response code\");")
    print(i + i + "}")
    print(i + i + "return it->second();")
    print(i + "}")

def main():
    if len(sys.argv) != 2:
        print("usage: {} FILE".format(sys.argv[0]))
        sys.exit(2)
    filename = sys.argv[1]
    text = convert(filename)
    items = []
    for data in interpret(text):
        data.sanitize()
        items.append(data)
    print("namespace ews")
    print("{")
    print_cxx_enum(items)
    print("")
    print_cxx_str_to_enum_function(items)
    print("}")

if __name__ == "__main__":
    main()

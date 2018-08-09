#!/usr/bin/env python

""" email_header_parser.py: Email header fields parser with user defined extractions to tab-delimited CSV file.
    The program processes any email files in .eml (text) type format located in a directory, parses the email headers,
    and puts all the header data in a tab-delimited text file. From the body of the email, it also pulls out some
    specific data. The emails must be individual files and not in a container, like an mbox or pst file..
    Usage: python email_header_parser.py [filepath to emails] [csv filename] """

__author__ = "Marsupilami8"
__version__ = "1.0"
__date__ = "2018-08-08"

import sys, os, fnmatch, email, re, csv

email_header_fields = set()  # Global variable for all the header fields pulled, to include non-standard ones

def main():

    if len(sys.argv) != 3:
        sys.exit("Error: Please provide a directory path to where the emails are located and "
                 "a name for the .CSV file. ")
    elif not os.path.exists(sys.argv[1]):
        sys.exit("Argument provided is not a directory.")
    else:
        processed_data = load_data(sys.argv[1])
        print("Processed %d files" % len(processed_data))
        if write_data_to_file(processed_data, sys.argv[2]):
            print("CSV file successfully written!")

def load_data(path):
    """ Loads the data from eml files located in a defined folder. The eml file is converted to an email message
        object, then the email message object gets parsed for the email header fields. The fields are placed in a
        keyed dictionary that is appended to a list."""

    pattern = '*.eml'  # File extension type to process
    data = []  #

    with os.scandir(path) as list_of_Entries:
        for entry in list_of_Entries:
            if entry.is_file() and fnmatch.fnmatch(entry, pattern):
                print("Working on parsing file: " + entry.name)
                file = open(entry, "r", encoding="utf8", errors='ignore')
                temp_dict = parse_header(file)
                data.append(temp_dict)
    return data

def parse_header(file):
    """ Parses the email headers from a passed file. This does not process MIME structured emails or anything
        with attachments. The input is an individual email file in eml format. """

    message = email.message_from_file(file)
    email_headers = [x.lower() for x in message.keys()] # Make lo case because some of the same fields may be up case

    # Set of just the duplicate headers found
    dup_email_header_fields = set(x for x in email_headers if email_headers.count(x) > 1)

    # Set of all the headers found
    uniq_email_header_fields = set(email_headers)

    add_email_header_tracker(uniq_email_header_fields)

    header = {}

    # Keep track of where the source eml file came from
    header['source-email-file'] = os.path.basename(file.name.lower())

    # Get the values for all the identified headers, if there are headers with same name concat the values
    for x in uniq_email_header_fields:
        if x not in dup_email_header_fields:
            header[x] = message.get(x)
        else:
            header[x] = message.get_all(x)

    # Pull out the email message text
    header['body'] = message.get_payload()

    # Pull out from the email message text items of interest, such as username, IP addresses
    header['username'], header['password'], header['ip'] = pull_artifacts(header['body'])

    return header  # Return a dictionary with type ['Email Header Name']: content

def add_email_header_tracker(header_f):
    """ Adds to a set any new email header fields encountered.  This is used to iterate over when making a CSV file."""

    for x in header_f:
        if x not in email_header_fields:
            email_header_fields.add(x)
    return True

def pull_artifacts(body):
    """ Searches the body of the text email for user defined artifacts described in a regex expression.
        This takes a lot of massaging of the regex to get the right pulls from all the different types of data.
        Add new regex as necessary to capture various artifacts."""

    # Find instances of email addresses/usernames (defined by field name), in the email body
    m1 = re.search('(username|email) *:(.*)', str(body), re.IGNORECASE)
    if m1 is not None:
        user = m1.group(2).strip()  # Captures the group of interest
    else:
        user = "*"

    # Find instances of user passwords (defined by field name), in the email body
    m2 = re.search('(email password|password|pass|pd|ps) *:(.*)', str(body), re.IGNORECASE)
    if m2 is not None:
        pwd = m2.group(2).strip()   # Captures the group of interest
    else:
        pwd = "*"

    # Find instances of an IP address in the email body
    m3 = re.search(':*(\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3})', str(body), re.IGNORECASE)
    if m3 is not None:
        ip = m3.group(1).strip()    # Captures the group of interest
    else:
        ip = "*"

    return user, pwd, ip  # Return maps to header['username'], header['password'], header['ip']

def write_data_to_file(data_list, output_file):
    """ Writes the data to a tab delimited CSV file from the list of email header dictionaries. Certain user defined
        columns are added to the CSV header row to reflect some data elements that were pulled from the parsing.
        Returns boolean value if the operation is successful or not"""

    csv_columns = sorted(email_header_fields)
    csv_columns.extend(['body', 'username', 'password', 'ip', 'source-email-file' ])  # user-defined fields

    csv_file = output_file

    try:
        with open(csv_file, 'w', newline='', encoding="utf8") as somefile:
            writer = csv.DictWriter(somefile, fieldnames=csv_columns, dialect='excel-tab')
            writer.writeheader()
            for dict_list in data_list:
                writer.writerow(dict_list)
        return True
    except IOError:
        print("I/O Error!")
        return False

if __name__ == "__main__": main()

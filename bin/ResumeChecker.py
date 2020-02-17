#!/usr/bin/env python
# coding=utf-8
"""
coding=utf-8

A utility to make handling many resumes easier by automatically pulling contact information, required skills and
custom text fields. These results are then surfaced as a convenient summary CSV.

"""
import argparse
import csv
import json
import functools
import glob
import logging
import os
import re
import sys
reload(sys)
sys.setdefaultencoding('utf8')

import pandas as pd

from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfpage import PDFPage
from cStringIO import StringIO
from subprocess import Popen,PIPE
from docx import opendocx, getdocumenttext

logging.basicConfig(level=logging.DEBUG)

__author__ = 'bjherger'
__license__ = 'http://opensource.org/licenses/MIT'
__version__ = '2.0'
__email__ = '13herger@gmail.com'
__status__ = 'Development'
__maintainer__ = 'bjherger'



def main():
    """
    Main method for ResumeParser. This utility will:
     - Read in `data_path` and `output_path` from command line arguments
     - Create a list of documents to scan
     - Read the text from those documents
     - Pull out desired information (e.g. contact info, skills, custom text fields)
     - Output summary CSV

    :return: None
    :rtype: None
    """
    logging.info('Begin Main')

    # Parse command line arguments
    logging.info('Parsing input arguments')
    parser = argparse.ArgumentParser(
        description='Script to parse PDF resumes, and create a csv file containing contact info '
                    'and required fields')
    parser.add_argument('--data_path', help='Path to folder containing documents ending in .pdf',
                        required=True)
    parser.add_argument('--output_path', help='Path to place output .csv file',
                        default='../data/output/resumes_output.csv')

    args = parser.parse_args()

    logging.info('Command line arguments: %s', vars(args))

    # Create resume resume_df
    resume_df = create_resume_df(args.data_path)

    # Output to CSV
    resume_df.to_csv(args.output_path, quoting=csv.QUOTE_ALL, encoding='utf-8')

    logging.info('End Main')

def convert_pdf_to_txt(input_pdf_path):
    """
    A utility function to convert a machine-readable PDF to raw text.

    This code is largely borrowed from existing solutions, and does not match the style of the rest of this repo.
    :param input_pdf_path: Path to the .pdf file which should be converted
    :type input_pdf_path: str
    :return: The text contents of the pdf
    :rtype: str
    """
    try:
        logging.debug('Converting pdf to txt: ' + str(input_pdf_path))
        # Setup pdf reader
        rsrcmgr = PDFResourceManager()
        retstr = StringIO()
        codec = 'utf-8'
        laparams = LAParams()
        device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
        interpreter = PDFPageInterpreter(rsrcmgr, device)
        password = ""
        maxpages = 0
        caching = True
        pagenos = set()

        # Iterate through pages
        path_open = file(input_pdf_path, 'rb')
        for page in PDFPage.get_pages(path_open, pagenos, maxpages=maxpages, password=password,
                                      caching=caching, check_extractable=True):
            interpreter.process_page(page)
        path_open.close()
        device.close()

        # Get full string from PDF
        full_string = retstr.getvalue()
        retstr.close()

        # Normalize a bit, removing line breaks
        full_string = full_string.replace("\r", "\n")
        full_string = full_string.replace("\n", " ")

        # Remove awkward LaTeX bullet characters
        full_string = re.sub(r"\(cid:\d{0,2}\)", " ", full_string)
        return full_string.encode('ascii', errors='ignore')
        
    except Exception, exception_instance:
        logging.error('Error in file: ' + input_pdf_path + str(exception_instance))
        return ''

def convert_doc_to_text(file_path):
    filename = file_path
    try:
       logging.debug('Converting : ' + str(file_path))
       if filename[-4:] == ".doc":
            cmd = ['antiword', file_path]
            p = Popen(cmd,stdout=PIPE)
            stdout,stderr = p.communicate()
            return stdout.decode('ascii','ignore')
       elif filename[-5:] == ".docx":
            document = opendocx(file_path)
            paratextlist = getdocumenttext(document)
            newparatextlist = []
            for paratext in paratextlist:
                newparatextlist.append(paratext.encode("utf-8"))
            return '\n\n'.join(newparatextlist)
       elif filename[-4:] == ".odt":
            cmd = ['odt2txt',file_path]
            p = Popen(cmd,stdout=PIPE)
            stdout, stderr = p.communicate()
            return stdout.decode('ascii','ignore')
       elif filename[-4:] == ".pdf":
            return convert_pdf_to_txt(file_path)
    except Exception, exception_instance:
        logging.error('Error in file: ' + file_path + str(exception_instance))
        return ''

def convert_file_to_text(file_path):
    input_pdf_path = file_path
    try:
        if ".pdf" in input_pdf_path:
            logging.debug('Converting pdf to txt: ' + str(input_pdf_path))
            # Setup pdf reader
            rsrcmgr = PDFResourceManager()
            retstr = StringIO()
            codec = 'utf-8'
            laparams = LAParams()
            device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
            interpreter = PDFPageInterpreter(rsrcmgr, device)
            password = ""
            maxpages = 0
            caching = True
            pagenos = set()

            # Iterate through pages
            path_open = file(input_pdf_path, 'rb')
            for page in PDFPage.get_pages(path_open, pagenos, maxpages=maxpages, password=password,
                                      caching=caching, check_extractable=True):
                interpreter.process_page(page)
            path_open.close()
            device.close()

            # Get full string from PDF
            full_string = retstr.getvalue()
            retstr.close()

            # Normalize a bit, removing line breaks
            full_string = full_string.replace("\r", "\n")
            full_string = full_string.replace("\n", " ")

            # Remove awkward LaTeX bullet characters
            full_string = re.sub(r"\(cid:\d{0,2}\)", " ", full_string)
            return full_string.encode('ascii', errors='ignore')
        else:
            filename = file_path
            logging.debug('Converting : ' + str(file_path))
            if filename[-4:] == ".doc":
                cmd = ['antiword', file_path]
                p = Popen(cmd,stdout=PIPE)
                stdout,stderr = p.communicate()
                return stdout.decode('ascii','ignore')
            elif filename[-5:] == ".docx":
                document = opendocx(file_path)
                paratextlist = getdocumenttext(document)
                newparatextlist = []
                for paratext in paratextlist:
                    newparatextlist.append(paratext.encode("utf-8"))
                return '\n\n'.join(newparatextlist)
            elif filename[-4:] == ".odt":
                cmd = ['odt2txt',file_path]
                p = Popen(cmd,stdout=PIPE)
                stdout, stderr = p.communicate()
                return stdout.decode('ascii','ignore')
            elif filename[-4:] == ".pdf":
                return convert_pdf_to_txt(file_path)
    except Exception, exception_instance:
        logging.error('Error in file: ' + input_pdf_path + str(exception_instance))
        return ''

def create_resume_df(data_path):
    """

    This function creates a Pandas DF with one row for every input resume, and columns including the resumes's
    file path and raw text.

    This is achieved through the following steps:
     - Create a list of documents to scan
     - Read the text from those documents
     - Pull out desired information (e.g. contact info, skills, custom text fields)
    :param data_path: Path to a folder containing resumes. Any files ending in .pdf in this folder will be treated as a
    resume.
    :type data_path: str
    :return: A Pandas DF with one row for every input resume, and columns including the resumes's
    file path and raw text
    :rtype: pd.DataFrame
    """

    # Create a list of documents to scan
    logging.info('Searching path: ' + str(data_path))

    # Find all files in the data_path which end in `.pdf`. These will all be treated as resumes
    path_glob = os.path.join(data_path, '*.pdf')
    path_glob_doc = os.path.join(data_path, '*.docx')

    # Create list of files
    file_list = glob.glob(path_glob)
    file_list_doc = glob.glob(path_glob_doc)
    joined_list = file_list + file_list_doc
    logging.info('Iterating through file_list: ' + str(file_list))
    logging.info('Iterating through file_list_doc: ' + str(file_list_doc))
    logging.info('Iterating through file_list_join: ' + str(joined_list))
    resume_summary_df = pd.DataFrame()

    # Store metadata, raw text, and word count
    #resume_summary_df["file_path"] = file_list
    resume_summary_df["file_path"] = joined_list
    #resume_summary_df["raw_text"] = resume_summary_df["file_path"].apply(convert_pdf_to_txt)
    #resume_summary_df["raw_text"] = resume_summary_df["file_path"].apply(convert_doc_to_text)
    resume_summary_df["raw_text"] = resume_summary_df["file_path"].apply(convert_file_to_text)
    resume_summary_df["num_words"] = resume_summary_df["raw_text"].apply(lambda x: len(x.split()))


    # Scrape contact information
    resume_summary_df["phone_number"] = resume_summary_df["raw_text"].apply(check_phone_number)
    resume_summary_df["area_code"] = resume_summary_df["phone_number"].apply(functools.partial(term_match, term=r"\d{3}"))
    resume_summary_df["email"] = resume_summary_df["raw_text"].apply(check_email)
    resume_summary_df["email_domain"] = resume_summary_df["email"].apply(functools.partial(term_match, term=r"@(.+)"))
    resume_summary_df["address"] = resume_summary_df["raw_text"].apply(check_address)
    #resume_summary_df["linkedin"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"linkedin"))
    #resume_summary_df["github"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"github"))

    # Scrape education information
    resume_summary_df["phd"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"ph.?d.?"))

    # Scrape skill information
    f = open('TermList.txt', 'r')
    x = f.readlines()
    for line in x:
        vTerm2 = line.strip()
       # vTerm2Count = vTerm2 + "_count"
        vTerm2Count = vTerm2
        print vTerm2
        print vTerm2Count
      #  resume_summary_df[vTerm2Count] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=vTerm2))
        resume_summary_df[vTerm2Count] = resume_summary_df["raw_text"].apply(functools.partial(term_string, term=vTerm2))
      
    """
    resume_summary_df["java_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"java"))
    resume_summary_df["python_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"python"))
    resume_summary_df["R_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r" R[ ,]"))
    resume_summary_df["latex_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"latex"))
    resume_summary_df["stata_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"stata"))
    resume_summary_df["CS_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"computer science"))
    resume_summary_df["mysql_count"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"mysql"))
    resume_summary_df["ms_office"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"microsoft office"))
    resume_summary_df["analytics"] = resume_summary_df["raw_text"].apply(functools.partial(term_count, term=r"analytics"))
    """

    # Return enriched DF
    return resume_summary_df

def create_json_file():
    # Open the CSV  
    #f = open( 'C:/Users/alhaadi.a/Desktop/Resume_Parser_sample_tools/Phyton/FastDataAnalytics_ResumeParser/data/output/resumes_output.csv', 'rU' ) 
    f = open( '../data/output/resumes_output.csv', 'rU' )  
# Change each fieldname to the appropriate field name. I know, so difficult.  
    reader = csv.DictReader( f, fieldnames = ( "position","file_path","raw_text","num_words","phone_number","area_code","email","email_domain","address","education" ))  
# Parse the CSV into JSON      
    out = json.dumps( [ row for row in reader] )  
    print "JSON parsed!"  
# Save the JSON  
    #f = open( 'C:/Users/alhaadi.a/Desktop/Resume_Parser_sample_tools/Phyton/FastDataAnalytics_ResumeParser/data/output/resumes_parsed.json', 'w')  
    f = open( '../data/output/resumes_parsed.json', 'w' )  
    f.write(out.replace("null","skills")) 
    print "JSON saved!"  




def check_phone_number(string_to_search):
    """
    Find first phone number in the string_to_search
    :param string_to_search: A string to check for a phone number in
    :type string_to_search: str
    :return: A string containing the first phone number, or None if no phone number is found.
    :rtype: str
    """
    try:
        regular_expression = re.compile(r"\(?"  # open parenthesis
                                        r"(\d{3})?"  # area code
                                        r"\)?"  # close parenthesis
                                        r"[\s\.-]{0,2}?"  # area code, phone separator
                                        r"(\d{3})"  # 3 digit exchange
                                        r"[\s\.-]{0,2}"  # separator bbetween 3 digit exchange, 4 digit local
                                        r"(\d{4})",  # 4 digit local
                                        re.IGNORECASE)
        result = re.search(regular_expression, string_to_search)
        if result:
            result = result.groups()
            result = "-".join(result)
        return result
    except Exception, exception_instance:
        logging.error('Issue parsing phone number: ' + string_to_search + str(exception_instance))
        return None


def check_email(string_to_search):
    """
       Find first email address in the string_to_search
       :param string_to_search: A string to check for an email address in
       :type string_to_search: str
       :return: A string containing the first email address, or None if no email address is found.
       :rtype: str
       """
    try:
        regular_expression = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}", re.IGNORECASE)
        result = re.search(regular_expression, string_to_search)
        if result:
            result = result.group()
        return result
    except Exception, exception_instance:
        logging.error('Issue parsing email number: ' + string_to_search + str(exception_instance))
        return None


def check_address(string_to_search):
    """
       Find first physical address in the string_to_search
       :param string_to_search: A string to check for a physical address in
       :type string_to_search: str
       :return: A string containing the first address, or None if no physical address is found.
       :rtype: str
       """
    try:
        regular_expression = re.compile(r"[0-9]+ [a-z0-9,\.# ]+\bCA\b", re.IGNORECASE)
        result = re.search(regular_expression, string_to_search)
        if result:
            result = result.group()

        return result
    except Exception, exception_instance:
        logging.error('Issue parsing email number: ' + string_to_search + str(exception_instance))

        return None


def term_count(string_to_search, term):
    """
    A utility function which counts the number of times `term` occurs in `string_to_search`
    :param string_to_search: A string which may or may not contain the term.
    :type string_to_search: str
    :param term: The term to search for the number of occurrences for
    :type term: str
    :return: The number of times the `term` occurs in the `string_to_search`
    :rtype: int
    """
    try:
        regular_expression = re.compile(term, re.IGNORECASE)
        result = re.findall(regular_expression, string_to_search)
        return len(result)
    except Exception, exception_instance:
        logging.error('Issue parsing term: ' + str(term) + ' from string: ' + str(
            string_to_search) + ': ' + str(exception_instance))
        return 0


def term_string(string_to_search, term):
    """
    A utility function which counts the number of times `term` occurs in `string_to_search`
    :param string_to_search: A string which may or may not contain the term.
    :type string_to_search: str
    :param term: The term to search for the number of occurrences for
    :type term: str
    :return: The number of times the `term` occurs in the `string_to_search`
    :rtype: int
    """
    try:
        regular_expression = re.compile(term, re.IGNORECASE)
        result = re.findall(regular_expression, string_to_search)
        for item in result:
         return item
    except Exception, exception_instance:
        logging.error('Issue parsing term: ' + str(term) + ' from string: ' + str(
            string_to_search) + ': ' + str(exception_instance))
        return 0


def term_match(string_to_search, term):
    """
    A utility function which return the first match to the `regex_pattern` in the `string_to_search`
    :param string_to_search: A string which may or may not contain the term.
    :type string_to_search: str
    :param term: The term to search for the number of occurrences for
    :type term: str
    :return: The first match of the `regex_pattern` in the `string_to_search`
    :rtype: str
    """
    try:
        regular_expression = re.compile(term, re.IGNORECASE)
        result = re.findall(regular_expression, string_to_search)
        return result[0]
    except Exception, exception_instance:
        logging.error('Issue parsing term: ' + str(term) + ' from string: ' +
                      str(string_to_search) + ': ' + str(exception_instance))
        return None


if __name__ == '__main__':
    main()
    create_json_file()

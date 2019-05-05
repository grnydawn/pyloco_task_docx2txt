================
'docx2text' task
================

'docx2text' task extracts texts from a Microsoft Word file while keeping
its original structure. For example, it outputs tables of the Word file in
the form of a nested Python dictionary of {<table no.>: {<row no.>:
{<cell no.>: text, ...}, ... }, ...}. A following task may take advantage
of this structural information for selecting a right text in the output.

Installation
------------

Before installing 'docx2text' task, please make sure that 'pyloco' is installed.
Run the following command if you need to install 'pyloco'.

>>> pip install pyloco

Or, if 'pyloco' is already installed, upgrade 'pyloco' with the following command

>>> pip install -U pyloco

To install 'docx2text' task, run the following 'pyloco' command.

>>> pyloco install docx2text

Command-line syntax
-------------------

usage: pyloco docx2text [-h] [-t type] [--general-arguments] path 

extracts structured texts from Microsoft Word file

positional arguments:
  path                  input MS Word file

optional arguments:
  -h, --help            show this help message and exit
  -t type, --type type  docx object type to extract (default='table')
  --general-arguments   Task-common arguments. Use --verbose to see a list of
                        general arguments

forward output variables:
   data                 output items extracted from the input MS Word file


Example(s)
----------

# Assuming my.docx has tables, following command extracts texts from
# the tables of the docx file and print extracted texts on screen
>>> pyloco docx2text my.docx -- print

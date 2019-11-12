# docx-text

Converts docx file from Sharepoint, One drive as well as individual files to text format.


## Installation

    $pip install docx-text

## Dependencies

  - Python3
  
## Usage

    >>>import doctext
    >>>doc_text = doctext.DocFile(url=download_url_of_file)
    >>>text = doc_text.get_text()
    # or you may directly enter the path to docs file.
    >>>doc_text = doctext.DocFile(doc=path_to_docx_file)


Adapted from https://github.com/ankushshah89/python-docx2txt

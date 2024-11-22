# PyDoc

PyDoc is a Python lib for processing and translating MS Word files from CN to EN. It provides the abilities to typeset, pre/post-process and translate MS Word files.

## How to Intsall

Clone the repo and run the following command to install the dependencies:

pip install -r requirements.txt

## How to Use

### Prerequisites

0. The styles of MS Word files are controlled based on YOUR OWN MS Word Style Template!!!
1. ID and Password of volcengine machine translation is required since the Translation feature is based on volcengine's machine translation API. In the future there will be more translation solutions to come.
2. Set up your `ACCESS_KEY_ID` and `ACCESS_KEY_SECET` in the `.env` file in the root folder of PyDoc.

   Example:

   `ACCESS_KEY_ID={your_key}`

   `ACCESS_KEY_SECRET={your_key}`
3. After registering on volcengine (https://console.volcengine.com/translate/usage), create and select (if there are multiple Glossaries) your own Glossary (optional).
4. usage: pydoc.py [-h] -i INPUT -o OUTPUT [-p] [-t]

   Process and translate Word documents.

   optional arguments:
   -h, --help            show this help message and exit
   -i INPUT, --input INPUT
   Input Word document path.
   -o OUTPUT, --output OUTPUT
   Output Word document path.
   -p, --preprocess   Preprocess the document before translation, remove the pre-body content in a MS Word file (*.docx)
   -t, --translate       Translate the document and return a Docx mixed with Chinese and English.

   Examples:

   Preprocess a doc without translating it: `python pydoc.py -p -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Translate a doc without preprocessing it: `python pydoc.py -t -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Preprocess and translate a doc:

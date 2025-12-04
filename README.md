# PyDoc

PyDoc is a Python lib for processing and translating MS Word files from CN to EN. It provides the abilities to typeset, pre/post-process and translate MS Word files.

## How to Intsall

Clone the repo and run the following command to install the dependencies:

pip install -r requirements.txt

## How to Use

### Prerequisites

0. The styles of MS Word files are controlled based on YOUR OWN MS Word Style Template!!!
1. ID and Password of volcengine machine translation is required since the Translation feature is based on volcengine's machine translation API.
2. Set up your `ACCESS_KEY_ID` and `ACCESS_KEY_SECET` in the `.env` file in the root folder of PyDoc.

   Example:

   `ACCESS_KEY_ID={your_key}`

   `ACCESS_KEY_SECRET={your_key}`
3. After registering on volcengine (https://console.volcengine.com/translate/usage), create and select (if there are multiple Glossaries) your own Glossary (optional).

4. For DeepL translation feature:
   - DeepL API authentication key is required. You can get it from https://www.deepl.com/pro-api
   - Set up your `DEEPL_AUTH_KEY` in the `.env` file in the root folder of PyDoc
   
   Example:
   
   `DEEPL_AUTH_KEY={your_key}`
4. usage: pydoc.py [-h] -i INPUT -o OUTPUT [-p] [-t] [-f] [--postprocess] [--deepl] [--deepl-source DEEPL_SOURCE] [--deepl-target DEEPL_TARGET] [--deepl-glossary DEEPL_GLOSSARY] [--deepl-key DEEPL_KEY]

   Process and translate Word documents.

   optional arguments:
   -h, --help            show this help message and exit
   -i INPUT, --input INPUT
                         Input Word document path.
   -o OUTPUT, --output OUTPUT
                         Output Word document path.
   -p, --preprocess      Preprocess the document before translation, remove the pre-body content in a MS Word file (*.docx)
   -t, --translate       Translate the document using VolcEngine API and return a Docx mixed with Chinese and English.
   -f, --check           Check the document structure and font consistency.
   --postprocess         Postprocess the document to add additional content.
   --deepl               Use DeepL API for translation.
   --deepl-source DEEPL_SOURCE
                         DeepL source language code (optional, auto-detected if not specified).
   --deepl-target DEEPL_TARGET
                         DeepL target language code (default: EN-US).
   --deepl-glossary DEEPL_GLOSSARY
                         Path to DeepL glossary JSON file.
   --deepl-key DEEPL_KEY
                         DeepL API authentication key (optional, can be set via DEEPL_AUTH_KEY environment variable).

   Examples:

   Preprocess a doc without translating it: `python pydoc.py -p -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Translate a doc without preprocessing it using VolcEngine: `python pydoc.py -t -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Translate a doc using DeepL without preprocessing: `python pydoc.py --deepl -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Translate a doc using DeepL with custom target language: `python pydoc.py --deepl --deepl-target FR -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Translate a doc using DeepL with glossary: `python pydoc.py --deepl --deepl-glossary "{path_to_glossary_json}" -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   Preprocess and translate a doc using DeepL: `python pydoc.py -p --deepl -i "{absolute_path_to_input_file}" -o "{absolute_path_to_output_file}"`

   DeepL Glossary JSON Format Example:
   ```json
   [
     {"source": "电子价签", "target": "Electronic Shelf Label"},
     {"source": "ESL", "target": "Electronic Shelf Label"},
     {"acronym": "AP", "expanded_form": "Wireless Access Point"}
   ]
   ```

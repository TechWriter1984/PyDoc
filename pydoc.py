import argparse
import os
from Preprocessor import Preprocessor
from Translator import Translator
from FileChecker import FileChecker
from Postprocessor import Postprocessor

glossary = [
    {"acronym": "All-Star", "expanded_form": "All-Star IoT Application Platform", "description": "Hanshow's new generation smart ESL and application management platform."},
    {"acronym": "ESL", "expanded_form": "Electronic Shelf Label", "description": "ESL is used for displaying product information such as promotions, price, and specifications."},
    {"acronym": "AP", "expanded_form": "Wireless Access Point", "description": "Access point for data transmission between ESL-Working and ESL."},
]

def process_document(input_file_path, output_file_path, preprocess, translate, check, postprocess):
    input_file_path = os.path.abspath(os.path.normpath(input_file_path))
    
    # Only normalize output_file_path if it is provided
    if output_file_path:
        output_file_path = os.path.abspath(os.path.normpath(output_file_path))

    if preprocess:
        print("Starting document preprocessing...")
        preprocessor = Preprocessor()
        preprocessor.process_word_file(input_file_path, output_file_path)
        print("Preprocessing completed.")

    if translate:
        print("Start document translation...")
        translator = Translator()
        translator.translate_word_file(input_file_path if not preprocess else output_file_path, output_file_path)
        print("Translation completed.")

    if check:
        print("Starting document check...")
        checker = FileChecker(output_file_path if (preprocess or translate) and output_file_path else input_file_path)
        checker.check_document_structure()
        checker.check_font_consistency()
        print("Document check completed.")

    if postprocess:
        print("Starting document postprocessing...")
        postprocessor = Postprocessor(glossary)
        postprocessor.process_word_file(output_file_path if (preprocess or translate) and output_file_path else input_file_path, output_file_path)
        print("Postprocessing completed.")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process and translate Word documents.')
    parser.add_argument('-i', '--input', required=True, type=str, help='Input Word document path.')
    parser.add_argument('-o', '--output', type=str, help='Output Word document path.')
    parser.add_argument('-p', '--preprocess', action='store_true', help='Preprocess the document before translation.')
    parser.add_argument('-t', '--translate', action='store_true', help='Translate the document.')
    parser.add_argument('-f', '--check', action='store_true', help='Check the document structure and font consistency.')
    parser.add_argument('--postprocess', action='store_true', help='Postprocess the document to add additional content.')

    args = parser.parse_args()

    input_file_path = args.input
    output_file_path = args.output

    # Check if at least one operation is specified
    if not any([args.preprocess, args.translate, args.check, args.postprocess]):
        print("Please specify at least one operation: preprocess (-p), translate (-t), check (-f), or postprocess (--postprocess).")
    
    # Check if input file exists
    elif not os.path.exists(input_file_path):
        print("The input file does not exist, please check the path.")
    
    # Check if -f flag is used without -o, allowing skipping output generation
    elif args.check and not args.preprocess and not args.translate and not args.postprocess and not output_file_path:
        print("Running check without output file generation...")
        process_document(input_file_path, None, args.preprocess, args.translate, args.check, args.postprocess)
    
    # Run with output file generation if -o is provided or other flags require it
    else:
        if not output_file_path:
            parser.error("Output file path (-o) is required if using preprocess (-p), translate (-t), or postprocess (--postprocess).")
        process_document(input_file_path, output_file_path, args.preprocess, args.translate, args.check, args.postprocess)

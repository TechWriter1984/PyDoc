import argparse
import os
from Preprocessor import Preprocessor
from Translator import Translator


def process_document(input_file_path, output_file_path, preprocess, translate):
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



if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Process and translate Word documents.')
    parser.add_argument('-i', '--input', required=True, type=str, help='Input Word document path.')
    parser.add_argument('-o', '--output', required=True, type=str, help='Output Word document path.')
    parser.add_argument('-p', '--preprocess', action='store_true', help='Preprocess the document before translation.')
    parser.add_argument('-t', '--translate', action='store_true', help='Translate the document.')

    args = parser.parse_args()

    input_file_path = args.input
    output_file_path = args.output

    # Check if at least one of preprocess or translate flags is set
    if not args.preprocess and not args.translate:
        print("Please specify at least one operation: preprocess (-p) or translate (-t)")
    elif os.path.exists(input_file_path):
        process_document(input_file_path, output_file_path, args.preprocess, args.translate)
    else:
        print("The input file does not exist, please check the path.")

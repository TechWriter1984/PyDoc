import argparse
import os
from Preprocessor import Preprocessor
from Translator import Translator
from FileChecker import FileChecker
from Postprocessor import Postprocessor
from deepl_translator import DeepLTranslator
from dotenv import load_dotenv
from doc_tester import check_document_parts as check_doc_parts

# 加载环境变量
load_dotenv()

glossary = [
    {"acronym": "All-Star", "expanded_form": "All-Star IoT Application Platform", "description": "Hanshow's new generation smart ESL and application management platform."},
    {"acronym": "ESL", "expanded_form": "Electronic Shelf Label", "description": "ESL is used for displaying product information such as promotions, price, and specifications."},
    {"acronym": "AP", "expanded_form": "Wireless Access Point", "description": "Access point for data transmission between ESL-Working and ESL."},
]

def process_document(input_file_path, output_file_path, preprocess, translate, check, postprocess, check_parts=False, deepl_translate=False, 
                    deepl_source_lang=None, deepl_target_lang='EN-US', deepl_glossary=None, deepl_auth_key=None, deepl_reuse_glossary=True):
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
    
    if deepl_translate:
        print("Starting DeepL document translation...")
        # 使用DeepL翻译器
        deepl_translator = DeepLTranslator(deepl_auth_key)
        deepl_translator.translate_file(
            input_path=input_file_path if not (preprocess or translate) else output_file_path,
            output_path=output_file_path,
            source_lang=deepl_source_lang,
            target_lang=deepl_target_lang,
            glossary_path=deepl_glossary,
            reuse_glossary=deepl_reuse_glossary
        )

    if check:
        print("Starting document check...")
        checker = FileChecker(output_file_path if (preprocess or translate) and output_file_path else input_file_path)
        checker.check_document_structure()
        checker.check_font_consistency()
        print("Document check completed.")
    
    if check_parts:
        print("\nStarting document parts integrity check...")
        try:
            result = check_doc_parts(output_file_path if (preprocess or translate) and output_file_path else input_file_path)
            print(f"Document parts check {'completed successfully' if result['status'] == 'success' else 'completed with warnings'}.")
        except Exception as e:
            print(f"Error during document parts check: {str(e)}")
        print("\n")

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
    parser.add_argument('--check-parts', action='store_true', help='Check the document parts integrity (cover, statement, TOC, etc.).')
    parser.add_argument('--postprocess', action='store_true', help='Postprocess the document to add additional content.')
    
    # DeepL翻译相关参数
    parser.add_argument('--deepl', action='store_true', help='Use DeepL API for translation.')
    parser.add_argument('--deepl-source', type=str, help='DeepL source language code (optional, auto-detected if not specified).')
    parser.add_argument('--deepl-target', type=str, default='EN-US', help='DeepL target language code (default: EN-US).')
    parser.add_argument('--deepl-glossary', type=str, help='Path to DeepL glossary JSON file.')
    parser.add_argument('--deepl-key', type=str, help='DeepL API authentication key (optional, can be set via DEEPL_AUTH_KEY environment variable).')
    parser.add_argument('--deepl-reuse-glossary', action='store_true', default=True, help='Reuse existing glossary with the same name, default: True')
    parser.add_argument('--deepl-no-reuse', action='store_false', dest='deepl_reuse_glossary', help='Create new glossary each time, do not reuse existing glossary')
    parser.add_argument('--deepl-list-glossaries', action='store_true', help='List all available DeepL glossaries')
    parser.add_argument('--deepl-cleanup', action='store_true', help='Delete all DeepL glossaries (use with caution)')

    args = parser.parse_args()

    input_file_path = args.input
    output_file_path = args.output

    # Handle special DeepL operations
    if args.deepl_list_glossaries:
        import sys
        from deepl_translator import DeepLTranslator
        from dotenv import load_dotenv
        
        load_dotenv()
        deepl_auth_key = args.deepl_key or os.getenv('DEEPL_AUTH_KEY')
        
        if not deepl_auth_key:
            print("DeepL API authentication key not provided, please set via DEEPL_AUTH_KEY environment variable or --deepl-key parameter")
            sys.exit(1)
            
        translator = DeepLTranslator(deepl_auth_key)
        translator.list_glossaries()
        sys.exit(0)
    
    if args.deepl_cleanup:
        import sys
        from deepl_translator import DeepLTranslator
        from dotenv import load_dotenv
        
        load_dotenv()
        deepl_auth_key = args.deepl_key or os.getenv('DEEPL_AUTH_KEY')
        
        if not deepl_auth_key:
            print("DeepL API authentication key not provided, please set via DEEPL_AUTH_KEY environment variable or --deepl-key parameter")
            sys.exit(1)
            
        translator = DeepLTranslator(deepl_auth_key)
        confirm = input("Are you sure you want to delete all glossaries? This cannot be undone! (y/n): ")
        if confirm.lower() == 'y':
            translator.delete_all_glossaries()
        else:
            print("Operation cancelled")
        sys.exit(0)
            
    # Check if at least one operation is specified
    if not any([args.preprocess, args.translate, args.check, args.check_parts, args.postprocess, args.deepl]) and not args.deepl_list_glossaries and not args.deepl_cleanup:
        print("Please specify at least one operation: preprocess (-p), translate (-t), check (-f), postprocess (--postprocess), or deepl (--deepl).")
    
    # Check if input file exists
    elif not os.path.exists(input_file_path):
        print("The input file does not exist, please check the path.")
    
    # Check if -f or --check-parts flag is used without -o, allowing skipping output generation
    elif (args.check or args.check_parts) and not args.preprocess and not args.translate and not args.postprocess and not output_file_path:
        print("Running check without output file generation...")
        process_document(input_file_path, None, args.preprocess, args.translate, args.check, args.postprocess, check_parts=args.check_parts)
    
    # Run with output file generation if -o is provided or other flags require it
    else:
        if not output_file_path:
            parser.error("Output file path (-o) is required if using preprocess (-p), translate (-t), or postprocess (--postprocess).")
        process_document(input_file_path, output_file_path, args.preprocess, args.translate, args.check, args.postprocess,
                        deepl_translate=args.deepl,
                        deepl_source_lang=args.deepl_source,
                        deepl_target_lang=args.deepl_target,
                        deepl_glossary=args.deepl_glossary,
                        deepl_auth_key=args.deepl_key,
                        deepl_reuse_glossary=args.deepl_reuse_glossary,
                        check_parts=args.check_parts)

import json

def generate_geolocation_file(cn_file, en_file, output_file):
    # Load data from the JSON files
    with open(cn_file, 'r', encoding='utf-8') as cn_f, open(en_file, 'r', encoding='utf-8') as en_f:
        cn_data = json.load(cn_f)
        en_data = json.load(en_f)

    # Open the output file
    with open(output_file, 'w', encoding='utf-8') as out_f:
        # Write pairs in "Chinese Value = English Value" format
        for cn_key, cn_value in cn_data.items():
            en_value = en_data.get(cn_key)
            if en_value:  # Only write if there is a matching English translation
                out_f.write(f"{cn_value}={en_value}\n")

    print(f"Geolocation data has been written to {output_file}")

# Define file paths
cn_file = 'D:\doc_marco_dev\data\input\cn.json'
en_file = 'D:\doc_marco_dev\data\input\en.json'
output_file = 'D:\doc_marco_dev\data\output\geolocation.txt'

# Generate the geolocation file
generate_geolocation_file(cn_file, en_file, output_file)

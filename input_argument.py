import argparse

# Create the parser
parser = argparse.ArgumentParser(description='A simple script with command-line arguments.')

# Add arguments using add_argument method
parser.add_argument('input_file', help='Path to the input file')
parser.add_argument('--output', '-o', help='Path to the output file (optional)')

# Parse the command-line arguments
args = parser.parse_args()

# Access the values of the arguments
input_file_path = args.input_file
output_file_path = args.output

# Now you can use input_file_path and output_file_path in your script
print(f'Input file: {input_file_path}')
print(f'Output file: {output_file_path}')

#!/usr/bin/env python3
import pandas as pd
import argparse
import os
import glob

# Set up command-line argument parsing
parser = argparse.ArgumentParser(
    description="Generate SQL comments for columns from Excel files. THe name of the excel file should be {tableName}.xlsx. This will write the files with liquibase notation, this will not effect direct execution but will facilitate future devops incorporation",
)
parser.add_argument('--author', '-a', type=str, default="wdubberley(generated)",
                    help="The name of the author for the changeset. Defaults to 'wdubberley(generated)' if not supplied.")
parser.add_argument('--filename', '-f', type=str,
                    help="Specific Excel file to process. If not supplied, processes all .xlsx files in the input directory.")
parser.add_argument('--directory', '-d', type=str, default="input",
                    help="Directory containing Excel files. Defaults to 'input' in the current directory if not supplied.")
parser.add_argument('--output_directory', '-o', type=str, default="output",
                    help="Directory to save the generated SQL files. Defaults to 'output' in the current directory if not supplied.")
parser.add_argument('--schema', '-s', type=str, required=True,
                    help="Target Schema, this is a required filed so that schema name is added to the table for sql execution")

args = parser.parse_args()

# Determine the files to process
if args.directory:
    # If a directory is provided, loop through all Excel files in that directory
    file_pattern = os.path.join(args.directory, "*.xlsx")
    files_to_process = glob.glob(file_pattern)
elif args.filename:
    # If a filename is provided, use that specific file
    files_to_process = [args.filename]
else:
    # If neither is provided, process all Excel files in the current directory
    files_to_process = glob.glob("*.xlsx")

# Check if no Excel files are found
if not files_to_process:
    print("No Excel files found to process.")
    exit()

# Determine the output directory
output_directory = args.output_directory if args.output_directory else os.getcwd()

# Create the output directory if it doesn't exist
os.makedirs(output_directory, exist_ok=True)

# Iterate over each Excel file found
for file in files_to_process:
    # Extract the table name from the filename (remove .xlsx extension)
    table_name = os.path.splitext(os.path.basename(file))[0]

    # Determine the output file path
    output_file_path = os.path.join(output_directory, f"{table_name}.sql")

    # Open the output .sql file
    with open(output_file_path, "w") as sql_file:
        # Load the Excel file into a DataFrame
        df = pd.read_excel(file)
        df = df.dropna(subset=['Description'])
        sql_file.write("-- liquibase formatted sql\n\n")
        # Iterate over the DataFrame and generate the SQL comments
        for index, row in df.iterrows():

            changeset_line = f"-- changeset {args.author}:{index+1}"
            comment_line = f"comment on {args.schema}.{table_name}.{row['Column']} is '{row['Description']}';"

            # Print the lines to the console
            print(changeset_line)
            print(comment_line)
            print()  # For the blank line between loops

            # Write the lines to the .sql file with a blank line in between
            sql_file.write(changeset_line + "\n")
            sql_file.write(comment_line + "\n")
            sql_file.write("\n")  # Blank line between each loop iteration

    print(f"SQL file generated: {output_file_path}")

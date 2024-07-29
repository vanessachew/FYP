# process generate summaries for more readability

# Open the original text file for reading
with open('top10.txt', 'r') as file:
    lines = file.readlines()

# Process each line to remove square brackets, quote marks, and commas
processed_lines = []
for line in lines:
    # Process the text to remove square brackets, quote marks, and commas
    line = line.replace('[', '').replace(']', '').replace("'", "").replace(',', '')
    # Remove excess white space between text chunks
    line = ' '.join(line.split())

    processed_lines.append(line)


# Write the processed lines to a new text file
with open('processed_text.txt', 'w') as file:
    for line in processed_lines:
        file.write(line + '\n')

import subprocess
import os
import re
import pypandoc

# Define the directory where the files are located
directory = "C:/Users/Connor/Desktop/docx-cleaner/INPUT"

# Change to the specified directory
os.chdir(directory)

# Define the pandoc command
command = ["pandoc", "-s", "ptesting.docx", "-t", "markdown", "-o", "pandocTesting.md"]

# Run the command
result = subprocess.run(command, capture_output=True, text=True)

# Print the output and errors (if any)
print("Output:", result.stdout)
print("Errors:", result.stderr)

##########################################################################
##########################################################################

def replace_latex_delimiters(text):
    # Replace \[ ... \] with $$ ... $$
    text = re.sub(r'\\\[(.*?)\\\]', r'$$\1$$', text, flags=re.DOTALL)
    # Replace \( ... \) with $ ... $
    text = re.sub(r'\\\((.*?)\\\)', r'$\1$', text, flags=re.DOTALL)
    return text


def replace_double_curly_braces(text):
    # Replace {{ with { { and }} with } }
    text = text.replace('{{', '{ {').replace('}}', '} }')
    return text


def replace_double_backslashes(text):
    # Replace \\ with \
    text = text.replace('\\\\', '\\')
    return text


def remove_unnecessary_braces(latex):
    # Regex pattern to match unnecessary curly braces around single variables
    pattern = r'\{([a-zA-Z0-9_]+)\}'

    # Substitute the matched pattern with the captured group (i.e., removing the braces)
    cleaned_latex = re.sub(pattern, r'\1', latex)

    return cleaned_latex


def read_md(file_path):
    # Read the Markdown file
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()

    return text


def process_latex_in_text(text):
    # Regex pattern to find LaTeX equations in text
    latex_pattern = r'\$\$.*?\$\$'

    def replace_latex(match):
        # Remove unnecessary braces in the matched LaTeX equation
        return remove_unnecessary_braces(match.group(0))

    # Replace LaTeX equations in the text
    processed_text = re.sub(latex_pattern, replace_latex, text, flags=re.DOTALL)

    return processed_text


def save_markdown(file_path, content):
    with open(file_path, 'w', encoding='utf-8') as f:
        f.write(content)


def main(input_md, output_md):
    # Read the .md file
    md_text = read_md(input_md)

    # Replace \[ ... \] with $$ ... $$ and \( ... \) with $ ... $
    md_text = replace_latex_delimiters(md_text)

    # Replace {{ and }} with { { and } }
    md_text = replace_double_curly_braces(md_text)

    # Replace \\ with \
    md_text = replace_double_backslashes(md_text)

    # Process LaTeX equations in the text
    processed_text = process_latex_in_text(md_text)

    # Ensure double newlines after headings and paragraphs for proper formatting
    markdown_text = re.sub(r'(\n\n+)', r'\n\n', processed_text)

    # Save the processed Markdown content to a file
    save_markdown(output_md, markdown_text)

if __name__ == '__main__':
    input_md = 'C:/Users/Connor/Desktop/docx-cleaner/INPUT/pandocTesting.md'  # Path to your input .md file
    output_md = "C:/Users/Connor/Desktop/docx-cleaner/OUTPUT/output.md"  # Path to your output .md file

    main(input_md, output_md)

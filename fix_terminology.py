import os
import re

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    new_content = content

    # Replace 'Step 7' with 'STEP 7'
    new_content = re.sub(r'\bStep 7\b', 'STEP 7', new_content)

    # Replace 'Step7' with 'STEP 7' where it's not part of a DLL or namespace or URL
    # Negative lookbehind/lookahead for . or / or no-space before/after if it's code
    new_content = re.sub(r'(?<!\.|/|-)\bStep7\b(?!\.|/|-)', 'STEP 7', new_content)

    # Replace 'DriveCliq' with 'DRIVE-CLiQ' where it's not part of a variable or file name or URL
    # e.g., avoid `DriveCliq` in code block if possible, but mostly it's text.
    # We will just replace DriveCliq with DRIVE-CLiQ in normal text.
    # Let's do a safe replacement
    new_content = re.sub(r'(?<!-)\bDriveCliq\b(?!-)', 'DRIVE-CLiQ', new_content)

    # Revert DRIVE-CLiQ in URLs and filenames
    new_content = new_content.replace('networks-and-DRIVE-CLiQ', 'networks-and-drivecliq')
    new_content = new_content.replace('DRIVE-CLiQ.dll', 'DriveCliq.dll') # if any

    # Fix variables that got changed if any
    new_content = new_content.replace('DRIVE-CLiQInterface', 'driveCliqInterface')

    # Check for Profinet -> PROFINET
    new_content = re.sub(r'\bProfinet\b', 'PROFINET', new_content)

    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Updated {filepath}")

for root, _, files in os.walk('tia-portal-ai-extensions-main'):
    for file in files:
        if file.endswith('.md'):
            process_file(os.path.join(root, file))

import os
import re

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    new_content = content

    # Check for Profinet -> PROFINET in any casing where it's not already PROFINET
    # Case insensitive substitution except PROFINET. Let's do re.sub(r'(?i)\bprofinet\b', 'PROFINET', new_content)
    new_content = re.sub(r'(?i)\bprofinet\b', 'PROFINET', new_content)

    # Revert PROFINET.dll if any
    new_content = new_content.replace('PROFINET.dll', 'Profinet.dll')

    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Updated {filepath}")

for root, _, files in os.walk('tia-portal-ai-extensions-main'):
    for file in files:
        if file.endswith('.md'):
            process_file(os.path.join(root, file))

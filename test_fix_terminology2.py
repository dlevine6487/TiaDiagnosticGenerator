import os
import re

def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    new_content = content

    # Replace WinCC Unified with WinCC Unified (if wrong capitalization)
    new_content = re.sub(r'(?i)\bWinCC Unified\b', 'WinCC Unified', new_content)

    if new_content != content:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(new_content)
        print(f"Updated {filepath}")

for root, _, files in os.walk('tia-portal-ai-extensions-main'):
    for file in files:
        if file.endswith('.md'):
            process_file(os.path.join(root, file))

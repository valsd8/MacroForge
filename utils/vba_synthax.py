# vba_converter.py
import string
import random

def generate_var_name(existing_names):
    """Return a random 5-letter name not in existing_names (existing_names is a set)."""
    while True:
        name = ''.join(random.choices(string.ascii_lowercase, k=5))
        if name not in existing_names:
            return name

def chunk_string(s, size=20):
    return [s[i:i+size] for i in range(0, len(s), size)]

def convertToVbaSynthax(input_string, chunk_size=20, declare=True):
    """
    Convert input_string into a multi-variable VBA assignment string.
    Returns the entire output as a single string (with newlines).
    """
    chunks = chunk_string(input_string, size=chunk_size)

    variables = {}
    var_names = []
    used_names = set()

    for i, chunk in enumerate(chunks):
        if i < 26:
            var_name = chr(ord('a') + i) 
        else:
            var_name = generate_var_name(used_names)
        variables[var_name] = chunk
        var_names.append(var_name)
        used_names.add(var_name)

    lines = []

    
    if declare:
        
        for var_name in var_names:
            lines.append(f'Dim {var_name} As String')
        lines.append("")  

   
    for var_name in var_names:
       
        chunk = variables[var_name].replace('"', '""')
        lines.append(f'{var_name} = "{chunk}"')
    lines.append("")  

    
    joined_vars = " & ".join(var_names)
    lines.append(f'cmd = {joined_vars}')

    return "\n".join(lines)

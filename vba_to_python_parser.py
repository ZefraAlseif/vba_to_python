import re
import argparse
from pathlib import Path

def vba_to_openpyxl(vba_code: str) -> str:
    python_lines = []
    with_stack = []

    python_lines.append("from openpyxl import load_workbook")
    python_lines.append("")
    python_lines.append("# Load workbook and select sheet")
    python_lines.append("wb = load_workbook(r'path_to_your_file.xlsx')  # TODO: change file path")
    python_lines.append("ws = wb.active  # or wb['SheetName']")
    python_lines.append("")

    for line in vba_code.splitlines():
        stripped = line.strip()

        # Skip empty lines
        if not stripped:
            python_lines.append("")
            continue

        # Comments
        if stripped.startswith("'"):
            python_lines.append("# " + stripped[1:].strip())
            continue

        # Sub / Function start
        m = re.match(r'^(Sub|Function)\s+(\w+)', stripped, re.IGNORECASE)
        if m:
            func_name = m.group(2)
            python_lines.append(f"def {func_name}():")
            continue

        # End Sub / End Function
        if re.match(r'^End\s+(Sub|Function)', stripped, re.IGNORECASE):
            python_lines.append("    pass  # End of function")
            continue

        # For loop
        m = re.match(r'^For\s+(\w+)\s*=\s*(\d+)\s+To\s+(\d+)', stripped, re.IGNORECASE)
        if m:
            var, start, end = m.groups()
            python_lines.append(f"    for {var} in range({start}, {int(end)+1}):")
            continue

        # Next
        if re.match(r'^Next', stripped, re.IGNORECASE):
            continue

        # If statements
        m = re.match(r'^If\s+(.+)\s+Then', stripped, re.IGNORECASE)
        if m:
            cond = m.group(1)
            python_lines.append(f"    if {cond}:  # TODO: adjust condition syntax")
            continue

        m = re.match(r'^ElseIf\s+(.+)\s+Then', stripped, re.IGNORECASE)
        if m:
            cond = m.group(1)
            python_lines.append(f"    elif {cond}:  # TODO: adjust condition syntax")
            continue

        if re.match(r'^Else$', stripped, re.IGNORECASE):
            python_lines.append("    else:")
            continue

        if re.match(r'^End\s+If', stripped, re.IGNORECASE):
            continue

        # With ... End With
        m = re.match(r'^With\s+(.+)', stripped, re.IGNORECASE)
        if m:
            obj = m.group(1)
            with_var = f"__with_obj_{len(with_stack)}"
            python_lines.append(f"    {with_var} = {obj}  # From VBA With statement")
            with_stack.append(with_var)
            continue

        if re.match(r'^End\s+With', stripped, re.IGNORECASE):
            if with_stack:
                with_stack.pop()
            continue

        # Select Case
        m = re.match(r'^Select\s+Case\s+(.+)', stripped, re.IGNORECASE)
        if m:
            var = m.group(1)
            python_lines.append(f"    match {var}:  # Python 3.10+ match/case")
            continue

        m = re.match(r'^Case\s+(.+)', stripped, re.IGNORECASE)
        if m:
            val = m.group(1)
            python_lines.append(f"        case {val}:")
            continue

        if re.match(r'^End\s+Select', stripped, re.IGNORECASE):
            continue

        # Do While / Loop Until
        m = re.match(r'^Do\s+While\s+(.+)', stripped, re.IGNORECASE)
        if m:
            cond = m.group(1)
            python_lines.append(f"    while {cond}:  # TODO: adjust condition syntax")
            continue

        m = re.match(r'^Loop\s+Until\s+(.+)', stripped, re.IGNORECASE)
        if m:
            cond = m.group(1)
            python_lines.append(f"    while not ({cond}):  # TODO: adjust condition syntax")
            continue

        if re.match(r'^Loop$', stripped, re.IGNORECASE):
            continue

        # Range references -> ws["A1"]
        line = re.sub(r'Range\("([^"]+)"\)\.Value', r'ws["\1"].value', line, flags=re.IGNORECASE)

        # Cells references -> ws.cell(row, col)
        line = re.sub(r'Cells\((\d+),\s*(\d+)\)\.Value', r'ws.cell(row=\1, column=\2).value', line, flags=re.IGNORECASE)

        # Replace '=' with proper spacing for Python (simple)
        if "=" in line and "==" not in line and "!=" not in line:
            line = line.replace("=", " = ")

        # Replace leading '.' inside With block
        if with_stack and stripped.startswith("."):
            line = line.replace(".", f"{with_stack[-1]}.", 1)

        # Indent
        line = "    " + line.strip()

        python_lines.append(line)

    python_lines.append("")
    python_lines.append("# Save changes")
    python_lines.append("wb.save(r'path_to_your_file.xlsx')  # TODO: change file path")

    return "\n".join(python_lines)


def main():
    parser = argparse.ArgumentParser(description="Convert VBA .bas macro to Python using openpyxl")
    parser.add_argument("input", help="Path to .bas file")
    parser.add_argument("-o", "--output", help="Path to save Python file", default=None)
    args = parser.parse_args()

    vba_code = Path(args.input).read_text(encoding="utf-8", errors="ignore")
    python_code = vba_to_openpyxl(vba_code)

    if args.output:
        Path(args.output).write_text(python_code, encoding="utf-8")
        print(f"âœ… Converted file saved to {args.output}")
    else:
        print(python_code)

if __name__ == "__main__":
    main()
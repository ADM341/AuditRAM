from auditram import AuditRAMHighlighter

file_path = "input.pdf"       # Upload path
search_text = "invoice"       # Text to highlight
output_file = "output.pdf"    # Final highlighted file

audit = AuditRAMHighlighter(file_path, search_text)
audit.run(output_file)

print("Done! Output created at:", output_file)

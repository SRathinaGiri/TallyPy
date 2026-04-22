import sys
content = open(r'C:\Users\acer\OneDrive\MyProjects\Python\TallyXML\tally_stock_vouchers.py').read()
escaped = content.replace('"', '""').replace('\n', '#(lf)')
with open('escaped_python.txt', 'w', encoding='utf-8') as f:
    f.write(escaped)

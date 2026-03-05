import os
import re

base_path = r"C:\Users\Armarinho\OneDrive\Documentos\Gestão&Controle\js"

# We replace: `R$ ${VAL.toFixed(2).replace('.', ',')}` with `${formatCurrencyBRL(VAL)}`
pattern1 = re.compile(r"R\$\s*\$\{\s*(.+?)\.toFixed\(2\)\.replace\('\.',\s*','\)\s*\}")
# We replace: 'R$ ' + VAL.toFixed(2).replace('.', ',') with formatCurrencyBRL(VAL)
pattern2 = re.compile(r"'R\$\s*'\s*\+\s*(.+?)\.toFixed\(2\)\.replace\('\.',\s*','\)")
# 'R$ ' + precoCusto.toFixed(...)
pattern3 = re.compile(r"'R\$\s*'\s*\+\s*([a-zA-Z0-parseFloat()\.||]+?)\.toFixed\(2\)\.replace\('\.',\s*','\)")


def process_file(filepath):
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    original = content
    content = pattern1.sub(r"${formatCurrencyBRL(\1)}", content)
    content = pattern2.sub(r"formatCurrencyBRL(\1)", content)

    if content != original:
        with open(filepath, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"Updated: {os.path.basename(filepath)}")

for filename in os.listdir(base_path):
    if filename.endswith(".js"):
        process_file(os.path.join(base_path, filename))

import docxcompose
import os
import shutil

src = os.path.join(os.path.dirname(docxcompose.__file__), 'templates')
dst = os.path.join(os.getcwd(), 'docxcompose_templates')

if os.path.exists(dst):
    shutil.rmtree(dst)

shutil.copytree(src, dst)
print(f"Copied templates to {dst}")

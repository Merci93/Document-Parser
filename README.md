# Document_Parser
Scripts to parse PDF and Word documents.
This repository contains the modules developed to parse PDF and Word documents, extracting Table of Contents, Images, texts, and tables. The extracted data are parsed and saved in separate folders `images`, `tables`, `table of contents` and `texts` in a parent directory `parsed_files` and a csv file generated with reports of extracted data from each document.

# Local Setup
### Windows
```
git clone https://github.com/Merci93/Document_Parser

cd Document_Parser

mkdir files_to_parse

python -m venv env

env/Scripts/activate

python.exe -m pip install --upgrade pip

pip install -r requirements.txt

python src/parse_all.py
```

### Linux / MacOS
```
git clone https://github.com/Merci93/Document_Parser

cd Document_Parser

mkdir files_to_parse

python -m venv env

source env/bin/activate

pip install --upgrade pip

pip install -r requirements.txt

python3 src/parse_all.py
```

>>> NOTE: The files to be parsed (Word or PDF) must be in the directory `files_to_parse`.

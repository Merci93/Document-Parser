# Document_Parser
Scripts to parse PDF and Word documents.
Documents are parsed and saved in separate folders `images`, `tables`, `table of contents` and `texts` in a parent directory `parsed_files`.

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

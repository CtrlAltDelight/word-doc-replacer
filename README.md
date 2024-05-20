# Word Doc Replacer

A simple tkinter GUI app that replaces mail merge codes in a word document with user input.

## Installation

To install the project, clone the repository:

```
git clone https://github.com/CtrlAltDelight/word-doc-replacer
cd word-doc-replacer
```

Then, create a conda environment from the `environment.yml` file and activate it:

```
conda env create -f environment.yml
conda activate word-replacer
```

Now, run the program:

```
python3 replacer.py
```

## Usage

Put mail merge codes in your word document for things you need to replace often in documents:

```
Dear {Name|Luke}

These codes are formatted like this: {Field|DefaultValue}
```

When you run the program and select a `.docx` file, it will scan for these merge codes and automatically populate each field with the default value. Change these values to whatever you wish, and then save a new document somewhere in your file system. Voila! You've just replaced mail merge codes in a word document.

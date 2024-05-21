# Word Doc Replacer

A simple tkinter GUI app that replaces mail merge codes in a word document with user input.

## Installation

To install the project, clone the repository and enter the project directory:

```bash
git clone https://github.com/CtrlAltDelight/word-doc-replacer
cd word-doc-replacer
```

To install dependencies, you can use conda or pip:

### Conda

Then, create a conda environment from the `environment.yml` file and activate it:

```bash
conda env create -f environment.yml
conda activate word-replacer
```

If you prefer not to use conda, you may use any other package manager to install the dependencies listed in `environment.yml`.

### Pip

```bash
pip install -r requirements.txt
```

### Other

The main two dependencies for this project are `tk` and `python-docx` if you want to install them individually.

## Running the program

To run the program:

```bash
python3 replacer.py
```

## Usage

Put mail merge codes in your word document for elements you need to replace frequently. These codes are formatted like this: `{Field|DefaultValue}`.

When you run the program and select a `.docx` file, it will scan for merge codes. The program will populate each field with the default value and prompt you to change them. Once you have made necessary changes, you can save the file as a new document.

Voila! You've just replaced merge codes in a word document!

## Example Use

Let's say you're a Jedi trying to recruit young force sensitives in your efforts against the dark side.

You could create a document like this:

```
Dear {Recipient|Luke}

We want to talk to you about {Subject|the Force}. Do you still live on {Planet|Tatooine}?

Sincerely,
{Sender|Uncle Ben}
```

This way, Uncle Ben can use this formatted document to make personalized messages to many young force sensitives and just switch out the key components of the message when necessary.

## TODO

This project is still in development. Here are some things I plan to add:

- [ ] Add support for more file types like .pages, .txt, .pdf, etc.
- [ ] Add support for more complex mail merge codes, such as loops and conditionals
- [ ] Make the GUI more user-friendly (add tooltips, better error messagest, etc.)
- [ ] Arrange the GUI elements in a more satisfying way
- [ ] Add a progress bar for long documents
- [ ] Add feature to preview document before saving

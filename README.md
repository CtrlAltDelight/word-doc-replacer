# Word Doc Replacer

A simple tkinter GUI app that replaces mail merge codes in a word document with user input.

## Installation

To install the project, clone the repository and enter the project directory:

```
git clone https://github.com/CtrlAltDelight/word-doc-replacer
cd word-doc-replacer
```

Then, create a conda environment from the `environment.yml` file and activate it:

```
conda env create -f environment.yml
conda activate word-replacer
```

If you would rather not use conda, you may use any other package manager to install the dependencies listed in `environment.yml`.

Otherwise, you can install the following dependencies manually (Check environment.yml for the most up-to-date list of dependencies):

```
tk
python-docx
```

Now, run the program:

```
python3 replacer.py
```

## Usage

Put mail merge codes in your word document for things you need to replace often in documents.

```
These codes are formatted like this: {Field|DefaultValue}
```

When you run the program and select a `.docx` file, it will scan for merge codes.

The program will populate each field with the default value and prompt you to change them. Once you have made necessary changes, you can save the file as a new document.

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

This way, Uncle Ben can use this formatted document to make personalized messages to many young force sensitives and just switch out the key components of the message.

## TODO

This project is still in development. Here are some things I plan to add:

- [ ] Add support for more file types like .pages, .txt, .pdf, etc.
- [ ] Add support for more complex mail merge codes like loops and conditionals
- [ ] Make the GUI more user-friendly (add tooltips, better error messagest, etc.)
- [ ] Arrange the GUI elements in a more satisfying way
- [ ] Add a progress bar for long documents
- [ ] Add feature to preview document before saving

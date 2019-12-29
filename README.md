# msoffice-tag-with-companyname
Look for company name in Microsoft Word Documents and tag file with name

A user at <https://forum.keyboardmaestro.com/t/apple-script-help/16492> asked:

> I'm looking for a simple Apple Script to use with Hazel that looks at the Company Name in the summary box of a Microsoft Word document and then tags the document with the name of the company. 
>
> I have 22,000 documents to sort from lots of different companies and it would help if I could initially tag each one with the name of the company it relates to. 

This is clearly an excellent case for automation.

[msoffice-tag-with-companyname.sh](https://github.com/tjluoma/msoffice-tag-with-companyname/blob/master/msoffice-tag-with-companyname.sh) seems to do just that. It does not use AppleScript. It uses `zsh` and [tag](https://github.com/jdberry/tag/) which can be installed using [brew](https://brew.sh) via `brew install tag` or using [MacPorts](https://www.macports.org) via `sudo port install tag`.

## Usage

The script is very straightforward, and _attempts_ to be at least minimally smart.

`msoffice-tag-with-companyname.sh FILENAME.docx`

will tag the file `FILENAME.docx` with the company name.

### What if there is no company name?

A message will be logged and nothing will be done.

### What if the file already has a tag with the company name? Or what if the file doesn’t exist? Or what if it does exist, but is not writable?

A message will be logged and nothing will be done.

### What if I try to run the command on a file that is not a Microsoft Office document?

If the file does not match a list of known filename extensions (which can be set in the script) then a message will be logged and nothing will be done.

### What if the `tag` command fails for some reason?

A message will be logged.

### Where is this log?

It will be stored at **$HOME/Library/Logs/msoffice-tag-with-companyname.log** which is the standard folder for saving logs in macOS.

## How do I use this with Hazel?

Download [msoffice-tag-with-companyname.sh](https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/msoffice-tag-with-companyname.sh) and save it somewhere such as **"$HOME/bin"**:

```
mkdir ~/bin/

cd ~/bin/

curl -sfLS "https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/msoffice-tag-with-companyname.sh" > msoffice-tag-with-companyname.sh

chmod 755 msoffice-tag-with-companyname.sh
```

Note that the `curl` command should be one long line.



```




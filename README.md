# msoffice-tag-with-companyname
Look for company name in Microsoft Word Documents and tag file with name

A user at <https://forum.keyboardmaestro.com/t/apple-script-help/16492> asked:

> I’m looking for a simple Apple Script to use with Hazel that looks at the Company Name in the summary box of a Microsoft Word document and then tags the document with the name of the company.
>
> I have 22,000 documents to sort from lots of different companies and it would help if I could initially tag each one with the name of the company it relates to.

This is clearly an excellent case for automation.

[msoffice-tag-with-companyname.sh](https://github.com/tjluoma/msoffice-tag-with-companyname/blob/master/msoffice-tag-with-companyname.sh) seems to do just that. It does not use AppleScript. It uses `zsh` and [tag](https://github.com/jdberry/tag/) which can be installed using [brew](https://brew.sh) via `brew install tag` or using [MacPorts](https://www.macports.org) via `sudo port install tag`.

## Usage

The script is very straightforward, and _attempts_ to be at least minimally smart.

`msoffice-tag-with-companyname.sh FILENAME.docx`

will tag the file `FILENAME.docx` with the company name.

`msoffice-tag-with-companyname.sh *.docx`

will tag all of the files with a `.docx` extension.

`msoffice-tag-with-companyname.sh *`

will attempt to tag all of the files in the current directory.

### What if there is no company name?

A message will be logged and nothing will be done.

### What if the file already has a tag with the company name? Or what if the file doesn’t exist? Or what if it does exist, but is not writable?

A message will be logged and nothing will be done.

### What if I try to run the command on a file that is not a Microsoft Office document?

If the file does not match a list of known filename extensions (which can be set in the script) then a message will be logged and nothing will be done.

By default the script will only try to tag the file if it has one of these file extensions:

* doc
* docx
* xls
* ppt

You can easily edit that list by changing this line in the script:

```
      'doc'|'docx'|'xls'|'ppt')
```

### What if the `tag` command fails for some reason?

A message will be logged.

### Where is this log?

It will be stored at **$HOME/Library/Logs/msoffice-tag-with-companyname.log** which is the standard folder for saving logs in macOS.

## How do I use this with Hazel?

## Step One: Download the script and make it executable

Download [msoffice-tag-with-companyname.sh](https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/msoffice-tag-with-companyname.sh) and save it somewhere such as **"$HOME/bin"**:

```
mkdir ~/bin/

cd ~/bin/

curl -sfLS 'https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/msoffice-tag-with-companyname.sh' > msoffice-tag-with-companyname.sh

chmod 755 msoffice-tag-with-companyname.sh
```

Note that the `curl` command should be one long line.

## Step Two: Tell Hazel to use the script:

Once you have your Hazel rule configured, choose “Run Shell Script” as the action. Click on “Embedded Script” and choose “Other” as shown here:

![Hazel screenshot before](https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/Hazel-MSOffice.png)

Then choose the `msoffice-tag-with-companyname.sh` script from **~/bin/**

It should look something like this:

![Hazel screenshot after](https://raw.githubusercontent.com/tjluoma/msoffice-tag-with-companyname/master/Hazel-MSOffice-2.png)


## Step Three

[There’s no step three](https://www.youtube.com/watch?v=6uXJlX50Lj8).


## Caveat

This script has only has limited testing, so I cannot guarantee that it will work on all Microsoft Word files, or Microsoft Office files.

If you find a file that _has_ a company name but _does not_ work with this script, enter this line in Terminal:

`mdls "/path/to/your/filename.ext"`

For example, if the file is on your Desktop and is called “My Report.doc” then you would do:

`mdls "$HOME/Desktop/My Report.doc"`

and copy/paste the output into a message here so I can see it. If the company name does not appear in the output of `mdls` then we’ll have to look for another way to find it.

★

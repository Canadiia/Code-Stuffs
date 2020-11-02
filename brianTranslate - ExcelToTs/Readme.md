This file will take translations and their comments from an Excel file and put them into a pre-setup .ts file
heeds the input file as a command line argument
should also have a command line argument that:
1. takes the translations and translates them to english
2. put them with their associated source text
3. save them to an excel file with the source next to the translation

Use **sys** and **getopt**

 - The Python **sys** module provides access to any command-line arguments via the `sys.argv`.
 - This serves two purposes:
     - `sys.argv` is the list of command-line arguments.
     - `len(sys.argv)` is the number of command-line arguments.
 - _Here `sys.argv[0]` is the program ie. script name._

getopt.getopt method

This method parses command line options and parameter list. Following is simple syntax for this method:

`getopt.getopt(args, options, [long_options])`

Here is the detail of the parameters:
 - `args` − This is the argument list to be parsed.
 - `options` − This is the string of option letters that the script wants to recognize, with options that require an argument should be followed by a colon (:).
 - `long_options` − This is optional parameter and if specified, must be a list of strings with the names of the long options, which should be supported. Long options, which require an argument should be followed by an equal sign ('='). To accept only long options, options should be an empty string.

This method returns value consisting of two elements: the first is a list of (option, value) pairs. The second is the list of program arguments left after the option list was stripped.
Each option-and-value pair returned has the option as its first element, prefixed with a hyphen for short options (e.g., '-x') or two hyphens for long options (e.g., '--long-option').

Exception `getopt.GetoptError`

This is raised when an unrecognized option is found in the argument list or when an option requiring an argument is given none.
The argument to the exception is a string indicating the cause of the error. The attributes `msg` and `opt` give the error message and related option.